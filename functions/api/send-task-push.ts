// functions/api/send-task-push.ts
import { dispatchWebPush } from './_webPush';

interface KVNamespaceLike {
  get(key: string, type: 'text'): Promise<string | null>;
  put(key: string, value: string): Promise<void>;
  delete(key: string): Promise<void>;
}

interface Env {
  EGMT_ROSTER?: KVNamespaceLike;
  EGMT_roster?: KVNamespaceLike;
}

interface PagesContext {
  request: Request;
  env: Env;
}

interface TaskItem {
  id: string;
  title: string;
  times: string[];
  scheduleType: 'once' | 'daily' | 'weekly';
  daysOfWeek?: number[];
  dateCreated?: string;
  targetShift?: 'ALL_ACTIVE' | 'M' | 'A' | 'N' | 'MID';
  notifyOnlyAlex?: boolean;
}

interface SendTaskPushPayload {
  taskId?: string;
  taskTitle: string;
  time?: string;
  targetShift?: 'ALL_ACTIVE' | 'M' | 'A' | 'N' | 'MID';
  body?: string;
  notifyOnlyAlex?: boolean;
  isManualTest?: boolean;
}

const jsonHeaders = {
  'Content-Type': 'application/json',
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export async function onRequestOptions() {
  return new Response(null, {
    status: 204,
    headers: jsonHeaders,
  });
}

function getBucharestTime(date: Date = new Date()): {
  isoDate: string;
  hhmm: string;
  prevHhmm: string;
  prevIsoDate: string;
  dayOfWeek: number;
} {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Europe/Bucharest',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  });

  const parts = formatter.formatToParts(date);
  const partMap: Record<string, string> = {};
  for (const p of parts) {
    partMap[p.type] = p.value;
  }
  const isoDate = `${partMap.year}-${partMap.month}-${partMap.day}`;
  const hhmm = `${partMap.hour}:${partMap.minute}`;

  // 1 minute prior for jitter tolerance
  const prevDateObj = new Date(date.getTime() - 60000);
  const prevParts = formatter.formatToParts(prevDateObj);
  const prevPartMap: Record<string, string> = {};
  for (const p of prevParts) {
    prevPartMap[p.type] = p.value;
  }
  const prevHhmm = `${prevPartMap.hour}:${prevPartMap.minute}`;

  const d = new Date(`${isoDate}T12:00:00Z`);
  const dayOfWeek = d.getUTCDay(); // 0 = Sunday, 1 = Monday, ...
  d.setUTCDate(d.getUTCDate() - 1);
  const prevIsoDate = d.toISOString().slice(0, 10);

  return { isoDate, hhmm, prevHhmm, prevIsoDate, dayOfWeek };
}


function checkIsOnShift(
  rows: Record<string, Record<string, string>>,
  employee: string,
  hhmm: string,
  isoDate: string,
  prevIsoDate: string
): { onShift: boolean; shift?: string } {
  const OFF_ALIASES = new Set(['OFF', 'ABS', 'AD', 'H8']);

  // Early morning: 00:00 - 05:59 (yesterday's Night shift)
  if (hhmm < '06:00') {
    const raw = (rows[employee]?.[prevIsoDate] ?? 'OFF').trim().toUpperCase();
    if (!OFF_ALIASES.has(raw) && raw === 'N') {
      return { onShift: true, shift: 'N' };
    }
    return { onShift: false };
  }

  // Morning: 06:00 - 13:59
  if (hhmm >= '06:00' && hhmm < '14:00') {
    const raw = (rows[employee]?.[isoDate] ?? 'OFF').trim().toUpperCase();
    if (raw === 'M') return { onShift: true, shift: 'M' };
    if (raw === 'MID' && hhmm >= '09:00') return { onShift: true, shift: 'MID' };
    return { onShift: false };
  }

  // Afternoon: 14:00 - 21:59
  if (hhmm >= '14:00' && hhmm < '22:00') {
    const raw = (rows[employee]?.[isoDate] ?? 'OFF').trim().toUpperCase();
    if (raw === 'A') return { onShift: true, shift: 'A' };
    if (raw === 'MID' && hhmm < '17:00') return { onShift: true, shift: 'MID' };
    return { onShift: false };
  }

  // Night: 22:00 - 23:59
  if (hhmm >= '22:00') {
    const raw = (rows[employee]?.[isoDate] ?? 'OFF').trim().toUpperCase();
    if (raw === 'N') return { onShift: true, shift: 'N' };
    return { onShift: false };
  }

  return { onShift: false };
}

async function dispatchTaskToShift(
  kv: KVNamespaceLike,
  task: { id: string; title: string; targetShift?: string; notifyOnlyAlex?: boolean },
  time: string,
  isoDate: string,
  prevIsoDate: string,
  customBody?: string,
  isManualTest?: boolean
) {
  const isOnlyAlex = Boolean(task.notifyOnlyAlex) || String(task.notifyOnlyAlex) === 'true';

  // Deduplicate within the minute to avoid duplicate pushes between cron runner and client triggers
  if (!isManualTest) {
    const dedupKey = `dedup:task:${task.id}:${isoDate}:${time}`;
    const alreadySent = await kv.get(dedupKey, 'text');
    if (alreadySent) {
      return {
        taskId: task.id,
        title: task.title,
        skipped: true,
        reason: 'Already dispatched for this minute',
      };
    }
    await kv.put(dedupKey, new Date().toISOString());
  }

  const targetShift = task.targetShift || 'ALL_ACTIVE';
  const eligibleEmployees: { name: string; shift: string }[] = [];

  if (isOnlyAlex) {
    // Test mode: Send alert exclusively to Alex
    // Check both aliases in KV, but pick the first active one so we don't send duplicates
    const alexAliases = ['Stoian Alexandru-Gabriel', 'Alexandru Stoian'];
    let foundAlias: string | null = null;
    for (const name of alexAliases) {
      const sub = await kv.get(`sub:${name}`, 'text');
      if (sub) {
        eligibleEmployees.push({ name, shift: 'TEST' });
        foundAlias = name;
        break; // Stop at first valid alias to prevent duplicate notifications to Alex
      }
    }
    if (!foundAlias) {
      eligibleEmployees.push({ name: alexAliases[0], shift: 'TEST' });
    }
  } else {
    const rawRoster = await kv.get('active_roster', 'text');
    if (!rawRoster) {
      return {
        taskId: task.id,
        title: task.title,
        skipped: true,
        reason: 'No active roster found in KV',
      };
    }

    const roster = JSON.parse(rawRoster) as {
      employees: string[];
      rows: Record<string, Record<string, string>>;
    };

    for (const emp of roster.employees) {
      const shiftStatus = checkIsOnShift(roster.rows, emp, time, isoDate, prevIsoDate);
      if (!shiftStatus.onShift) continue;
      if (targetShift !== 'ALL_ACTIVE' && shiftStatus.shift !== targetShift) continue;
      eligibleEmployees.push({ name: emp, shift: shiftStatus.shift || 'ON_SHIFT' });
    }
  }

  const delivered: string[] = [];
  const dispatchErrors: { name: string; error: string }[] = [];

  for (const { name, shift } of eligibleEmployees) {
    const rawSub = await kv.get(`sub:${name}`, 'text');
    if (!rawSub) continue;

    const pushRes = await dispatchWebPush(rawSub, {
      title: `⏰ Task Alert: ${task.title}`,
      body: customBody || (isOnlyAlex
        ? `Task reminder for ${time}`
        : `Shift ${shift} task reminder for ${time}`),
      icon: '/icon.svg',
      badge: '/icon.svg',
      tag: `task-${task.id}-${time}`,
      data: { url: '/', taskId: task.id, shift },
    });

    if (pushRes.success) {
      delivered.push(name);
    } else {
      dispatchErrors.push({ name, error: pushRes.error || 'Failed to dispatch' });
    }
  }

  return {
    taskId: task.id,
    title: task.title,
    skipped: false,
    eligibleCount: eligibleEmployees.length,
    delivered,
    dispatchErrors,
  };
}

// GET /api/send-task-push (Automated Cron Runner)
// Evaluates all active tasks stored in KV against current Bucharest time and dispatches pushes
export async function onRequestGet(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(
      JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound' }),
      { status: 503, headers: jsonHeaders }
    );
  }

  try {
    const { isoDate, hhmm, prevHhmm, prevIsoDate, dayOfWeek } = getBucharestTime();
    const url = new URL(context.request.url);
    const showStatus = url.searchParams.get('status') === 'true';

    // 1. If status requested, return health and last ping info without dispatching
    if (showStatus) {
      const rawPing = await kv.get('last_cron_ping', 'text');
      let lastPing: { timestamp: string; bucharestTime: string; userAgent?: string } | null = null;
      let secondsAgo: number | null = null;
      if (rawPing) {
        try {
          lastPing = JSON.parse(rawPing);
          if (lastPing?.timestamp) {
            secondsAgo = Math.round((Date.now() - new Date(lastPing.timestamp).getTime()) / 1000);
          }
        } catch {
          // Ignore JSON parse error on legacy ping values
        }
      }
      return new Response(
        JSON.stringify({
          success: true,
          currentBucharestTime: hhmm,
          lastPing,
          secondsAgo,
          isHealthy: secondsAgo !== null && secondsAgo <= 180,
        }),
        { status: 200, headers: jsonHeaders }
      );
    }

    // 2. Record this ping into KV for audit and health monitoring
    const userAgent = context.request.headers.get('user-agent') || 'unknown';
    const nowIso = new Date().toISOString();
    await kv.put(
      'last_cron_ping',
      JSON.stringify({
        timestamp: nowIso,
        bucharestTime: hhmm,
        userAgent,
      })
    );

    // 3. Fetch active tasks from KV
    const rawTasks = await kv.get('site_tasks', 'text');
    let tasks: TaskItem[] = [];
    if (rawTasks) {
      const parsed = JSON.parse(rawTasks);
      tasks = Array.isArray(parsed) ? parsed : parsed.tasks || [];
    }

    if (tasks.length === 0) {
      return new Response(
        JSON.stringify({
          success: true,
          evaluatedTime: hhmm,
          bucharestDate: isoDate,
          message: 'No tasks configured in KV',
          matchedTasks: 0,
        }),
        { status: 200, headers: jsonHeaders }
      );
    }

    // 4. Filter tasks scheduled for current minute OR previous minute (jitter tolerance)
    const matchingTasksWithTimes: { task: TaskItem; time: string }[] = [];
    for (const t of tasks) {
      const isValidDate =
        t.scheduleType === 'daily' ||
        (t.scheduleType === 'once' && t.dateCreated === isoDate) ||
        (t.scheduleType === 'weekly' && t.daysOfWeek?.includes(dayOfWeek));
      if (!isValidDate || !t.times) continue;

      for (const checkTime of [prevHhmm, hhmm]) {
        if (t.times.includes(checkTime)) {
          // Avoid duplicate entry if task has duplicate time definition
          if (!matchingTasksWithTimes.some(m => m.task.id === t.id && m.time === checkTime)) {
            matchingTasksWithTimes.push({ task: t, time: checkTime });
          }
        }
      }
    }

    if (matchingTasksWithTimes.length === 0) {
      return new Response(
        JSON.stringify({
          success: true,
          evaluatedTime: hhmm,
          jitterCheckedTime: prevHhmm,
          bucharestDate: isoDate,
          dayOfWeek,
          matchedTasks: 0,
          message: `No tasks scheduled for ${hhmm} or ${prevHhmm}`,
        }),
        { status: 200, headers: jsonHeaders }
      );
    }

    // 5. Dispatch Web Push for each matching task (dedupKey prevents duplicate deliveries)
    const results = [];
    for (const item of matchingTasksWithTimes) {
      const res = await dispatchTaskToShift(kv, item.task, item.time, isoDate, prevIsoDate);
      results.push(res);
    }

    return new Response(
      JSON.stringify({
        success: true,
        evaluatedTime: hhmm,
        bucharestDate: isoDate,
        matchedTasks: matchingTasksWithTimes.length,
        results,
      }),
      { status: 200, headers: jsonHeaders }
    );
  } catch (err) {
    return new Response(
      JSON.stringify({ error: 'Failed to execute cron task push', details: String(err) }),
      { status: 500, headers: jsonHeaders }
    );
  }
}

// POST /api/send-task-push (Client or Webhook Trigger)
export async function onRequestPost(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(
      JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound' }),
      { status: 503, headers: jsonHeaders }
    );
  }

  try {
    const payload = (await context.request.json()) as Partial<SendTaskPushPayload>;
    const taskTitle = (payload.taskTitle || 'Scheduled Task').trim();
    const targetShift = payload.targetShift || 'ALL_ACTIVE';
    const time = payload.time || '';

    const { isoDate, hhmm, prevIsoDate } = getBucharestTime();
    const checkHHMM = time || hhmm;

    const res = await dispatchTaskToShift(
      kv,
      {
        id: payload.taskId || 'adhoc',
        title: taskTitle,
        targetShift,
        notifyOnlyAlex: payload.notifyOnlyAlex === true || String(payload.notifyOnlyAlex) === 'true',
      },
      checkHHMM,
      isoDate,
      prevIsoDate,
      payload.body,
      Boolean(payload.isManualTest)
    );

    return new Response(
      JSON.stringify({
        success: true,
        taskTitle,
        targetShift,
        evaluatedTime: checkHHMM,
        bucharestDate: isoDate,
        ...res,
      }),
      { status: 200, headers: jsonHeaders }
    );
  } catch (err) {
    return new Response(
      JSON.stringify({ error: 'Failed to process task push', details: String(err) }),
      { status: 500, headers: jsonHeaders }
    );
  }
}
