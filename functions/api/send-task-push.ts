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
}

interface SendTaskPushPayload {
  taskId?: string;
  taskTitle: string;
  time?: string;
  targetShift?: 'ALL_ACTIVE' | 'M' | 'A' | 'N' | 'MID';
  body?: string;
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

  const d = new Date(`${isoDate}T12:00:00Z`);
  const dayOfWeek = d.getUTCDay(); // 0 = Sunday, 1 = Monday, ...
  d.setUTCDate(d.getUTCDate() - 1);
  const prevIsoDate = d.toISOString().slice(0, 10);

  return { isoDate, hhmm, prevIsoDate, dayOfWeek };
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
  task: { id: string; title: string; targetShift?: string },
  time: string,
  isoDate: string,
  prevIsoDate: string,
  customBody?: string
) {
  // Deduplicate within the minute to avoid duplicate pushes if multiple devices or crons trigger
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

  // Mark dedup immediately
  await kv.put(dedupKey, new Date().toISOString());

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

  const targetShift = task.targetShift || 'ALL_ACTIVE';
  const eligibleEmployees: { name: string; shift: string }[] = [];

  for (const emp of roster.employees) {
    const shiftStatus = checkIsOnShift(roster.rows, emp, time, isoDate, prevIsoDate);
    if (!shiftStatus.onShift) continue;
    if (targetShift !== 'ALL_ACTIVE' && shiftStatus.shift !== targetShift) continue;
    eligibleEmployees.push({ name: emp, shift: shiftStatus.shift || 'ON_SHIFT' });
  }

  const delivered: string[] = [];
  const dispatchErrors: { name: string; error: string }[] = [];

  for (const { name, shift } of eligibleEmployees) {
    const rawSub = await kv.get(`sub:${name}`, 'text');
    if (!rawSub) continue;

    const pushRes = await dispatchWebPush(rawSub, {
      title: `⏰ Task Alert: ${task.title}`,
      body: customBody || `Shift ${shift} task reminder for ${time}`,
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
    const { isoDate, hhmm, prevIsoDate, dayOfWeek } = getBucharestTime();

    // 1. Fetch active tasks from KV
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

    // 2. Filter tasks scheduled for current minute
    const matchingTasks = tasks.filter((t) => {
      if (!t.times || !t.times.includes(hhmm)) return false;
      if (t.scheduleType === 'daily') return true;
      if (t.scheduleType === 'once' && t.dateCreated === isoDate) return true;
      if (t.scheduleType === 'weekly' && t.daysOfWeek?.includes(dayOfWeek)) return true;
      return false;
    });

    if (matchingTasks.length === 0) {
      return new Response(
        JSON.stringify({
          success: true,
          evaluatedTime: hhmm,
          bucharestDate: isoDate,
          dayOfWeek,
          matchedTasks: 0,
          message: `No tasks scheduled for ${hhmm}`,
        }),
        { status: 200, headers: jsonHeaders }
      );
    }

    // 3. Dispatch Web Push for each matching task
    const results = [];
    for (const task of matchingTasks) {
      const res = await dispatchTaskToShift(kv, task, hhmm, isoDate, prevIsoDate);
      results.push(res);
    }

    return new Response(
      JSON.stringify({
        success: true,
        evaluatedTime: hhmm,
        bucharestDate: isoDate,
        matchedTasks: matchingTasks.length,
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
      },
      checkHHMM,
      isoDate,
      prevIsoDate,
      payload.body
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
