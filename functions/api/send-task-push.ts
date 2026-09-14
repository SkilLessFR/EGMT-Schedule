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

function getBucharestTime(date: Date = new Date()): { isoDate: string; hhmm: string; prevIsoDate: string } {
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
  d.setUTCDate(d.getUTCDate() - 1);
  const prevIsoDate = d.toISOString().slice(0, 10);

  return { isoDate, hhmm, prevIsoDate };
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

    // 1. Fetch active roster from KV to determine shift presence
    const rawRoster = await kv.get('active_roster', 'text');
    if (!rawRoster) {
      return new Response(
        JSON.stringify({
          error: 'No active roster found in KV. Roster is needed to filter by active shift.',
        }),
        { status: 404, headers: jsonHeaders }
      );
    }

    const roster = JSON.parse(rawRoster) as {
      employees: string[];
      rows: Record<string, Record<string, string>>;
    };

    const eligibleEmployees: { name: string; shift: string }[] = [];
    const skippedOffShift: { name: string; reason: string }[] = [];

    for (const emp of roster.employees) {
      const shiftStatus = checkIsOnShift(roster.rows, emp, checkHHMM, isoDate, prevIsoDate);
      if (!shiftStatus.onShift) {
        skippedOffShift.push({ name: emp, reason: 'Off shift or not working at this hour' });
        continue;
      }

      if (targetShift !== 'ALL_ACTIVE' && shiftStatus.shift !== targetShift) {
        skippedOffShift.push({
          name: emp,
          reason: `On shift ${shiftStatus.shift}, but task is designated for ${targetShift}`,
        });
        continue;
      }

      eligibleEmployees.push({ name: emp, shift: shiftStatus.shift || 'ON_SHIFT' });
    }

    // 2. Dispatch push notifications to on-shift colleagues with active subscriptions
    const delivered: string[] = [];
    const noSubscription: string[] = [];
    const dispatchErrors: { name: string; error: string }[] = [];

    for (const { name, shift } of eligibleEmployees) {
      const rawSub = await kv.get(`sub:${name}`, 'text');
      if (!rawSub) {
        noSubscription.push(name);
        continue;
      }

      const pushRes = await dispatchWebPush(rawSub, {
        title: `⏰ Task Alert: ${taskTitle}`,
        body: payload.body || `Shift ${shift} task reminder for ${checkHHMM}`,
        icon: '/icon.svg',
        badge: '/icon.svg',
        tag: `task-${payload.taskId || 'alert'}-${checkHHMM}`,
        data: { url: '/', taskId: payload.taskId, shift },
      });

      if (pushRes.success) {
        delivered.push(name);
      } else {
        dispatchErrors.push({ name, error: pushRes.error || 'Failed to dispatch' });
      }
    }

    return new Response(
      JSON.stringify({
        success: true,
        taskTitle,
        targetShift,
        evaluatedTime: checkHHMM,
        bucharestDate: isoDate,
        eligibleCount: eligibleEmployees.length,
        delivered,
        noSubscription,
        dispatchErrors,
        skippedOffShiftCount: skippedOffShift.length,
        skippedOffShift,
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
