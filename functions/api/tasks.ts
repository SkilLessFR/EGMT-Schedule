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

export interface TaskItem {
  id: string;
  title: string;
  times: string[];
  scheduleType: 'once' | 'daily' | 'weekly';
  daysOfWeek?: number[];
  dateCreated?: string;
}

export interface StoredTasksPayload {
  tasks: TaskItem[];
  updatedAt: string;
  updatedBy?: string;
}

const DEFAULT_SEED_TASKS: TaskItem[] = [
  {
    id: "59522edd-39f3-4d39-87cc-5c0e972e1a02",
    title: "S43 Change Dice",
    times: ["05:45", "10:45", "16:45", "22:45"],
    scheduleType: "daily",
  },
  {
    id: "1af0132c-ead7-4505-934b-4407caeb1650",
    title: "A10 Change Dice",
    times: ["00:45", "12:45"],
    scheduleType: "daily",
  },
  {
    id: "95b41a63-9d70-47d7-93c4-ab490bd09c07",
    title: "VBB Bile change",
    times: ["07:45"],
    scheduleType: "daily",
  },
  {
    id: "3265b03b-4374-4801-b9fe-8d718df98c54",
    title: "S43 Change Dice in 1 hour",
    times: ["05:00", "10:00", "16:00", "22:00"],
    scheduleType: "daily",
    dateCreated: "2026-07-31",
  },
  {
    id: "a7035514-b725-464c-9ec3-f34907fde8af",
    title: "A10 Change Dice in 1 hour",
    times: ["00:00", "12:00"],
    scheduleType: "daily",
    dateCreated: "2026-07-31",
  },
  {
    id: "c7763170-4a50-499d-834e-f56afca7c487",
    title: "VBB Bile Change in 1 hour",
    times: ["07:00"],
    scheduleType: "daily",
    dateCreated: "2026-07-31",
  },
];

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

// GET /api/tasks
export async function onRequestGet(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(
      JSON.stringify({
        success: true,
        tasks: DEFAULT_SEED_TASKS,
        updatedAt: new Date().toISOString(),
        source: 'local-fallback',
      }),
      { status: 200, headers: jsonHeaders }
    );
  }

  try {
    const rawData = await kv.get('site_tasks', 'text');
    if (!rawData) {
      const initialPayload: StoredTasksPayload = {
        tasks: DEFAULT_SEED_TASKS,
        updatedAt: new Date().toISOString(),
        updatedBy: 'initial_seed',
      };
      await kv.put('site_tasks', JSON.stringify(initialPayload));

      return new Response(
        JSON.stringify({
          success: true,
          tasks: initialPayload.tasks,
          updatedAt: initialPayload.updatedAt,
          source: 'seeded_kv',
        }),
        { status: 200, headers: jsonHeaders }
      );
    }

    const parsed = JSON.parse(rawData);
    if (Array.isArray(parsed)) {
      return new Response(
        JSON.stringify({
          success: true,
          tasks: parsed,
          updatedAt: new Date().toISOString(),
          source: 'kv',
        }),
        { status: 200, headers: jsonHeaders }
      );
    }

    return new Response(
      JSON.stringify({
        success: true,
        tasks: parsed.tasks ?? [],
        updatedAt: parsed.updatedAt ?? new Date().toISOString(),
        updatedBy: parsed.updatedBy,
        source: 'kv',
      }),
      { status: 200, headers: jsonHeaders }
    );
  } catch (error) {
    return new Response(
      JSON.stringify({
        error: 'Failed to read tasks from KV',
        details: String(error),
        tasks: DEFAULT_SEED_TASKS,
      }),
      { status: 500, headers: jsonHeaders }
    );
  }
}

// POST /api/tasks
export async function onRequestPost(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(
      JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound in Cloudflare Pages' }),
      { status: 503, headers: jsonHeaders }
    );
  }

  try {
    const body = await context.request.json();
    let tasksList: TaskItem[] = [];
    let updatedBy: string | undefined;

    if (Array.isArray(body)) {
      tasksList = body;
    } else if (body && Array.isArray(body.tasks)) {
      tasksList = body.tasks;
      updatedBy = typeof body.updatedBy === 'string' ? body.updatedBy.trim() : undefined;
    } else {
      return new Response(
        JSON.stringify({ error: 'Invalid payload. "tasks" must be an array of Task objects.' }),
        { status: 400, headers: jsonHeaders }
      );
    }

    const sanitizedTasks: TaskItem[] = tasksList.map((t) => ({
      id: t.id || crypto.randomUUID(),
      title: (t.title || 'Untitled Task').trim(),
      times: Array.isArray(t.times) ? Array.from(new Set(t.times)).sort() : ['00:00'],
      scheduleType: t.scheduleType || 'daily',
      daysOfWeek: Array.isArray(t.daysOfWeek) ? [...t.daysOfWeek].sort() : undefined,
      dateCreated: t.dateCreated || new Date().toISOString().slice(0, 10),
    }));

    const payload: StoredTasksPayload = {
      tasks: sanitizedTasks,
      updatedAt: new Date().toISOString(),
      updatedBy,
    };

    await kv.put('site_tasks', JSON.stringify(payload));

    return new Response(
      JSON.stringify({
        success: true,
        count: sanitizedTasks.length,
        updatedAt: payload.updatedAt,
        updatedBy: payload.updatedBy,
      }),
      { status: 200, headers: jsonHeaders }
    );
  } catch (error) {
    return new Response(
      JSON.stringify({ error: 'Failed to update tasks in KV', details: String(error) }),
      { status: 500, headers: jsonHeaders }
    );
  }
}
