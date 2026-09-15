interface KVNamespaceLike {
  get(key: string, type: 'text'): Promise<string | null>;
  put(key: string, value: string): Promise<void>;
}

interface Env {
  EGMT_ROSTER?: KVNamespaceLike;
  EGMT_roster?: KVNamespaceLike;
}

interface PagesContext {
  request: Request;
  env: Env;
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

export async function onRequestGet(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound' }), {
      status: 503,
      headers: jsonHeaders,
    });
  }

  const url = new URL(context.request.url);
  const employee = url.searchParams.get('employee');

  if (!employee) {
    return new Response(JSON.stringify({ error: 'Employee query parameter is required' }), {
      status: 400,
      headers: jsonHeaders,
    });
  }

  try {
    const rawSub = await kv.get(`sub:${employee}`, 'text');
    return new Response(JSON.stringify({
      employee,
      subscribed: Boolean(rawSub),
    }), {
      status: 200,
      headers: jsonHeaders,
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: 'Failed to read subscription from KV', details: String(err) }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}

export async function onRequestPost(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound' }), {
      status: 503,
      headers: jsonHeaders,
    });
  }

  try {
    const payload = await context.request.json() as { employee?: string; subscription?: unknown };
    const employee = payload.employee?.trim();
    const subscription = payload.subscription;

    if (!employee || !subscription) {
      return new Response(JSON.stringify({ error: 'Employee name and push subscription object are required.' }), {
        status: 400,
        headers: jsonHeaders,
      });
    }

    // Read existing subscriptions for this employee to support multiple devices (e.g. Phone + Laptop)
    const rawExisting = await kv.get(`sub:${employee}`, 'text');
    let existingSubs: Array<{ endpoint?: string; keys?: unknown }> = [];
    if (rawExisting) {
      try {
        const parsed = JSON.parse(rawExisting);
        if (Array.isArray(parsed.subscriptions)) {
          existingSubs = parsed.subscriptions;
        } else if (parsed.subscription) {
          existingSubs = [parsed.subscription];
        } else if (parsed.endpoint) {
          existingSubs = [parsed];
        }
      } catch {
        // Fallback
      }
    }

    const newSubObj = subscription as { endpoint?: string; keys?: unknown };
    const newEndpoint = newSubObj?.endpoint;

    // Filter out duplicate or matching endpoint and keep up to 5 devices per employee
    const updatedSubs = [
      newSubObj,
      ...existingSubs.filter((s) => s?.endpoint && s.endpoint !== newEndpoint),
    ].slice(0, 5);

    // Store the employee's subscriptions in KV
    await kv.put(`sub:${employee}`, JSON.stringify({
      subscription, // Backwards-compatible single subscription
      subscriptions: updatedSubs, // Multi-device subscriptions array
      updatedAt: new Date().toISOString(),
    }));

    // Update active subscriber registry
    try {
      const rawList = await kv.get('subscribers_index', 'text');
      const list = rawList ? (JSON.parse(rawList) as string[]) : [];
      if (!list.includes(employee)) {
        list.push(employee);
        await kv.put('subscribers_index', JSON.stringify(list));
      }
    } catch {
      // Non-critical index update
    }

    return new Response(JSON.stringify({
      success: true,
      message: `Push subscription linked to ${employee} in Cloudflare KV.`,
    }), {
      status: 200,
      headers: jsonHeaders,
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: 'Failed to save subscription in KV', details: String(error) }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}
