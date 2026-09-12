interface KVNamespaceLike {
  get(key: string, type: 'text'): Promise<string | null>;
  put(key: string, value: string): Promise<void>;
}

interface Env {
  EGMT_ROSTER?: KVNamespaceLike;
  EGMT_roster?: KVNamespaceLike;
  ADMIN_PIN?: string;
}

interface PagesContext {
  request: Request;
  env: Env;
}

const jsonHeaders = {
  'Content-Type': 'application/json',
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, X-Admin-Pin',
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
    return new Response(JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound in Cloudflare Pages' }), {
      status: 503,
      headers: jsonHeaders,
    });
  }

  try {
    const rawRoster = await kv.get('active_roster', 'text');
    if (!rawRoster) {
      return new Response(JSON.stringify({ error: 'No active roster uploaded yet in KV' }), {
        status: 404,
        headers: jsonHeaders,
      });
    }

    return new Response(rawRoster, {
      status: 200,
      headers: {
        ...jsonHeaders,
        'Cache-Control': 'public, max-age=60, s-maxage=120',
      },
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: 'Failed to read from KV', details: String(error) }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}

export async function onRequestPost(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound in Cloudflare Pages' }), {
      status: 503,
      headers: jsonHeaders,
    });
  }

  // Verify PIN if configured on Cloudflare
  const expectedPin = context.env.ADMIN_PIN?.trim();
  if (expectedPin) {
    const clientPin = context.request.headers.get('X-Admin-Pin')?.trim();
    if (clientPin !== expectedPin) {
      return new Response(JSON.stringify({ error: 'Invalid or missing Admin PIN.' }), {
        status: 401,
        headers: jsonHeaders,
      });
    }
  }

  try {
    const payload = await context.request.json();
    if (!payload || !Array.isArray(payload.employees) || !Array.isArray(payload.dateColumns) || !payload.rows) {
      return new Response(JSON.stringify({ error: 'Invalid roster structure. Must include employees, dateColumns, and rows.' }), {
        status: 400,
        headers: jsonHeaders,
      });
    }

    await kv.put('active_roster', JSON.stringify(payload));

    return new Response(
      JSON.stringify({
        success: true,
        fileName: payload.fileName ?? 'uploaded.xlsx',
        employeesCount: payload.employees.length,
        datesCount: payload.dateColumns.length,
      }),
      {
        status: 200,
        headers: jsonHeaders,
      },
    );
  } catch (error) {
    return new Response(JSON.stringify({ error: 'Failed to save roster to KV', details: String(error) }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}
