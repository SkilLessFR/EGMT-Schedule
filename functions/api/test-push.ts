// functions/api/test-push.ts
import { dispatchWebPush } from './_webPush';

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
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export async function onRequestOptions() {
  return new Response(null, {
    status: 204,
    headers: jsonHeaders,
  });
}

// POST /api/test-push
export async function onRequestPost(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound' }), {
      status: 503,
      headers: jsonHeaders,
    });
  }

  try {
    const body = (await context.request.json().catch(() => ({}))) as {
      employee?: string;
      delaySeconds?: number;
    };
    const employee = body.employee?.trim();

    if (!employee) {
      return new Response(JSON.stringify({ error: 'Employee name is required' }), {
        status: 400,
        headers: jsonHeaders,
      });
    }

    const rawSub = await kv.get(`sub:${employee}`, 'text');
    if (!rawSub) {
      return new Response(
        JSON.stringify({
          error: `No Web Push subscription found for ${employee}. Please enable notifications in the PWA on your home screen first.`,
        }),
        { status: 404, headers: jsonHeaders }
      );
    }

    const delay = Math.min(Math.max(body.delaySeconds ?? 3, 0), 30);
    if (delay > 0) {
      await new Promise((resolve) => setTimeout(resolve, delay * 1000));
    }

    const result = await dispatchWebPush(rawSub, {
      title: `EGMT Schedule Alert 🔔`,
      body: `Hello ${employee}! Background push notifications are active even when the app is closed.`,
      icon: '/logo.png',
      badge: '/logo.png',
      tag: `test-background-${Date.now()}`,
      data: { url: '/', type: 'test' },
    });

    if (result.success) {
      return new Response(
        JSON.stringify({
          success: true,
          message: `Background push notification delivered to ${employee}'s device!`,
        }),
        { status: 200, headers: jsonHeaders }
      );
    }

    return new Response(
      JSON.stringify({
        error: `Failed to deliver push to device: ${result.error}`,
      }),
      { status: 502, headers: jsonHeaders }
    );
  } catch (err) {
    return new Response(
      JSON.stringify({ error: 'Failed to process test push', details: String(err) }),
      { status: 500, headers: jsonHeaders }
    );
  }
}
