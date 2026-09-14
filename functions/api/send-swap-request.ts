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

export interface SwapRequestPayload {
  id: string;
  fromEmployee: string;
  toEmployee: string;
  date: string;
  requesterShift: string;
  candidateShift: string;
  message: string;
  createdAt: string;
  status: 'pending' | 'accepted' | 'declined';
}

const jsonHeaders = {
  'Content-Type': 'application/json',
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export async function onRequestOptions() {
  return new Response(null, {
    status: 204,
    headers: jsonHeaders,
  });
}

// GET /api/send-swap-request?employee=...
export async function onRequestGet(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound' }), {
      status: 503,
      headers: jsonHeaders,
    });
  }

  const url = new URL(context.request.url);
  const employee = url.searchParams.get('employee')?.trim();

  if (!employee) {
    return new Response(JSON.stringify({ error: 'Employee query parameter is required' }), {
      status: 400,
      headers: jsonHeaders,
    });
  }

  try {
    const rawInbox = await kv.get(`inbox:${employee}`, 'text');
    const requests: SwapRequestPayload[] = rawInbox ? JSON.parse(rawInbox) : [];

    // Also check if the recipient has an active Web Push subscription on file
    const rawSub = await kv.get(`sub:${employee}`, 'text');

    return new Response(JSON.stringify({
      employee,
      requests,
      hasPushSubscription: Boolean(rawSub),
    }), {
      status: 200,
      headers: jsonHeaders,
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: 'Failed to retrieve swap requests from KV', details: String(err) }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}

// POST /api/send-swap-request
export async function onRequestPost(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound' }), {
      status: 503,
      headers: jsonHeaders,
    });
  }

  try {
    const payload = await context.request.json() as Partial<SwapRequestPayload>;
    const fromEmployee = payload.fromEmployee?.trim();
    const toEmployee = payload.toEmployee?.trim();
    const date = payload.date?.trim();

    if (!fromEmployee || !toEmployee || !date) {
      return new Response(JSON.stringify({ error: 'fromEmployee, toEmployee, and date are required.' }), {
        status: 400,
        headers: jsonHeaders,
      });
    }

    const newRequest: SwapRequestPayload = {
      id: payload.id || crypto.randomUUID(),
      fromEmployee,
      toEmployee,
      date,
      requesterShift: payload.requesterShift || 'OFF',
      candidateShift: payload.candidateShift || 'OFF',
      message: payload.message || `${fromEmployee} requested a shift swap for ${date}.`,
      createdAt: payload.createdAt || new Date().toISOString(),
      status: 'pending',
    };

    // 1. Append to recipient's inbox in Cloudflare KV
    const rawInbox = await kv.get(`inbox:${toEmployee}`, 'text');
    const inbox: SwapRequestPayload[] = rawInbox ? JSON.parse(rawInbox) : [];
    
    // Filter out duplicates if already sent recently
    const filteredInbox = inbox.filter((r) => !(r.fromEmployee === fromEmployee && r.date === date && r.status === 'pending'));
    filteredInbox.unshift(newRequest);

    // Limit inbox size to recent 30 items
    const trimmedInbox = filteredInbox.slice(0, 30);
    await kv.put(`inbox:${toEmployee}`, JSON.stringify(trimmedInbox));

    // 2. Dispatch real Web Push notification if recipient has an active subscription in KV
    const rawSub = await kv.get(`sub:${toEmployee}`, 'text');
    let pushDispatched = false;
    let pushError: string | undefined;

    if (rawSub) {
      const pushRes = await dispatchWebPush(rawSub, {
        title: `🔄 Shift Swap Request from ${fromEmployee}`,
        body: `${fromEmployee} requested to swap shifts with you for ${date} (${payload.requesterShift || 'OFF'} ↔ ${payload.candidateShift || 'OFF'}).`,
        icon: '/logo.png',
        badge: '/logo.png',
        tag: `swap-req-${newRequest.id}`,
        data: { url: '/', id: newRequest.id },
      });
      pushDispatched = pushRes.success;
      if (!pushRes.success) {
        pushError = pushRes.error;
      }
    }

    return new Response(JSON.stringify({
      success: true,
      requestId: newRequest.id,
      toEmployee,
      hasPushSubscription: Boolean(rawSub),
      pushDispatched,
      pushError,
      message: pushDispatched
        ? `Notification delivered directly to ${toEmployee}'s phone!`
        : Boolean(rawSub)
          ? `Swap request queued (push delivery attempted).`
          : `Swap request queued in ${toEmployee}'s inbox.`,
    }), {
      status: 200,
      headers: jsonHeaders,
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: 'Failed to send swap request', details: String(error) }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}

// DELETE /api/send-swap-request?employee=...&id=...
export async function onRequestDelete(context: PagesContext) {
  const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
  if (!kv) {
    return new Response(JSON.stringify({ error: 'KV namespace EGMT_ROSTER is not bound' }), {
      status: 503,
      headers: jsonHeaders,
    });
  }

  const url = new URL(context.request.url);
  const employee = url.searchParams.get('employee')?.trim();
  const requestId = url.searchParams.get('id')?.trim();

  if (!employee || !requestId) {
    return new Response(JSON.stringify({ error: 'employee and id query parameters are required' }), {
      status: 400,
      headers: jsonHeaders,
    });
  }

  try {
    const rawInbox = await kv.get(`inbox:${employee}`, 'text');
    if (rawInbox) {
      const inbox: SwapRequestPayload[] = JSON.parse(rawInbox);
      const updatedInbox = inbox.filter((r) => r.id !== requestId);
      await kv.put(`inbox:${employee}`, JSON.stringify(updatedInbox));
    }

    return new Response(JSON.stringify({ success: true, message: 'Request dismissed' }), {
      status: 200,
      headers: jsonHeaders,
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: 'Failed to delete request from KV', details: String(err) }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}
