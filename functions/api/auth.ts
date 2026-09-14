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

const DEFAULT_EMPLOYEE_CODES: Record<string, string> = {
  'Palade Alexandru-Ionut': '697',
  'Debona Andrei-Laurentiu': '4747',
  'Stan Antonio-Claudiu': '1646',
  'Olteanu Ionut-Lucian': '14776',
  'Urzica Ion-Alexandru': '2868',
  'Nica Bogdan-Alexandru': '5816',
  'Toma Mihai': '14331',
  'Gheorghe Adrian Marinescu': '14739',
  'Jimboreanu Adelin-Cosmin': '2732',
  'Tir George Cristian': '5',
  'Prisecaru Alexandru-Ionut': '14845',
  'Stoian Alexandru-Gabriel': '14846',
  'Corneci Stefan-Cristian': '9120',
};

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

export async function onRequestPost(context: PagesContext) {
  try {
    const payload = await context.request.json() as { employee?: string; code?: string };
    const employee = payload.employee?.trim();
    const code = payload.code?.trim();

    if (!employee || !code) {
      return new Response(JSON.stringify({ success: false, error: 'Employee name and verification code are required.' }), {
        status: 400,
        headers: jsonHeaders,
      });
    }

    const kv = context.env.EGMT_ROSTER ?? context.env.EGMT_roster;
    let expectedCode = DEFAULT_EMPLOYEE_CODES[employee];

    // Check Cloudflare KV for dynamic code overrides if configured
    if (kv) {
      try {
        const individualCode = await kv.get(`code:${employee}`, 'text');
        if (individualCode) {
          expectedCode = individualCode.trim();
        } else {
          const rawCodes = await kv.get('employee_codes', 'text');
          if (rawCodes) {
            const parsedCodes = JSON.parse(rawCodes) as Record<string, string>;
            if (parsedCodes[employee]) {
              expectedCode = parsedCodes[employee].trim();
            }
          }
        }
      } catch (kvErr) {
        console.warn('KV read failed during auth check, falling back to defaults:', kvErr);
      }
    }

    if (!expectedCode) {
      return new Response(JSON.stringify({ success: false, error: 'Employee code not configured.' }), {
        status: 404,
        headers: jsonHeaders,
      });
    }

    if (code === expectedCode) {
      return new Response(JSON.stringify({
        success: true,
        employee,
        authenticatedAt: new Date().toISOString(),
      }), {
        status: 200,
        headers: jsonHeaders,
      });
    }

    return new Response(JSON.stringify({ success: false, error: 'Invalid verification code.' }), {
      status: 401,
      headers: jsonHeaders,
    });
  } catch (error) {
    return new Response(JSON.stringify({ success: false, error: 'Failed to process authentication request.', details: String(error) }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}
