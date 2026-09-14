// src/authConfig.ts
/**
 * Employee authentication codes and Cloudflare KV integration configuration.
 */

export const EMPLOYEE_CODES: Record<string, string> = {
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

// Aliases mapping colloquial names to full roster names
export const EMPLOYEE_ALIASES: Record<string, string> = {
  'Alex Palade': 'Palade Alexandru-Ionut',
  'Andrei Debona': 'Debona Andrei-Laurentiu',
  'Antonio Stan': 'Stan Antonio-Claudiu',
  'Lucian Olteanu': 'Olteanu Ionut-Lucian',
  'Alex Urzica': 'Urzica Ion-Alexandru',
  'Bogdan Nica': 'Nica Bogdan-Alexandru',
  'Toma Mihai': 'Toma Mihai',
  'Adrian Marinescu': 'Gheorghe Adrian Marinescu',
  'Adelin Jimboreanu': 'Jimboreanu Adelin-Cosmin',
  'George Tir': 'Tir George Cristian',
  'Alexandru Prisecaru': 'Prisecaru Alexandru-Ionut',
  'Alexandru Stoian': 'Stoian Alexandru-Gabriel',
  'Stefan Corneci': 'Corneci Stefan-Cristian',
};

export const AUTH_STORAGE_KEY = 'work-schedule-auth-user';
export const AUTH_TOKEN_STORAGE_KEY = 'work-schedule-auth-token';

/**
 * Resolve the expected authentication PIN code for an employee name.
 */
export function getExpectedCodeForEmployee(employeeName: string): string | null {
  if (!employeeName) return null;
  const trimmed = employeeName.trim();

  // 1. Direct match
  if (EMPLOYEE_CODES[trimmed]) return EMPLOYEE_CODES[trimmed];

  // 2. Alias match
  const canonical = EMPLOYEE_ALIASES[trimmed];
  if (canonical && EMPLOYEE_CODES[canonical]) return EMPLOYEE_CODES[canonical];

  // 3. Normalized lowercase search
  const lower = trimmed.toLowerCase();
  for (const [name, code] of Object.entries(EMPLOYEE_CODES)) {
    if (name.toLowerCase() === lower) return code;
  }

  // 4. Substring partial match
  for (const [name, code] of Object.entries(EMPLOYEE_CODES)) {
    const parts = lower.split(/[\s-]+/).filter(Boolean);
    const nameLower = name.toLowerCase();
    if (parts.length >= 2 && parts.every((p) => nameLower.includes(p))) {
      return code;
    }
  }

  return null;
}

/**
 * Verify employee code locally or via Cloudflare KV /api/auth.
 */
export async function verifyEmployeeCode(employee: string, code: string): Promise<{ success: boolean; error?: string }> {
  const cleanCode = code.trim();
  if (!cleanCode) {
    return { success: false, error: 'Please enter your verification code.' };
  }

  // 1. Attempt Cloudflare Pages API check if available
  try {
    const res = await fetch('/api/auth', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ employee, code: cleanCode }),
    });

    if (res.ok) {
      const data = await res.json();
      if (data.success) {
        return { success: true };
      }
    }
  } catch {
    // API not reachable or running in pure client-side mode, fall back to local verification
  }

  // 2. Local fallback verification
  const expected = getExpectedCodeForEmployee(employee);
  if (!expected) {
    return { success: false, error: 'Employee not found in authentication roster.' };
  }

  if (expected === cleanCode) {
    return { success: true };
  }

  return { success: false, error: 'Incorrect verification code for ' + employee };
}

/**
 * Register push subscription to Cloudflare KV for the authenticated employee.
 */
export async function syncPushSubscriptionToKv(employee: string, subscription: unknown): Promise<boolean> {
  if (!employee || !subscription) return false;

  try {
    const res = await fetch('/api/subscribe', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ employee, subscription }),
    });
    return res.ok;
  } catch (err) {
    console.warn('Could not sync push subscription to Cloudflare KV:', err);
    return false;
  }
}
