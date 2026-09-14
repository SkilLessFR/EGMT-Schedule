// src/swapNotificationService.ts
/**
 * Handles sending and receiving shift swap push notifications via Cloudflare KV.
 */
import { showAppNotification } from './notificationService';

export interface SwapRequestItem {
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

export interface SendSwapRequestArgs {
  fromEmployee: string;
  toEmployee: string;
  date: string;
  requesterShift: string;
  candidateShift: string;
  message: string;
}

const LOCAL_INBOX_KEY = 'work-schedule-local-inbox';

export async function sendSwapNotificationRequest(args: SendSwapRequestArgs): Promise<{
  success: boolean;
  message: string;
  pushed?: boolean;
}> {
  // 1. Copy message to clipboard as seamless backup
  try {
    if (typeof navigator !== 'undefined' && navigator.clipboard) {
      await navigator.clipboard.writeText(args.message);
    }
  } catch {
    // Non-critical clipboard copy
  }

  // 2. Dispatch to Cloudflare KV backend
  try {
    const res = await fetch('/api/send-swap-request', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ...args,
        createdAt: new Date().toISOString(),
      }),
    });

    if (res.ok) {
      const data = await res.json();
      return {
        success: true,
        message: data.message || `Swap request sent to ${args.toEmployee}!`,
        pushed: data.hasPushSubscription,
      };
    }
  } catch (err) {
    console.warn('Network send failed, queuing locally:', err);
  }

  // 3. Fallback: store in local storage if offline/pure client mode
  try {
    const localStore = JSON.parse(localStorage.getItem(LOCAL_INBOX_KEY) || '[]');
    localStore.unshift({
      id: crypto.randomUUID(),
      ...args,
      createdAt: new Date().toISOString(),
      status: 'pending',
    });
    localStorage.setItem(LOCAL_INBOX_KEY, JSON.stringify(localStore.slice(0, 20)));
  } catch {
    // Ignore
  }

  // If testing on the same device or self-testing, dispatch local notification
  if (args.fromEmployee === args.toEmployee) {
    showAppNotification(`🔄 Shift Swap Request from ${args.fromEmployee}`, {
      body: `Requested swap for ${args.date} (${args.requesterShift} ↔ ${args.candidateShift}).`,
      tag: `swap-self-test-${Date.now()}`,
    }).catch(() => {});
  }

  return {
    success: true,
    message: `Swap request registered for ${args.toEmployee}! (Text copied to clipboard)`,
  };
}

export async function getIncomingSwapRequests(employee: string): Promise<SwapRequestItem[]> {
  if (!employee) return [];

  // 1. Try Cloudflare Pages Function
  try {
    const res = await fetch(`/api/send-swap-request?employee=${encodeURIComponent(employee)}`);
    if (res.ok) {
      const data = await res.json();
      return (data.requests || []) as SwapRequestItem[];
    }
  } catch {
    // Fall back to local storage below
  }

  // 2. Local fallback
  try {
    const localStore = JSON.parse(localStorage.getItem(LOCAL_INBOX_KEY) || '[]') as SwapRequestItem[];
    return localStore.filter((item) => item.toEmployee === employee);
  } catch {
    return [];
  }
}

export async function dismissIncomingSwapRequest(employee: string, id: string): Promise<boolean> {
  try {
    await fetch(`/api/send-swap-request?employee=${encodeURIComponent(employee)}&id=${encodeURIComponent(id)}`, {
      method: 'DELETE',
    });
  } catch {
    // Non-critical
  }

  try {
    const localStore = JSON.parse(localStorage.getItem(LOCAL_INBOX_KEY) || '[]') as SwapRequestItem[];
    const next = localStore.filter((item) => item.id !== id);
    localStorage.setItem(LOCAL_INBOX_KEY, JSON.stringify(next));
  } catch {
    // Ignore
  }

  return true;
}
