// functions/api/_webPush.ts
import { buildPushPayload } from '@block65/webcrypto-web-push';

export const VAPID_KEYS = {
  subject: 'mailto:admin@egmt-schedule.pages.dev',
  publicKey: 'BNvxNLCsValrfhO74mJsm5f2pQEXpomtF_X5dAYwCU3OaWM4O_dKNakyLIyEm74BOZeE6sj8axIlUcZdw4MD3HI',
  privateKey: '32y3CDccTAAAHmSlnw9gJnDIsEV3GuUE5YaEgm2_9Gk',
};

export interface PushNotificationData {
  title: string;
  body: string;
  icon?: string;
  badge?: string;
  tag?: string;
  data?: Record<string, unknown>;
}

export interface StoredSubscription {
  endpoint: string;
  keys: {
    p256dh: string;
    auth: string;
  };
}

export async function dispatchWebPush(
  rawSubData: string | Record<string, unknown>,
  notification: PushNotificationData
): Promise<{ success: boolean; status?: number; error?: string }> {
  try {
    let sub: StoredSubscription | null = null;
    if (typeof rawSubData === 'string') {
      const parsed = JSON.parse(rawSubData);
      sub = (parsed.subscription || parsed) as StoredSubscription;
    } else if (typeof rawSubData === 'object' && rawSubData !== null) {
      sub = ((rawSubData as Record<string, unknown>).subscription || rawSubData) as StoredSubscription;
    }

    if (!sub || !sub.endpoint || !sub.keys || !sub.keys.p256dh || !sub.keys.auth) {
      return { success: false, error: 'Invalid subscription structure' };
    }

    const payload = await buildPushPayload(
      {
        data: JSON.stringify({
          title: notification.title,
          body: notification.body,
          icon: notification.icon || '/logo.png',
          badge: notification.badge || '/logo.png',
          tag: notification.tag || `push-${Date.now()}`,
          data: notification.data || { url: '/' },
        }),
        options: {
          urgency: 'high',
          ttl: 86400,
        },
      },
      sub,
      VAPID_KEYS
    );

    // Merge high urgency and Apple APNs priority headers to prevent iOS from deferring/batching delivery
    const mergedHeaders: Record<string, string> = {
      ...payload.headers,
      urgency: 'high',
      'apns-priority': '10',
      'apns-push-type': 'alert',
    };

    const res = await fetch(sub.endpoint, {
      method: 'POST',
      headers: mergedHeaders,
      body: payload.body,
    });

    if (res.ok || res.status === 201) {
      return { success: true, status: res.status };
    }

    const errorText = await res.text().catch(() => '');
    return {
      success: false,
      status: res.status,
      error: `Push service returned ${res.status}: ${errorText}`,
    };
  } catch (err) {
    return {
      success: false,
      error: String(err),
    };
  }
}
