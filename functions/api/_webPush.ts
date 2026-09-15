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
): Promise<{ success: boolean; status?: number; error?: string; deliveredCount?: number }> {
  try {
    let subs: StoredSubscription[] = [];

    if (typeof rawSubData === 'string') {
      try {
        const parsed = JSON.parse(rawSubData);
        if (Array.isArray(parsed.subscriptions)) {
          subs = parsed.subscriptions;
        } else if (parsed.subscription) {
          subs = [parsed.subscription];
        } else if (parsed.endpoint) {
          subs = [parsed];
        }
      } catch {
        return { success: false, error: 'Invalid JSON in subscription data' };
      }
    } else if (typeof rawSubData === 'object' && rawSubData !== null) {
      const obj = rawSubData as Record<string, unknown>;
      if (Array.isArray(obj.subscriptions)) {
        subs = obj.subscriptions as StoredSubscription[];
      } else if (obj.subscription) {
        subs = [obj.subscription as StoredSubscription];
      } else if (obj.endpoint) {
        subs = [obj as unknown as StoredSubscription];
      }
    }

    const validSubs = subs.filter(
      (s) => s && typeof s.endpoint === 'string' && s.keys && s.keys.p256dh && s.keys.auth
    );

    if (validSubs.length === 0) {
      return { success: false, error: 'No valid subscription structures found' };
    }

    let successCount = 0;
    const errors: string[] = [];

    for (const sub of validSubs) {
      try {
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
          successCount++;
        } else {
          const errorText = await res.text().catch(() => '');
          errors.push(`Status ${res.status}: ${errorText}`);
        }
      } catch (subErr) {
        errors.push(String(subErr));
      }
    }

    return {
      success: successCount > 0,
      deliveredCount: successCount,
      error: errors.length > 0 ? errors.join('; ') : undefined,
    };
  } catch (err) {
    return {
      success: false,
      error: String(err),
    };
  }
}
