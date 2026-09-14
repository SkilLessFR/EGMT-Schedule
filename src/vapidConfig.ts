// src/vapidConfig.ts
/**
 * VAPID Application Server Public Key for Web Push Protocol (RFC 8292).
 * Allows browsers (especially iOS Safari 16.4+ in Home Screen PWA mode)
 * to register with Apple Push Notification service (APNs) and receive
 * push notifications even when the app is completely closed.
 */
export const VAPID_PUBLIC_KEY = 'BNvxNLCsValrfhO74mJsm5f2pQEXpomtF_X5dAYwCU3OaWM4O_dKNakyLIyEm74BOZeE6sj8axIlUcZdw4MD3HI';

export function urlBase64ToUint8Array(base64String: string): Uint8Array<ArrayBuffer> {
  const padding = '='.repeat((4 - (base64String.length % 4)) % 4);
  const base64 = (base64String + padding).replace(/-/g, '+').replace(/_/g, '/');
  const rawData = window.atob(base64);
  const buffer = new ArrayBuffer(rawData.length);
  const outputArray = new Uint8Array(buffer);
  for (let i = 0; i < rawData.length; ++i) {
    outputArray[i] = rawData.charCodeAt(i);
  }
  return outputArray;
}
