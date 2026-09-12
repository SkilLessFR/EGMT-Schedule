/**
 * notificationService.ts
 * Comprehensive iOS & modern Web Push notification manager for EGMT Schedule PWA.
 */

export interface IosPwaStatus {
  isIos: boolean;
  isStandalone: boolean;
  canPrompt: boolean;
  statusMessage: string;
}

export const NOTIFICATION_PREF_KEY = 'egmt_notifications_enabled';
export const SHIFT_ALERTS_PREF_KEY = 'egmt_shift_alerts_enabled';

/**
 * Check if the browser supports notifications and service workers.
 */
export function isNotificationSupported(): boolean {
  if (typeof window === 'undefined' || typeof navigator === 'undefined') return false;
  return 'Notification' in window && 'serviceWorker' in navigator;
}

/**
 * Detects iOS environment and whether the PWA is running in standalone mode (Home Screen).
 * On iOS (since 16.4+), Web Push / notifications are ONLY supported when added to Home Screen.
 */
export function getIosPwaStatus(): IosPwaStatus {
  if (typeof window === 'undefined' || typeof navigator === 'undefined') {
    return { isIos: false, isStandalone: false, canPrompt: false, statusMessage: '' };
  }

  const userAgent = navigator.userAgent || '';
  const isIos =
    /iPad|iPhone|iPod/.test(userAgent) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

  const isStandalone =
    window.matchMedia('(display-mode: standalone)').matches ||
    (navigator as unknown as { standalone?: boolean }).standalone === true;

  if (isIos && !isStandalone) {
    return {
      isIos: true,
      isStandalone: false,
      canPrompt: false,
      statusMessage:
        "iOS requires adding this app to your Home Screen before enabling notifications: Tap the Share button (⎋) in Safari, select 'Add to Home Screen', and launch the app from your home screen.",
    };
  }

  return {
    isIos,
    isStandalone,
    canPrompt: isNotificationSupported(),
    statusMessage: isIos
      ? 'Running as iOS Home Screen PWA. Ready for push notifications!'
      : 'Ready for push notifications.',
  };
}

/**
 * Get current notification permission state.
 */
export function getNotificationPermission(): NotificationPermission | 'unsupported' {
  if (typeof window === 'undefined' || !('Notification' in window)) {
    return 'unsupported';
  }
  return Notification.permission;
}

/**
 * Registers the Service Worker at /sw.js
 */
export async function registerServiceWorker(): Promise<ServiceWorkerRegistration | null> {
  if (typeof navigator === 'undefined' || !('serviceWorker' in navigator)) {
    return null;
  }

  try {
    const registration = await navigator.serviceWorker.register('/sw.js', { scope: '/' });
    return registration;
  } catch (err) {
    console.warn('Service Worker registration encountered an error:', err);
    return null;
  }
}

/**
 * Prompts user for notification permission from a direct user gesture (button click).
 * On iOS, this MUST be triggered directly within the onClick event callstack.
 */
export async function requestNotificationPermission(): Promise<NotificationPermission> {
  if (!isNotificationSupported()) {
    console.warn('Notifications not supported in this environment.');
    return 'denied';
  }

  // Ensure service worker is registered
  registerServiceWorker().catch(() => {});

  let permission: NotificationPermission;
  try {
    // Standard modern promise-based permission request
    permission = await Notification.requestPermission();
  } catch {
    // Fallback for older WebKit / Safari callback format
    permission = await new Promise<NotificationPermission>((resolve) => {
      Notification.requestPermission((status) => resolve(status));
    });
  }

  if (permission === 'granted') {
    localStorage.setItem(NOTIFICATION_PREF_KEY, 'true');
    // Dispatch test/welcome notification immediately to confirm
    showAppNotification('Notifications Enabled! 🔔', {
      body: 'EGMT Schedule can now send you shift alerts and alarm reminders.',
      tag: 'welcome-notification',
    }).catch((err) => console.log('Welcome notification suppressed:', err));
  } else {
    localStorage.setItem(NOTIFICATION_PREF_KEY, 'false');
  }

  return permission;
}

/**
 * Display a notification using the active ServiceWorkerRegistration.
 * On iOS Safari PWA standalone, registration.showNotification is required because `new Notification()` fails.
 */
export async function showAppNotification(title: string, options?: NotificationOptions): Promise<boolean> {
  if (typeof window === 'undefined' || !('Notification' in window)) {
    return false;
  }

  if (Notification.permission !== 'granted') {
    return false;
  }

  const notificationOptions: NotificationOptions = {
    icon: '/icon.svg',
    badge: '/icon.svg',
    ...options,
  };

  try {
    if ('serviceWorker' in navigator) {
      let reg = await navigator.serviceWorker.getRegistration();
      if (!reg) {
        reg = await registerServiceWorker();
      }
      if (reg) {
        await reg.showNotification(title, notificationOptions);
        return true;
      }
    }

    // Fallback if service worker registration is unavailable
    if (typeof Notification === 'function') {
      new Notification(title, notificationOptions);
      return true;
    }
  } catch (err) {
    console.warn('Failed to display notification:', err);
  }

  return false;
}

/**
 * Inspect or generate a Push Subscription (for server-sent Web Push if desired).
 */
export async function getPushSubscription(applicationServerKey?: string): Promise<PushSubscription | null> {
  if (typeof navigator === 'undefined' || !('serviceWorker' in navigator)) return null;

  try {
    const reg = await navigator.serviceWorker.ready;
    if (!reg.pushManager) return null;

    let sub = await reg.pushManager.getSubscription();
    if (!sub && applicationServerKey) {
      sub = await reg.pushManager.subscribe({
        userVisibleOnly: true,
        applicationServerKey,
      });
    }
    return sub;
  } catch (err) {
    console.warn('Push subscription retrieval failed:', err);
    return null;
  }
}
