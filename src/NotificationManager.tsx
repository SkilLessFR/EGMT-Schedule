import { useState, useEffect, useCallback } from 'react';
import { Bell, BellOff, CheckCircle2, Smartphone, Info, Send, ChevronDown } from 'lucide-react';
import {
  getIosPwaStatus,
  getNotificationPermission,
  requestNotificationPermission,
  showAppNotification,
  getPushSubscription,
  triggerRemoteTestPush,
  type IosPwaStatus,
  NOTIFICATION_PREF_KEY,
  SHIFT_ALERTS_PREF_KEY,
} from './notificationService';
import { syncPushSubscriptionToKv } from './authConfig';
import { GLASS_CARD, ONLY_ON_SHIFT_NOTIFS_KEY } from './scheduleUtils';

const IOS_SWITCH_ON = '#34c759';
const IOS_SWITCH_OFF = 'rgba(120, 120, 128, 0.16)';

export function useNotifications(authenticatedEmployee?: string | null) {
  const [permission, setPermission] = useState<NotificationPermission | 'unsupported'>('default');
  const [iosStatus, setIosStatus] = useState<IosPwaStatus>({
    isIos: false,
    isStandalone: false,
    canPrompt: false,
    statusMessage: '',
  });
  const [taskAlertsEnabled, setTaskAlertsEnabled] = useState<boolean>(() => {
    const saved = localStorage.getItem(NOTIFICATION_PREF_KEY);
    return saved !== null ? saved === 'true' : true;
  });
  const [shiftAlertsEnabled, setShiftAlertsEnabled] = useState<boolean>(() => {
    const saved = localStorage.getItem(SHIFT_ALERTS_PREF_KEY);
    return saved !== null ? saved === 'true' : true;
  });
  const [onlyOnShiftEnabled, setOnlyOnShiftEnabled] = useState<boolean>(() => {
    const saved = localStorage.getItem(ONLY_ON_SHIFT_NOTIFS_KEY);
    return saved !== null ? saved === 'true' : true;
  });
  const [testSent, setTestSent] = useState(false);

  useEffect(() => {
    setPermission(getNotificationPermission());
    setIosStatus(getIosPwaStatus());
  }, []);

  const handleEnableNotifications = useCallback(async () => {
    // Note: requestNotificationPermission MUST be called directly from user click
    const result = await requestNotificationPermission();
    setPermission(result);
    setIosStatus(getIosPwaStatus());
    if (result === 'granted') {
      setTaskAlertsEnabled(true);
      localStorage.setItem(NOTIFICATION_PREF_KEY, 'true');

      // Sync Web Push subscription to Cloudflare KV for this employee
      const empToSync = authenticatedEmployee || localStorage.getItem('work-schedule-employee');
      if (empToSync) {
        try {
          const sub = await getPushSubscription();
          if (sub) {
            syncPushSubscriptionToKv(empToSync, sub.toJSON()).catch(() => {});
          }
        } catch {
          // Non-blocking
        }
      }
    }
    return result;
  }, [authenticatedEmployee]);

  const [backgroundPushScheduled, setBackgroundPushScheduled] = useState<boolean>(false);

  const handleSendTestNotification = useCallback(async () => {
    const title = authenticatedEmployee ? `Alert for ${authenticatedEmployee} 🔔` : 'EGMT Schedule Notification 🔔';
    const body = authenticatedEmployee
      ? `Push notifications are active for ${authenticatedEmployee}! Shift alerts and alarms will be delivered to this device.`
      : 'Push notifications are fully working on this device!';

    const sent = await showAppNotification(title, {
      body,
      tag: 'test-notification-' + Date.now(),
    });
    if (sent) {
      setTestSent(true);
      setTimeout(() => setTestSent(false), 3500);

      const empToSync = authenticatedEmployee || localStorage.getItem('work-schedule-employee');
      if (empToSync) {
        try {
          const sub = await getPushSubscription();
          if (sub) {
            await syncPushSubscriptionToKv(empToSync, sub.toJSON());
            const bgRes = await triggerRemoteTestPush(empToSync, 4);
            if (bgRes.success) {
              setBackgroundPushScheduled(true);
              setTimeout(() => setBackgroundPushScheduled(false), 7000);
            }
          }
        } catch {
          // Non-blocking
        }
      }
    }
    return sent;
  }, [authenticatedEmployee]);

  const toggleTaskAlerts = useCallback(() => {
    setTaskAlertsEnabled((prev) => {
      const next = !prev;
      localStorage.setItem(NOTIFICATION_PREF_KEY, String(next));
      return next;
    });
  }, []);

  const toggleShiftAlerts = useCallback(() => {
    setShiftAlertsEnabled((prev) => {
      const next = !prev;
      localStorage.setItem(SHIFT_ALERTS_PREF_KEY, String(next));
      return next;
    });
  }, []);

  const toggleOnlyOnShift = useCallback(() => {
    setOnlyOnShiftEnabled((prev) => {
      const next = !prev;
      localStorage.setItem(ONLY_ON_SHIFT_NOTIFS_KEY, String(next));
      return next;
    });
  }, []);

  const toggleMasterPush = useCallback(async (enable?: boolean) => {
    const isCurrentlyActive = (taskAlertsEnabled || shiftAlertsEnabled) && permission !== 'denied';
    const shouldEnable = enable !== undefined ? enable : !isCurrentlyActive;
    if (shouldEnable) {
      if (permission !== 'granted') {
        await handleEnableNotifications();
      }
      setTaskAlertsEnabled(true);
      localStorage.setItem(NOTIFICATION_PREF_KEY, 'true');
      setOnlyOnShiftEnabled(true);
      localStorage.setItem(ONLY_ON_SHIFT_NOTIFS_KEY, 'true');
      setShiftAlertsEnabled(false);
      localStorage.setItem(SHIFT_ALERTS_PREF_KEY, 'false');
    } else {
      setTaskAlertsEnabled(false);
      localStorage.setItem(NOTIFICATION_PREF_KEY, 'false');
      setShiftAlertsEnabled(false);
      localStorage.setItem(SHIFT_ALERTS_PREF_KEY, 'false');
      setOnlyOnShiftEnabled(false);
      localStorage.setItem(ONLY_ON_SHIFT_NOTIFS_KEY, 'false');
    }
  }, [taskAlertsEnabled, shiftAlertsEnabled, permission, handleEnableNotifications]);

  return {
    permission,
    iosStatus,
    taskAlertsEnabled,
    shiftAlertsEnabled,
    onlyOnShiftEnabled,
    testSent,
    backgroundPushScheduled,
    handleEnableNotifications,
    handleSendTestNotification,
    toggleTaskAlerts,
    toggleShiftAlerts,
    toggleOnlyOnShift,
    toggleMasterPush,
  };
}

export function NotificationSettingsCard({ authenticatedEmployee }: { authenticatedEmployee?: string | null }) {
  const {
    permission,
    iosStatus,
    taskAlertsEnabled,
    shiftAlertsEnabled,
    onlyOnShiftEnabled,
    testSent,
    backgroundPushScheduled,
    handleEnableNotifications,
    handleSendTestNotification,
    toggleTaskAlerts,
    toggleShiftAlerts,
    toggleOnlyOnShift,
    toggleMasterPush,
  } = useNotifications(authenticatedEmployee);

  const [isExpanded, setIsExpanded] = useState(false);
  const isMasterPushEnabled = (taskAlertsEnabled || shiftAlertsEnabled) && permission !== 'denied';

  return (
    <div className={`p-4 ${GLASS_CARD} space-y-3 transition-all`}>
      {/* Header with Collapsible Trigger & Master Switch */}
      <div className="flex items-center justify-between gap-3">
        <button
          type="button"
          onClick={() => setIsExpanded((prev) => !prev)}
          className="flex items-center gap-2.5 min-w-0 flex-1 text-left cursor-pointer select-none"
        >
          <div className={`flex size-9 shrink-0 items-center justify-center rounded-2xl transition-colors ${
            isMasterPushEnabled ? 'bg-blue-500/10 text-blue-500' : 'bg-zinc-500/10 text-zinc-400'
          }`}>
            <Bell className="size-5" />
          </div>
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-1.5">
              <h3 className="text-[15px] font-bold truncate">Push Notifications</h3>
              <ChevronDown
                className={`size-4 text-zinc-400 transition-transform duration-200 shrink-0 ${
                  isExpanded ? 'rotate-180' : ''
                }`}
              />
            </div>
            <p className="text-[12px] text-zinc-400 dark:text-zinc-500 truncate">
              {permission === 'denied'
                ? 'Blocked in browser'
                : isMasterPushEnabled
                ? (authenticatedEmployee ? `Active for ${authenticatedEmployee}` : 'Active (Alarms on)')
                : 'All alerts disabled'}
            </p>
          </div>
        </button>

        {/* Master Push Toggle Switch */}
        <div className="flex items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={async () => {
              const nextState = !isMasterPushEnabled;
              await toggleMasterPush(nextState);
              if (nextState && permission !== 'granted') {
                setIsExpanded(true);
              }
            }}
            style={{ backgroundColor: isMasterPushEnabled ? IOS_SWITCH_ON : IOS_SWITCH_OFF }}
            className="relative h-[29px] w-[49px] shrink-0 rounded-full p-0.5 transition-colors duration-200 cursor-pointer"
            aria-label="Toggle all push notifications"
          >
            <div
              className={`h-[25px] w-[25px] rounded-full bg-white shadow-sm ring-1 ring-black/[0.04] transition-transform duration-200 ${
                isMasterPushEnabled ? 'translate-x-[20px]' : 'translate-x-0'
              }`}
            />
          </button>
        </div>
      </div>

      {/* Collapsible Details */}
      {isExpanded && (
        <div className="border-t border-zinc-950/5 dark:border-white/5 pt-3 space-y-3 animate-fadeIn">
          {/* Status Badge */}
          <div className="flex items-center justify-between">
            <span className="text-[12px] text-zinc-400">System Permission:</span>
            {permission === 'granted' ? (
              <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/10 px-2.5 py-0.5 text-[11px] font-semibold text-emerald-500">
                <CheckCircle2 className="size-3" /> Granted
              </span>
            ) : permission === 'denied' ? (
              <span className="inline-flex items-center gap-1 rounded-full bg-red-500/10 px-2.5 py-0.5 text-[11px] font-semibold text-red-500">
                <BellOff className="size-3" /> Blocked
              </span>
            ) : (
              <span className="inline-flex items-center gap-1 rounded-full bg-amber-500/10 px-2.5 py-0.5 text-[11px] font-semibold text-amber-500">
                Not Enabled
              </span>
            )}
          </div>

          {/* Special iOS instructions if on iOS and not yet added to Home Screen */}
          {iosStatus.isIos && !iosStatus.isStandalone && (
            <div className="rounded-2xl border border-amber-500/30 bg-amber-500/10 p-3.5 text-left text-amber-600 dark:text-amber-400">
              <div className="flex items-start gap-2.5">
                <Smartphone className="mt-0.5 size-5 shrink-0" />
                <div className="space-y-1 text-[13px]">
                  <p className="font-bold">iOS Home Screen Step Required</p>
                  <p className="text-[12px] opacity-90">
                    Apple requires Web Push apps to be added to the Home Screen before notifications can be enabled:
                  </p>
                  <ol className="list-decimal space-y-1 pl-4 pt-1 text-[12px] opacity-95">
                    <li>Tap the <strong>Share</strong> button (⎋) in Safari's toolbar.</li>
                    <li>Scroll down and tap <strong>"Add to Home Screen"</strong>.</li>
                    <li>Open <strong>Work Schedule</strong> from your home screen.</li>
                    <li>Tap <strong>Enable Notifications</strong> below.</li>
                  </ol>
                </div>
              </div>
            </div>
          )}

          {/* Main Action Button */}
          {permission !== 'granted' ? (
            <div>
              <button
                type="button"
                onClick={handleEnableNotifications}
                className="flex w-full items-center justify-center gap-2 rounded-2xl bg-blue-600 px-4 py-3 text-[15px] font-bold text-white shadow-lg shadow-blue-500/20 transition-all active:scale-[0.98] hover:bg-blue-500 cursor-pointer"
              >
                <Bell className="size-4" />
                Enable Notifications
              </button>
              <p className="mt-1.5 text-center text-[11px] text-zinc-400">
                {iosStatus.isIos && !iosStatus.isStandalone
                  ? 'Works after adding to Home Screen'
                  : 'Prompts your device to allow notification alerts'}
              </p>
            </div>
          ) : (
            <div className="space-y-3">
              <button
                type="button"
                onClick={handleSendTestNotification}
                className={`flex w-full items-center justify-center gap-2 rounded-2xl px-4 py-2.5 text-[14px] font-bold transition-all active:scale-[0.98] cursor-pointer ${
                  testSent
                    ? 'bg-emerald-600 text-white'
                    : 'bg-zinc-950/5 text-zinc-800 hover:bg-zinc-950/10 dark:bg-white/10 dark:text-white dark:hover:bg-white/15'
                }`}
              >
                {testSent ? <CheckCircle2 className="size-4" /> : <Send className="size-4" />}
                {testSent ? 'Test Notification Dispatched!' : 'Send Test Notification'}
              </button>

              {backgroundPushScheduled && (
                <div className="rounded-2xl border border-cyan-500/30 bg-cyan-500/10 p-3 text-left text-cyan-600 dark:text-cyan-400 animate-fadeIn">
                  <div className="flex items-start gap-2.5">
                    <Smartphone className="mt-0.5 size-4 shrink-0 text-cyan-500" />
                    <div className="space-y-1 text-[12px]">
                      <p className="font-bold">Background Push Scheduled (4s)</p>
                      <p className="text-[11px] opacity-90 leading-relaxed">
                        Close the PWA or lock your iPhone right now! Cloudflare is sending the Web Push notification directly to your device.
                      </p>
                    </div>
                  </div>
                </div>
              )}

              {/* Configuration Toggles */}
              <div className="space-y-2 border-t border-zinc-950/5 pt-3 dark:border-white/5">
                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-[14px] font-semibold">Alarm & Task Alerts</p>
                    <p className="text-[11px] text-zinc-400">Notify when recurring alarms trigger</p>
                  </div>
                  <button
                    type="button"
                    onClick={toggleTaskAlerts}
                    style={{ backgroundColor: taskAlertsEnabled ? IOS_SWITCH_ON : IOS_SWITCH_OFF }}
                    className="relative h-[27px] w-[45px] shrink-0 rounded-full p-0.5 transition-colors duration-200 cursor-pointer"
                  >
                    <div
                      className={`h-[23px] w-[23px] rounded-full bg-white shadow-sm transition-transform duration-200 ${
                        taskAlertsEnabled ? 'translate-x-[18px]' : 'translate-x-0'
                      }`}
                    />
                  </button>
                </div>

                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-[14px] font-semibold">Shift Reminders</p>
                    <p className="text-[11px] text-zinc-400">Receive alert before upcoming shift starts</p>
                  </div>
                  <button
                    type="button"
                    onClick={toggleShiftAlerts}
                    style={{ backgroundColor: shiftAlertsEnabled ? IOS_SWITCH_ON : IOS_SWITCH_OFF }}
                    className="relative h-[27px] w-[45px] shrink-0 rounded-full p-0.5 transition-colors duration-200 cursor-pointer"
                  >
                    <div
                      className={`h-[23px] w-[23px] rounded-full bg-white shadow-sm transition-transform duration-200 ${
                        shiftAlertsEnabled ? 'translate-x-[18px]' : 'translate-x-0'
                      }`}
                    />
                  </button>
                </div>

                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-[14px] font-semibold">Only Alert When On Shift</p>
                    <p className="text-[11px] text-zinc-400">Silence task alarms when off shift or on another shift</p>
                  </div>
                  <button
                    type="button"
                    onClick={toggleOnlyOnShift}
                    style={{ backgroundColor: onlyOnShiftEnabled ? IOS_SWITCH_ON : IOS_SWITCH_OFF }}
                    className="relative h-[27px] w-[45px] shrink-0 rounded-full p-0.5 transition-colors duration-200 cursor-pointer"
                  >
                    <div
                      className={`h-[23px] w-[23px] rounded-full bg-white shadow-sm transition-transform duration-200 ${
                        onlyOnShiftEnabled ? 'translate-x-[18px]' : 'translate-x-0'
                      }`}
                    />
                  </button>
                </div>
              </div>
            </div>
          )}

          {permission === 'denied' && (
            <div className="flex items-start gap-2 rounded-xl bg-red-500/10 p-2.5 text-[12px] text-red-500">
              <Info className="mt-0.5 size-4 shrink-0" />
              <span>
                Notifications are blocked. To re-enable, go to iOS Settings → Safari (or Web App) → Notifications and select "Allow".
              </span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export function NotificationPromptBanner({ onEnable }: { onEnable?: () => void }) {
  const { permission, iosStatus, handleEnableNotifications } = useNotifications();

  if (permission === 'granted' || permission === 'denied') {
    return null;
  }

  const handleClick = async () => {
    await handleEnableNotifications();
    if (onEnable) onEnable();
  };

  return (
    <div className="mb-4 flex items-center justify-between gap-3 rounded-2xl border border-blue-500/30 bg-blue-500/10 p-3 text-blue-600 dark:text-blue-400">
      <div className="flex items-center gap-2.5 min-w-0">
        <Bell className="size-5 shrink-0 animate-bounce" />
        <p className="truncate text-[13px] font-medium">
          {iosStatus.isIos && !iosStatus.isStandalone
            ? 'Add to Home Screen for iOS notifications'
            : 'Enable notifications for shift & alarm alerts'}
        </p>
      </div>
      <button
        type="button"
        onClick={handleClick}
        className="shrink-0 rounded-xl bg-blue-600 px-3 py-1.5 text-[12px] font-bold text-white shadow-sm hover:bg-blue-500 active:scale-95"
      >
        Enable
      </button>
    </div>
  );
}
