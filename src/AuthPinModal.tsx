// src/AuthPinModal.tsx
import React, { useState, useEffect, useRef } from 'react';
import { ShieldCheck, X, AlertCircle, Loader2, CheckCircle2 } from 'lucide-react';
import { verifyEmployeeCode, syncPushSubscriptionToKv, AUTH_STORAGE_KEY } from './authConfig';
import { getPushSubscription } from './notificationService';

interface AuthPinModalProps {
  isOpen: boolean;
  targetEmployee: string | null;
  onClose: () => void;
  onAuthenticated: (employee: string) => void;
}

export default function AuthPinModal({
  isOpen,
  targetEmployee,
  onClose,
  onAuthenticated,
}: AuthPinModalProps) {
  const [code, setCode] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [isVerifying, setIsVerifying] = useState(false);
  const [success, setSuccess] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      setCode('');
      setError(null);
      setSuccess(false);
      setIsVerifying(false);
      setTimeout(() => {
        inputRef.current?.focus();
      }, 100);
    }
  }, [isOpen, targetEmployee]);

  if (!isOpen || !targetEmployee) return null;

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!code.trim() || isVerifying) return;

    setIsVerifying(true);
    setError(null);

    try {
      const res = await verifyEmployeeCode(targetEmployee, code.trim());
      if (res.success) {
        setSuccess(true);
        // Persist authenticated identity in localStorage
        localStorage.setItem(AUTH_STORAGE_KEY, targetEmployee);

        // Sync push subscription to Cloudflare KV in background if available
        try {
          const sub = await getPushSubscription();
          if (sub) {
            syncPushSubscriptionToKv(targetEmployee, sub.toJSON()).catch(() => {});
          }
        } catch {
          // Non-blocking
        }

        setTimeout(() => {
          onAuthenticated(targetEmployee);
          onClose();
        }, 400);
      } else {
        setError(res.error || 'Incorrect code. Please try again.');
        setIsVerifying(false);
        inputRef.current?.select();
      }
    } catch (err) {
      setError('An error occurred during verification.');
      setIsVerifying(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 backdrop-blur-sm animate-fadeIn">
      <div className="relative w-full max-w-[360px] rounded-3xl border border-zinc-950/10 bg-white/95 p-6 shadow-2xl backdrop-blur-xl dark:border-white/10 dark:bg-zinc-900/95 text-zinc-950 dark:text-white">
        {/* Close Button */}
        <button
          type="button"
          onClick={onClose}
          disabled={isVerifying}
          className="absolute right-4 top-4 flex size-8 items-center justify-center rounded-full text-zinc-400 hover:bg-zinc-950/5 dark:hover:bg-white/10 transition-colors"
        >
          <X className="size-4" />
        </button>

        <div className="flex flex-col items-center text-center">
          <div className="flex size-14 items-center justify-center rounded-2xl bg-blue-500/10 text-blue-500 mb-3.5">
            {success ? <CheckCircle2 className="size-7 text-emerald-500" /> : <ShieldCheck className="size-7" />}
          </div>

          <h3 className="text-[19px] font-bold tracking-tight">Employee Verification</h3>
          <p className="mt-1 text-[13px] text-zinc-500 dark:text-zinc-400">
            Enter the PIN code to authenticate as
          </p>
          <p className="mt-0.5 text-[14px] font-semibold text-blue-600 dark:text-blue-400">
            {targetEmployee}
          </p>

          <form onSubmit={handleSubmit} className="mt-5 w-full space-y-3">
            <div className="relative">
              <input
                ref={inputRef}
                type="password"
                inputMode="numeric"
                pattern="[0-9]*"
                maxLength={8}
                value={code}
                onChange={(e) => setCode(e.target.value.replace(/\D/g, ''))}
                placeholder="Enter Code"
                disabled={isVerifying || success}
                className="w-full rounded-2xl border border-zinc-950/10 bg-zinc-950/5 px-4 py-3 text-center font-mono text-[24px] font-bold tracking-widest outline-none focus:border-blue-500 focus:bg-transparent dark:border-white/10 dark:bg-white/5 dark:focus:border-blue-400 transition-all"
              />
            </div>

            {error && (
              <div className="flex items-center justify-center gap-1.5 text-[12px] font-medium text-red-500 animate-shake">
                <AlertCircle className="size-3.5 shrink-0" />
                <span>{error}</span>
              </div>
            )}

            <button
              type="submit"
              disabled={!code.trim() || isVerifying || success}
              className="flex w-full items-center justify-center gap-2 rounded-2xl bg-blue-600 py-3 text-[15px] font-bold text-white shadow-md shadow-blue-500/20 hover:bg-blue-500 active:scale-[0.98] disabled:opacity-50 transition-all cursor-pointer"
            >
              {isVerifying ? (
                <>
                  <Loader2 className="size-4 animate-spin" /> Verifying...
                </>
              ) : success ? (
                <>
                  <CheckCircle2 className="size-4" /> Verified!
                </>
              ) : (
                'Verify & Authenticate'
              )}
            </button>

            <p className="text-[11px] text-zinc-400 dark:text-zinc-500">
              Your login will stay saved on this device so you can receive notifications.
            </p>
          </form>
        </div>
      </div>
    </div>
  );
}
