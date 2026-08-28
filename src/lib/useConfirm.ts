import { createElement, useRef, useState } from 'react';
import ConfirmDialog from '../components/ConfirmDialog';

interface ConfirmOptions {
  title: string;
  message: string;
  danger?: boolean;
  confirmLabel?: string;
  cancelLabel?: string;
}

/**
 * Promise-based replacement for window.confirm(), so it drops straight into an
 * existing `if (!(await confirm(...))) return;` guard without restructuring the
 * caller. Mirrors the shape of the per-page showToast/toast-state pattern used
 * everywhere else in this app, just promise-based instead of fire-and-forget.
 */
export function useConfirm() {
  const [options, setOptions] = useState<ConfirmOptions | null>(null);
  const resolver = useRef<((v: boolean) => void) | null>(null);

  function confirm(opts: ConfirmOptions): Promise<boolean> {
    setOptions(opts);
    return new Promise((resolve) => {
      resolver.current = resolve;
    });
  }

  function settle(result: boolean) {
    setOptions(null);
    resolver.current?.(result);
    resolver.current = null;
  }

  const dialog = createElement(ConfirmDialog, {
    open: options !== null,
    title: options?.title || '',
    message: options?.message || '',
    danger: options?.danger,
    confirmLabel: options?.confirmLabel,
    cancelLabel: options?.cancelLabel,
    onConfirm: () => settle(true),
    onCancel: () => settle(false),
  });

  return { confirm, dialog };
}
