import { isDevMode } from '@angular/core';

/** Enable in DevTools with sessionStorage.setItem('photoDebug', '1'). */
export function tracePhoto(msid: string | null | undefined, stage: string,
  details: Record<string, unknown> = {}): void {
  if (!isDevMode() || typeof window === 'undefined') return;
  try {
    if (window.sessionStorage.getItem('photoDebug') !== '1') return;
    // A string keeps the timestamp and values as they were at the event.
    console.info('[photo-debug] ' + JSON.stringify({
      atMs: Math.round(performance.now()), msid, stage, ...details,
    }));
  } catch {
    // Diagnostics must never interfere with loading when storage is unavailable.
  }
}
