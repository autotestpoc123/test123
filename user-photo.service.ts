import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Injectable, OnDestroy } from '@angular/core';
import { DomSanitizer, SafeUrl } from '@angular/platform-browser';
import { Observable, of, ReplaySubject, Subject } from 'rxjs';
import { catchError, finalize, map, takeUntil } from 'rxjs/operators';
import { URL_GET_USER_PHOTO } from '../data/constants';

const MAX_CACHE_SIZE = 400;
// Absolute freshness from response time. Expired entries are replaced on access;
// displayed photos keep their URL until their existing owners release it.
const PHOTO_CACHE_TTL_MS = 60 * 60_000;
const NEGATIVE_CACHE_MS = 30_000;
const REVOKE_DELAY_MS = 5_000;
const MAX_CONCURRENT_PHOTO_REQUESTS = 4;

export type PhotoWidth = 48 | 256;
export type PhotoRequestPriority = 'visible' | 'prefetch';
type PhotoPriorityProvider = () => PhotoRequestPriority;
const VISIBLE_PRIORITY: PhotoPriorityProvider = () => 'visible';

interface PhotoCacheEntry {
  key: string;
  msid: string;
  size?: PhotoWidth;
  result: ReplaySubject<SafeUrl | null>;
  state: 'queued' | 'active' | 'complete';
  consumers: Map<object, PhotoPriorityProvider>;
  objectUrl?: string;
  safeUrl?: SafeUrl;
  references: number;
  evicted: boolean;
  pending: boolean;
  expiresAt?: number;
  revokeTimer?: ReturnType<typeof setTimeout>;
  negativeTimer?: ReturnType<typeof setTimeout>;
}


@Injectable({ providedIn: 'root' })
export class UserPhotoService implements OnDestroy {
  private cache = new Map<string, PhotoCacheEntry>();
  private readonly entries = new Set<PhotoCacheEntry>();
  private readonly urls = new Map<SafeUrl, PhotoCacheEntry>();
  // In-flight work remains deduplicated even if the LRU evicts its cache entry.
  private readonly inflight = new Map<string, PhotoCacheEntry>();
  private readonly queue: PhotoCacheEntry[] = [];
  private activeRequests = 0;
  private draining = false;
  private readonly destroying$ = new Subject<void>();
  private destroyed = false;
  private readonly unloadHandler = () => this.releaseAll();

  constructor(private http: HttpClient, private sanitizer: DomSanitizer) {
    if (typeof window !== 'undefined') {
      window.addEventListener('beforeunload', this.unloadHandler);
    }
  }

  getPhoto(msid: string | undefined | null, size?: PhotoWidth,
    priority: PhotoPriorityProvider = VISIBLE_PRIORITY): Observable<SafeUrl | null> {
    if (this.destroyed || !msid || !/^[A-Za-z0-9]{2,20}$/.test(msid)) {
      return of(null);
    }
    const key = size !== undefined ? `${msid.toLowerCase()}:${size}` : msid.toLowerCase();

    // Register each subscriber separately: a queued request is unnecessary only
    // when ALL views have left. Re-subscribing resolves the current cache entry.
    return new Observable(subscriber => {
      if (this.destroyed) {
        subscriber.complete();
        return;
      }
      let entry = this.cache.get(key) ?? this.inflight.get(key);
      if (entry?.state === 'complete' && entry.expiresAt !== undefined && entry.expiresAt <= Date.now()) {
        if (this.cache.get(key) === entry) this.cache.delete(key);
        if (this.inflight.get(key) === entry) this.inflight.delete(key);
        this.retire(entry);
        entry = undefined;
      }
      if (!entry) {
        entry = {
          key, msid, size, result: new ReplaySubject<SafeUrl | null>(1),
          state: 'queued', consumers: new Map(),
          references: 0, evicted: false, pending: true,
        };
        this.entries.add(entry);
        this.inflight.set(key, entry);
        this.queue.push(entry);
      }
      // Refresh LRU position, including an active entry rescued from eviction.
      entry.evicted = false;
      this.cache.delete(key);
      this.cache.set(key, entry);
      const consumer = {};
      entry.consumers.set(consumer, priority);
      const subscription = entry.result.subscribe(subscriber);
      const current = entry;
      subscriber.add(() => {
        current.consumers.delete(consumer);
        if (current.state === 'queued' && current.consumers.size === 0) {
          this.cancelQueued(current);
        }
      });
      this.enforceLru();
      this.drainQueue();
      return subscription;
    });
  }

  private cancelQueued(entry: PhotoCacheEntry): void {
    const index = this.queue.indexOf(entry);
    if (index >= 0) this.queue.splice(index, 1);
    if (this.cache.get(entry.key) === entry) this.cache.delete(entry.key);
    if (this.inflight.get(entry.key) === entry) this.inflight.delete(entry.key);
    entry.pending = false;
    entry.state = 'complete';
    entry.result.complete();
    this.retire(entry);
  }

  private drainQueue(): void {
    if (this.destroyed || this.draining) return;
    this.draining = true;
    try {
      while (this.activeRequests < MAX_CONCURRENT_PHOTO_REQUESTS && this.queue.length) {
        // Providers read the latest position at dispatch time; equal priorities
        // stay FIFO. A visible consumer promotes a shared prefetch request.
        const visible = this.queue.findIndex(entry =>
          [...entry.consumers.values()].some(priority => priority() === 'visible'));
        const entry = this.queue.splice(visible >= 0 ? visible : 0, 1)[0];
        this.startRequest(entry);
      }
    } finally {
      this.draining = false;
    }
  }

  private startRequest(entry: PhotoCacheEntry): void {
    const { msid, size, key } = entry;
    let cacheLifetimeMs = PHOTO_CACHE_TTL_MS;
    entry.state = 'active';
    this.activeRequests++;
    const url = URL_GET_USER_PHOTO + encodeURIComponent(msid)
      + (size !== undefined ? `?size=${size}` : '');
    // The service owns active requests. A view unsubscribing only removes work
    // that has not started; active responses still populate the cache.
    this.http.get(url, {
      responseType: 'blob', observe: 'response',
      // Memory misses, including a new page's first load, must validate HTTP
      // cache entries too. Unchanged photos can still use the browser's ETag/304
      // path; in-memory hits above make no HTTP request.
      headers: { 'Cache-Control': 'no-cache' },
    })
      .pipe(
        takeUntil(this.destroying$),
        map((resp) => {
          if (resp.status === 204 || !resp.body || resp.body.size === 0) {
            return null;
          }
          // The photo API returns JPEG only. Proxies/authentication can return
          // HTML with status 200; reject it before creating or trusting a URL.
          const mime = (resp.headers.get('Content-Type') || resp.body.type)
            .split(';')[0].trim().toLowerCase();
          if (mime !== 'image/jpeg') {
            throw new Error(`Unexpected photo content type: ${mime || '(missing)'}`);
          }
          const objectUrl = URL.createObjectURL(resp.body);
          entry.objectUrl = objectUrl;
          entry.safeUrl = this.sanitizer.bypassSecurityTrustUrl(objectUrl);
          this.urls.set(entry.safeUrl, entry);
          return entry.safeUrl;
        }),
        catchError((err: unknown) => {
          const status = err instanceof HttpErrorResponse ? err.status : undefined;
          if (status === 404 || status === 204 || status === 403) {
            return of(null);
          }
          console.warn('[user-photo] request failed', {
            msid,
            status,
            message: err instanceof HttpErrorResponse || err instanceof Error
              ? err.message : 'Unknown photo request error',
          });
          cacheLifetimeMs = NEGATIVE_CACHE_MS;
          entry.negativeTimer = setTimeout(() => {
            if (this.cache.get(key) === entry) {
              this.cache.delete(key);
              if (this.inflight.get(key) === entry) this.inflight.delete(key);
              this.retire(entry);
            }
          }, NEGATIVE_CACHE_MS);
          return of(null);
        }),
        finalize(() => {
          entry.pending = false;
          entry.state = 'complete';
          this.activeRequests--;
          if (this.inflight.get(key) === entry) this.inflight.delete(key);
          if (entry.evicted) this.scheduleRevoke(entry);
          this.drainQueue();
        }),
      ).subscribe({
        next: url => {
          entry.expiresAt = Date.now() + cacheLifetimeMs;
          entry.result.next(url);
        },
        error: error => entry.result.error(error),
        complete: () => entry.result.complete(),
      });
  }

  private enforceLru(): void {
    while (this.cache.size > MAX_CACHE_SIZE) {
      const oldestKey = this.cache.keys().next().value;
      if (!oldestKey) {
        break;
      }
      const entry = this.cache.get(oldestKey);
      this.cache.delete(oldestKey);
      if (entry) this.retire(entry);
    }
  }
  /**
   * Drop every requested size from the cache. Existing consumers retain
   * their photo until releasePhoto(); late responses remain tracked too.
   */
  invalidate(msid: string | undefined | null, size?: PhotoWidth): void {
    if (!msid) {
      return;
    }

    const lower = msid.toLowerCase();
    const keys = new Set<string>();

    if (size !== undefined) {
      keys.add(`${lower}:${size}`);
    } else {
      for (const key of [...this.cache.keys(), ...this.inflight.keys()]) {
        if (key === lower || key.startsWith(`${lower}:`)) keys.add(key);
      }
    }

    keys.forEach((key) => {
      const entry = this.cache.get(key) ?? this.inflight.get(key);
      if (!entry) return;
      this.cache.delete(key);
      if (this.inflight.get(key) === entry) this.inflight.delete(key);
      this.retire(entry);
    });
  }

  /** Call when putting a getPhoto() result into a view; completion is not release. */
  retainPhoto(url: SafeUrl | null): void {
    if (!url) return;
    const entry = this.urls.get(url);
    if (!entry) return;
    entry.references++;
    if (entry.revokeTimer !== undefined) {
      clearTimeout(entry.revokeTimer);
      entry.revokeTimer = undefined;
    }
  }

  /** Pair each retained view with a release on replacement, close or destroy. */
  releasePhoto(url: SafeUrl | null): void {
    if (!url) return;
    const entry = this.urls.get(url);
    if (!entry || entry.references === 0) return;
    entry.references--;
    if (entry.evicted) this.scheduleRevoke(entry);
  }

  private retire(entry: PhotoCacheEntry): void {
    entry.evicted = true;
    if (entry.negativeTimer !== undefined) clearTimeout(entry.negativeTimer);
    this.scheduleRevoke(entry);
  }

  private scheduleRevoke(entry: PhotoCacheEntry): void {
    if (entry.pending || entry.references > 0 || entry.revokeTimer !== undefined) return;
    if (!entry.objectUrl || this.destroyed) {
      this.disposeEntry(entry);
      return;
    }
    // Allow Angular to replace the old src before releasing its last URL.
    entry.revokeTimer = setTimeout(() => {
      entry.revokeTimer = undefined;
      if (entry.evicted && entry.references === 0) this.disposeEntry(entry);
    }, REVOKE_DELAY_MS);
  }

  private disposeEntry(entry: PhotoCacheEntry): void {
    if (entry.revokeTimer !== undefined) clearTimeout(entry.revokeTimer);
    if (entry.negativeTimer !== undefined) clearTimeout(entry.negativeTimer);
    if (entry.objectUrl) URL.revokeObjectURL(entry.objectUrl);
    if (entry.safeUrl) this.urls.delete(entry.safeUrl);
    entry.objectUrl = undefined;
    this.entries.delete(entry);
  }

  private releaseAll(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.destroying$.next();
    this.destroying$.complete();
    this.queue.splice(0).forEach(entry => {
      entry.pending = false;
      entry.state = 'complete';
      entry.result.complete();
    });
    this.inflight.clear();
    this.entries.forEach((entry) => this.disposeEntry(entry));
    this.cache.clear();
  }

  ngOnDestroy(): void {
    if (typeof window !== 'undefined') {
      window.removeEventListener('beforeunload', this.unloadHandler);
    }
    this.releaseAll();
  }
}
