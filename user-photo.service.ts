import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Injectable, OnDestroy } from '@angular/core';
import { DomSanitizer, SafeUrl } from '@angular/platform-browser';
import { defer, EMPTY, Observable, of, Subject } from 'rxjs';
import { catchError, finalize, map, shareReplay, takeUntil } from 'rxjs/operators';
import { URL_GET_USER_PHOTO } from '../data/constants';
import { tracePhoto } from './photo-diagnostics';

const MAX_CACHE_SIZE = 400;
const NEGATIVE_CACHE_MS = 30_000;
const REVOKE_DELAY_MS = 5_000;

export type PhotoWidth = 48 | 512;

interface PhotoCacheEntry {
  observable: Observable<SafeUrl | null>;
  objectUrl?: string;
  safeUrl?: SafeUrl;
  references: number;
  evicted: boolean;
  pending: boolean;
  revokeTimer?: ReturnType<typeof setTimeout>;
  negativeTimer?: ReturnType<typeof setTimeout>;
}


@Injectable({ providedIn: 'root' })
export class UserPhotoService implements OnDestroy {
  private cache = new Map<string, PhotoCacheEntry>();
  private readonly entries = new Set<PhotoCacheEntry>();
  private readonly urls = new Map<SafeUrl, PhotoCacheEntry>();
  private readonly destroying$ = new Subject<void>();
  private destroyed = false;
  private readonly unloadHandler = () => this.releaseAll();

  constructor(private http: HttpClient, private sanitizer: DomSanitizer) {
    if (typeof window !== 'undefined') {
      window.addEventListener('beforeunload', this.unloadHandler);
    }
  }

  getPhoto(msid: string | undefined | null, size?: PhotoWidth): Observable<SafeUrl | null> {
    if (this.destroyed || !msid || !/^[A-Za-z0-9]{2,20}$/.test(msid)) {
      return of(null);
    }
    const key = size !== undefined ? `${msid.toLowerCase()}:${size}` : msid.toLowerCase();

    const cached = this.cache.get(key);
    if (cached) {
      tracePhoto(msid, 'cache-hit');
      // refresh LRU position: move hot entry to the end of insertion order
      this.cache.delete(key);
      this.cache.set(key, cached);
      return cached.observable;
    }

    const entry: PhotoCacheEntry = {
      observable: of(null), references: 0, evicted: false, pending: false,
    };
    this.entries.add(entry);

    const url = URL_GET_USER_PHOTO + encodeURIComponent(msid)
         + (size !== undefined ? `?size=${size}` : '');
    const request$ = defer(() => {
      if (this.destroyed) return EMPTY;
      entry.pending = true;
      this.entries.add(entry);
      tracePhoto(msid, 'request-start', {size});
      return this.http.get(url, {
        responseType: 'blob',
        observe: 'response',
      });
    })
      .pipe(
        takeUntil(this.destroying$),
        map((resp) => {
          tracePhoto(msid, 'response', {
            status: resp.status, bytes: resp.body?.size ?? 0, type: resp.body?.type,
          });
          if (resp.status === 204 || !resp.body || resp.body.size === 0) {
            return null;
          }
          const objectUrl = URL.createObjectURL(resp.body);
          entry.objectUrl = objectUrl;
          entry.safeUrl = this.sanitizer.bypassSecurityTrustUrl(objectUrl);
          this.urls.set(entry.safeUrl, entry);
          return entry.safeUrl;
        }),
        catchError((err: HttpErrorResponse) => {
          tracePhoto(msid, 'request-error', { status: err.status });
          if (err.status === 404 || err.status === 204 || err.status === 403) {
            return of(null);
          }
          console.warn('[user-photo] request failed', {
            msid,
            status: err.status,
            message: err.message,
          });
          entry.negativeTimer = setTimeout(() => {
            if (this.cache.get(key) === entry) {
              this.cache.delete(key);
              this.retire(entry);
            }
          }, NEGATIVE_CACHE_MS);
          return of(null);
        }),
        finalize(() => {
          entry.pending = false;
          if (entry.evicted) this.scheduleRevoke(entry);
        }),
        shareReplay(1)
      );

    entry.observable = request$;
    this.cache.set(key, entry);
    this.enforceLru();
    return request$;
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
    const keys: string[] = [];

    if (size !== undefined) {
      keys.push(`${lower}:${size}`);
    } else {
      this.cache.forEach((_, key) => {
        if (key === lower || key.startsWith(`${lower}:`)) keys.push(key);
      });
    }

    keys.forEach((key) => {
      const entry = this.cache.get(key);
      if (!entry) return;
      this.cache.delete(key);
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
    this.destroyed = true;
    this.destroying$.next();
    this.destroying$.complete();
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
