import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Injectable, OnDestroy } from '@angular/core';
import { DomSanitizer, SafeUrl } from '@angular/platform-browser';
import { Observable, of } from 'rxjs';
import { catchError, map, shareReplay } from 'rxjs/operators';
import { URL_GET_USER_PHOTO } from '../data/constants';

const MAX_CACHE_SIZE = 200;
const NEGATIVE_CACHE_MS = 30_000;
const REVOKE_DELAY_MS = 5_000;

interface PhotoCacheEntry {
  observable: Observable<SafeUrl | null>;
  objectUrl?: string;
}

@Injectable({ providedIn: 'root' })
export class UserPhotoService implements OnDestroy {
  private cache = new Map<string, PhotoCacheEntry>();
  private readonly unloadHandler = () => this.releaseAll();

  constructor(private http: HttpClient, private sanitizer: DomSanitizer) {
    if (typeof window !== 'undefined') {
      window.addEventListener('beforeunload', this.unloadHandler);
    }
  }

  getPhoto(msid: string | undefined | null): Observable<SafeUrl | null> {
    if (!msid || !/^[A-Za-z0-9]{2,20}$/.test(msid)) {
      return of(null);
    }
    const key = msid.toLowerCase();

    const cached = this.cache.get(key);
    if (cached) {
      // refresh LRU position: move hot entry to the end of insertion order
      this.cache.delete(key);
      this.cache.set(key, cached);
      return cached.observable;
    }

    const entry: PhotoCacheEntry = { observable: of(null) };

    const request$ = this.http
      .get(URL_GET_USER_PHOTO + encodeURIComponent(msid), {
        responseType: 'blob',
        observe: 'response',
      })
      .pipe(
        map((resp) => {
          if (resp.status === 204 || !resp.body || resp.body.size === 0) {
            return null;
          }
          const objectUrl = URL.createObjectURL(resp.body);
          entry.objectUrl = objectUrl;
          return this.sanitizer.bypassSecurityTrustUrl(objectUrl);
        }),
        catchError((err: HttpErrorResponse) => {
          if (err.status === 404 || err.status === 204 || err.status === 403) {
            return of(null);
          }
          console.warn('[user-photo] request failed', {
            msid,
            status: err.status,
            message: err.message,
          });
          setTimeout(() => this.cache.delete(key), NEGATIVE_CACHE_MS);
          return of(null);
        }),
        shareReplay(1)
      );

    entry.observable = request$;
    this.cache.set(key, entry);
    this.enforceLru();
    return request$;
  }

  private enforceLru(): void {
    // Only evict the cache entry; do NOT revoke objectUrl here because
    // already-rendered <img> elements still reference it. Revocation only
    // happens on explicit invalidate() or page unload.
    while (this.cache.size > MAX_CACHE_SIZE) {
      const oldestKey = this.cache.keys().next().value;
      if (!oldestKey) {
        break;
      }
      this.cache.delete(oldestKey);
    }
  }

  /**
   * Drop cached photo for msid and schedule blob URL revocation.
   * Contract: caller should only invoke this after any in-flight request
   * for this msid has settled (e.g. after a successful upload response).
   * Calling during an in-flight request will leak the resulting blob URL
   * until page unload.
   */
  invalidate(msid: string | undefined | null): void {
    if (!msid) {
      return;
    }
    const key = msid.toLowerCase();
    const entry = this.cache.get(key);
    if (!entry) {
      return;
    }
    this.cache.delete(key);
    const url = entry.objectUrl;
    if (url) {
      // Delay revoke so callers swapping <img src> after invalidate()
      // don't see a broken image during the DOM transition.
      setTimeout(() => URL.revokeObjectURL(url), REVOKE_DELAY_MS);
    }
  }

  private releaseAll(): void {
    this.cache.forEach((entry) => {
      if (entry.objectUrl) {
        URL.revokeObjectURL(entry.objectUrl);
      }
    });
    this.cache.clear();
  }

  ngOnDestroy(): void {
    if (typeof window !== 'undefined') {
      window.removeEventListener('beforeunload', this.unloadHandler);
    }
    this.releaseAll();
  }
}
