import { fakeAsync, TestBed, tick } from '@angular/core/testing';
import { SafeUrl } from '@angular/platform-browser';
import { HttpClientTestingModule, HttpTestingController } from '@angular/common/http/testing';
import { UserPhotoService } from './user-photo.service';
import { URL_GET_USER_PHOTO } from '../data/constants';

describe('UserPhotoService', () => {
  let service: UserPhotoService;
  let httpMock: HttpTestingController;
  let createdUrls: string[];
  let revokedUrls: string[];
  let originalCreate: typeof URL.createObjectURL;
  let originalRevoke: typeof URL.revokeObjectURL;

  beforeEach(() => {
    TestBed.configureTestingModule({
      imports: [HttpClientTestingModule],
      providers: [UserPhotoService],
    });
    service = TestBed.inject(UserPhotoService);
    httpMock = TestBed.inject(HttpTestingController);

    createdUrls = [];
    revokedUrls = [];
    originalCreate = URL.createObjectURL;
    originalRevoke = URL.revokeObjectURL;
    URL.createObjectURL = (blob: any) => {
      const fake = 'blob:fake-' + createdUrls.length;
      createdUrls.push(fake);
      return fake;
    };
    URL.revokeObjectURL = (url: string) => {
      revokedUrls.push(url);
    };
  });

  afterEach(() => {
    httpMock.verify();
    service.ngOnDestroy();
    URL.createObjectURL = originalCreate;
    URL.revokeObjectURL = originalRevoke;
  });

  it('returns null for missing msid without HTTP call', (done) => {
    service.getPhoto(undefined).subscribe((url) => {
      expect(url).toBeNull();
      done();
    });
    httpMock.expectNone(() => true);
  });

  it('returns null for malformed msid without HTTP call', (done) => {
    service.getPhoto('../../etc/passwd').subscribe((url) => {
      expect(url).toBeNull();
      done();
    });
    httpMock.expectNone(() => true);
  });

  it('returns null on 204 No Content', (done) => {
    service.getPhoto('abc123').subscribe((url) => {
      expect(url).toBeNull();
      done();
    });
    const req = httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123');
    req.flush(null, { status: 204, statusText: 'No Content' });
  });

  it('returns null on 404 without bubbling error', (done) => {
    service.getPhoto('abc123').subscribe({
      next: (url) => {
        expect(url).toBeNull();
        done();
      },
      error: () => fail('should not error'),
    });
    const req = httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123');
    req.flush(new Blob(['not found'], { type: 'image/jpeg' }), { status: 404, statusText: 'Not Found' });
  });

  it('wraps blob into SafeUrl on success', (done) => {
    service.getPhoto('abc123').subscribe((url) => {
      expect(url).not.toBeNull();
      expect(createdUrls.length).toBe(1);
      done();
    });
    const req = httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123');
    req.flush(new Blob(['x'], { type: 'image/jpeg' }), { status: 200, statusText: 'OK' });
  });

  it('deduplicates concurrent requests for same msid', () => {
    service.getPhoto('abc123').subscribe();
    service.getPhoto('abc123').subscribe();
    service.getPhoto('ABC123').subscribe();
    const reqs = httpMock.match(URL_GET_USER_PHOTO + 'abc123');
    expect(reqs.length).toBe(1);
    reqs[0].flush(new Blob(['x'], { type: 'image/jpeg' }), { status: 200, statusText: 'OK' });
  });

  for (const mime of ['text/html', 'application/json', 'image/svg+xml', '']) {
    it(`rejects 200 ${mime || 'missing MIME'} without creating a URL and retries after 30 seconds`, fakeAsync(() => {
      const warn = spyOn(console, 'warn');
      const received: (SafeUrl | null)[] = [];
      service.getPhoto('invalidmime', 48).subscribe(url => received.push(url));
      httpMock.expectOne(URL_GET_USER_PHOTO + 'invalidmime?size=48')
        .flush(new Blob(['unexpected'], { type: mime }));
      expect(received).toEqual([null]);
      expect(createdUrls).toEqual([]);
      expect(warn).toHaveBeenCalledWith('[user-photo] request failed', jasmine.objectContaining({
        message: `Unexpected photo content type: ${mime || '(missing)'}`,
      }));
      tick(29_999);
      service.getPhoto('invalidmime', 48).subscribe(url => expect(url).toBeNull());
      httpMock.expectNone(URL_GET_USER_PHOTO + 'invalidmime?size=48');
      tick(2);
      service.getPhoto('invalidmime', 48).subscribe(url => expect(url).not.toBeNull());
      httpMock.expectOne(URL_GET_USER_PHOTO + 'invalidmime?size=48')
        .flush(new Blob(['jpeg'], { type: 'image/jpeg' }));
      expect(createdUrls.length).toBe(1);
    }));
  }

  it('accepts a JPEG header with case and parameters when the blob type is missing', () => {
    service.getPhoto('jpegheader', 512).subscribe(url => expect(url).not.toBeNull());
    httpMock.expectOne(URL_GET_USER_PHOTO + 'jpegheader?size=512')
      .flush(new Blob(['jpeg']), { headers: { 'Content-Type': 'IMAGE/JPEG; charset=binary' } });
    expect(createdUrls.length).toBe(1);
  });

  it('rejects an HTML response header even if the blob claims JPEG and releases the queue slot', () => {
    spyOn(console, 'warn');
    const busy = fillRequestSlots();
    service.getPhoto('afterinvalid', 48).subscribe();
    busy[0].flush(new Blob(['html'], { type: 'image/jpeg' }), {
      headers: { 'Content-Type': 'text/html' },
    });
    expect(createdUrls).toEqual([]);
    httpMock.expectOne(URL_GET_USER_PHOTO + 'afterinvalid?size=48').flush(null);
    busy.slice(1).forEach(request => request.flush(null));
  });

  function fillRequestSlots() {
    return Array.from({ length: 6 }, (_, i) => {
      const id = 'busy' + i;
      service.getPhoto(id, 48).subscribe();
      return httpMock.expectOne(URL_GET_USER_PHOTO + id + '?size=48');
    });
  }

  it('loads 50 distinct photos without exceeding six active HTTP requests', () => {
    let delivered = 0;
    for (let i = 0; i < 50; i++) {
      service.getPhoto('person' + i, 48).subscribe(() => delivered++);
    }
    let batch = httpMock.match(() => true);
    expect(batch.length).toBe(6);
    while (batch.length) {
      batch.forEach(request => request.flush(new Blob(['photo'], { type: 'image/jpeg' })));
      batch = httpMock.match(() => true);
      expect(batch.length).toBeLessThanOrEqual(6);
    }
    expect(delivered).toBe(50);
  });

  it('dispatches visible photos before queued prefetch photos and keeps equal priorities FIFO', () => {
    const busy = fillRequestSlots();
    service.getPhoto('nearby1', 48, () => 'prefetch').subscribe();
    service.getPhoto('nearby2', 48, () => 'prefetch').subscribe();
    service.getPhoto('visible1', 48).subscribe();
    httpMock.expectNone(URL_GET_USER_PHOTO + 'visible1?size=48');
    busy[0].flush(null, { status: 204, statusText: 'No Content' });
    const visible = httpMock.expectOne(URL_GET_USER_PHOTO + 'visible1?size=48');
    expect(visible.request.method).toBe('GET');
    visible.flush(new Blob(['visible'], { type: 'image/jpeg' }));
    httpMock.expectOne(URL_GET_USER_PHOTO + 'nearby1?size=48').flush(new Blob(['nearby'], { type: 'image/jpeg' }));
    httpMock.expectOne(URL_GET_USER_PHOTO + 'nearby2?size=48').flush(new Blob(['nearby'], { type: 'image/jpeg' }));
    busy.slice(1).forEach(request => request.flush(null));
  });

  it('re-evaluates viewport priority when a slot becomes available', () => {
    const busy = fillRequestSlots();
    let nowVisible = false;
    service.getPhoto('firstnearby', 48, () => 'prefetch').subscribe();
    service.getPhoto('scrolltarget', 48, () => nowVisible ? 'visible' : 'prefetch').subscribe();
    nowVisible = true;
    busy[0].flush(null);
    const target = httpMock.expectOne(URL_GET_USER_PHOTO + 'scrolltarget?size=48');
    expect(target.request.method).toBe('GET');
    target.flush(null);
    httpMock.expectOne(URL_GET_USER_PHOTO + 'firstnearby?size=48').flush(null);
    busy.slice(1).forEach(request => request.flush(null));
  });

  it('removes work only after the last queued consumer leaves and allows re-subscription', () => {
    const busy = fillRequestSlots();
    const photo$ = service.getPhoto('sharedqueued', 48);
    const first = photo$.subscribe();
    const second = photo$.subscribe();
    first.unsubscribe();
    expect(second.closed).toBeFalse();
    busy[0].flush(null);
    httpMock.expectOne(URL_GET_USER_PHOTO + 'sharedqueued?size=48').flush(null);
    second.unsubscribe();
    const unwanted$ = service.getPhoto('unwanted', 48);
    // Refill the free slot so unwanted stays queued.
    service.getPhoto('replacementbusy', 48).subscribe();
    const replacement = httpMock.expectOne(URL_GET_USER_PHOTO + 'replacementbusy?size=48');
    const abandoned = unwanted$.subscribe();
    abandoned.unsubscribe();
    busy[1].flush(null);
    httpMock.expectNone(URL_GET_USER_PHOTO + 'unwanted?size=48');
    unwanted$.subscribe();
    httpMock.expectOne(URL_GET_USER_PHOTO + 'unwanted?size=48').flush(null);
    replacement.flush(null);
    busy.slice(2).forEach(request => request.flush(null));
  });

  it('uses the remaining shared consumer priority when a visible consumer leaves', () => {
    const busy = fillRequestSlots();
    service.getPhoto('sharedpriority', 48, () => 'prefetch').subscribe();
    const visible = service.getPhoto('sharedpriority', 48).subscribe();
    service.getPhoto('stillvisible', 48).subscribe();
    visible.unsubscribe();
    busy[0].flush(null);
    const next = httpMock.expectOne(URL_GET_USER_PHOTO + 'stillvisible?size=48');
    expect(next.request.method).toBe('GET');
    next.flush(null);
    httpMock.expectOne(URL_GET_USER_PHOTO + 'sharedpriority?size=48').flush(null);
    busy.slice(1).forEach(request => request.flush(null));
  });

  it('replays cached photos even while every request slot is occupied', () => {
    let photo: SafeUrl | null = null;
    service.getPhoto('cached', 48).subscribe(url => photo = url);
    httpMock.expectOne(URL_GET_USER_PHOTO + 'cached?size=48').flush(new Blob(['cached'], { type: 'image/jpeg' }));
    const busy = fillRequestSlots();
    service.getPhoto('cached', 48).subscribe(url => expect(url).toBe(photo));
    httpMock.expectNone(URL_GET_USER_PHOTO + 'cached?size=48');
    busy.forEach(request => request.flush(null));
  });

  it('keeps an active request alive and cached when its only view leaves', () => {
    const sub = service.getPhoto('started', 48).subscribe();
    const request = httpMock.expectOne(URL_GET_USER_PHOTO + 'started?size=48');
    sub.unsubscribe();
    expect(request.cancelled).toBeFalse();
    request.flush(new Blob(['completed'], { type: 'image/jpeg' }));
    service.getPhoto('started', 48).subscribe(url => expect(url).not.toBeNull());
    httpMock.expectNone(URL_GET_USER_PHOTO + 'started?size=48');
  });

  it('releases a request slot on error and dispatches the next queued photo', () => {
    const warn = spyOn(console, 'warn');
    const busy = fillRequestSlots();
    service.getPhoto('aftererror', 48).subscribe();
    busy[0].flush(null, { status: 503, statusText: 'Unavailable' });
    expect(warn).toHaveBeenCalled();
    httpMock.expectOne(URL_GET_USER_PHOTO + 'aftererror?size=48').flush(null);
    busy.slice(1).forEach(request => request.flush(null));
  });

  it('completes queued consumers without starting HTTP when destroyed', () => {
    const busy = fillRequestSlots();
    let complete = false;
    service.getPhoto('neverstart', 48).subscribe({ complete: () => complete = true });
    service.ngOnDestroy();
    expect(complete).toBeTrue();
    expect(busy.every(request => request.cancelled)).toBeTrue();
    httpMock.expectNone(URL_GET_USER_PHOTO + 'neverstart?size=48');
  });

  it('invalidate removes entry and schedules revoke', fakeAsync(() => {
    service.getPhoto('abc123').subscribe();
    const req = httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123');
    req.flush(new Blob(['x'], { type: 'image/jpeg' }), { status: 200, statusText: 'OK' });
    service.invalidate('abc123');
    expect(revokedUrls.length).toBe(0);
    tick(5_001);
    expect(revokedUrls.length).toBe(1);
    service.getPhoto('abc123').subscribe();
    httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123').flush(new Blob(['y'], { type: 'image/jpeg' }));
  }));

  it('LRU: hot entry survives cache overflow', () => {
    // Fill the 400-entry cache with null results, then force actual eviction.
    for (let i = 0; i < 400; i++) {
      const id = 'u' + String(i).padStart(4, '0');
      service.getPhoto(id).subscribe();
      const r = httpMock.expectOne(URL_GET_USER_PHOTO + id);
      r.flush(null, { status: 204, statusText: 'No Content' });
    }

    // Touch the first-inserted entry to refresh its LRU position
    service.getPhoto('u0000').subscribe();
    httpMock.expectNone(URL_GET_USER_PHOTO + 'u0000');

    // Insert one more distinct entry -> triggers eviction of the actual LRU
    service.getPhoto('u9999').subscribe();
    const r9999 = httpMock.expectOne(URL_GET_USER_PHOTO + 'u9999');
    r9999.flush(null, { status: 204, statusText: 'No Content' });

    // u0000 should still be cached (was refreshed, not the oldest anymore).
    // u0001 was the true LRU after refresh and should be evicted.
    service.getPhoto('u0000').subscribe();
    httpMock.expectNone(URL_GET_USER_PHOTO + 'u0000');

    service.getPhoto('u0001').subscribe();
    const r0001 = httpMock.expectOne(URL_GET_USER_PHOTO + 'u0001');
    expect(r0001.request.method).toBe('GET');
    r0001.flush(null, { status: 204, statusText: 'No Content' });
  });

  function overflowCache(): void {
    for (let i = 0; i < 400; i++) {
      const id = 'overflow' + i;
      service.getPhoto(id, 48).subscribe();
      httpMock.expectOne(URL_GET_USER_PHOTO + id + '?size=48')
        .flush(null, { status: 204, statusText: 'No Content' });
    }
  }

  it('deduplicates an active request even after LRU eviction', () => {
    const received: (SafeUrl | null)[] = [];
    service.getPhoto('activeevicted', 48).subscribe(url => received.push(url));
    const active = httpMock.expectOne(URL_GET_USER_PHOTO + 'activeevicted?size=48');
    overflowCache();
    service.getPhoto('activeevicted', 48).subscribe(url => received.push(url));
    httpMock.expectNone(URL_GET_USER_PHOTO + 'activeevicted?size=48');
    active.flush(new Blob(['shared'], { type: 'image/jpeg' }));
    expect(received.length).toBe(2);
    expect(received[0]).toBe(received[1]);
  });

  it('invalidates active entries outside the LRU and keeps replacement work separate', fakeAsync(() => {
    service.getPhoto('activeevicted', 48).subscribe();
    const old = httpMock.expectOne(URL_GET_USER_PHOTO + 'activeevicted?size=48');
    overflowCache();
    service.invalidate('ACTIVEEVICTED');
    let replacement: SafeUrl | null = null;
    service.getPhoto('activeevicted', 48).subscribe(url => replacement = url);
    const current = httpMock.expectOne(URL_GET_USER_PHOTO + 'activeevicted?size=48');
    old.flush(new Blob(['old'], { type: 'image/jpeg' }));
    current.flush(new Blob(['replacement'], { type: 'image/jpeg' }));
    service.getPhoto('activeevicted', 48).subscribe(url => expect(url).toBe(replacement));
    httpMock.expectNone(URL_GET_USER_PHOTO + 'activeevicted?size=48');
    tick(5_001);
    expect(revokedUrls).toEqual([createdUrls[0]]);
  }));

  it('does not let cancellation of invalidated queued work remove its replacement', () => {
    const busy = fillRequestSlots();
    const old = service.getPhoto('replacequeued', 48).subscribe();
    service.invalidate('replacequeued');
    let received = false;
    service.getPhoto('replacequeued', 48).subscribe(() => received = true);
    old.unsubscribe();
    busy[0].flush(null);
    httpMock.expectOne(URL_GET_USER_PHOTO + 'replacequeued?size=48').flush(null);
    expect(received).toBeTrue();
    busy.slice(1).forEach(request => request.flush(null));
  });

  it('keeps an evicted URL alive until both avatar and modal release it', fakeAsync(() => {
    let photo!: SafeUrl;
    service.getPhoto('abc123', 512).subscribe(url => {
      photo = url!;
      service.retainPhoto(photo);
      service.retainPhoto(photo);
    });
    httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123?size=512').flush(new Blob(['x'], { type: 'image/jpeg' }));
    overflowCache();
    tick(5_001);
    expect(revokedUrls).toEqual([]);
    service.releasePhoto(photo);
    tick(5_001);
    expect(revokedUrls).toEqual([]);
    service.releasePhoto(photo);
    tick(5_001);
    expect(revokedUrls).toEqual([createdUrls[0]]);
  }));

  it('allows a late evicted response to be retained before revocation', fakeAsync(() => {
    let photo!: SafeUrl;
    service.getPhoto('abc123', 48).subscribe(url => {
      photo = url!;
      service.retainPhoto(photo);
    });
    const pending = httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123?size=48');
    overflowCache();
    pending.flush(new Blob(['late'], { type: 'image/jpeg' }));
    tick(5_001);
    expect(revokedUrls).toEqual([]);
    service.releasePhoto(photo);
    tick(5_001);
    expect(revokedUrls).toEqual([createdUrls[0]]);
  }));

  it('cancels scheduled revocation if a view retains the URL during the grace period', fakeAsync(() => {
    let photo!: SafeUrl;
    service.getPhoto('abc123', 48).subscribe(url => photo = url!);
    httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123?size=48').flush(new Blob(['x'], { type: 'image/jpeg' }));
    overflowCache();
    tick(4_000);
    service.retainPhoto(photo);
    tick(2_000);
    expect(revokedUrls).toEqual([]);
    service.releasePhoto(photo);
    tick(5_001);
    expect(revokedUrls).toEqual([createdUrls[0]]);
  }));

  it('releases a late evicted response with no remaining consumer', fakeAsync(() => {
    const subscription = service.getPhoto('abc123', 48).subscribe();
    const pending = httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123?size=48');
    subscription.unsubscribe();
    overflowCache();
    pending.flush(new Blob(['late'], { type: 'image/jpeg' }));
    expect(revokedUrls).toEqual([]);
    tick(5_001);
    expect(revokedUrls).toEqual([createdUrls[0]]);
  }));

  it('retains an invalidated in-flight result until its consumer releases it', fakeAsync(() => {
    let photo!: SafeUrl;
    service.getPhoto('abc123', 48).subscribe(url => {
      photo = url!;
      service.retainPhoto(photo);
    });
    const pending = httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123?size=48');
    service.invalidate('abc123');
    pending.flush(new Blob(['old'], { type: 'image/jpeg' }));
    service.getPhoto('abc123', 48).subscribe();
    httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123?size=48').flush(new Blob(['new'], { type: 'image/jpeg' }));
    tick(5_001);
    expect(revokedUrls).toEqual([]);
    service.releasePhoto(photo);
    tick(5_001);
    expect(revokedUrls).toEqual([createdUrls[0]]);
  }));

  it('cleans up retired URLs and cancels pending requests on destroy', fakeAsync(() => {
    service.getPhoto('abc123', 48).subscribe();
    httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123?size=48').flush(new Blob(['x'], { type: 'image/jpeg' }));
    service.invalidate('abc123');
    service.getPhoto('pending1', 48).subscribe();
    const pending = httpMock.expectOne(URL_GET_USER_PHOTO + 'pending1?size=48');
    service.ngOnDestroy();
    expect(pending.cancelled).toBeTrue();
    expect(revokedUrls).toEqual([createdUrls[0]]);
    tick(5_001);
    expect(revokedUrls.length).toBe(1);
  }));

  it('does not let an old error timer evict a replacement entry', fakeAsync(() => {
    spyOn(console, 'warn');
    service.getPhoto('abc123', 48).subscribe();
    httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123?size=48')
      .flush(null, { status: 503, statusText: 'Service Unavailable' });
    service.invalidate('abc123');
    let replacement: SafeUrl | null = null;
    service.getPhoto('abc123', 48).subscribe(url => replacement = url);
    httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123?size=48').flush(new Blob(['new'], { type: 'image/jpeg' }));
    tick(30_001);
    service.getPhoto('abc123', 48).subscribe(url => expect(url).toBe(replacement));
    httpMock.expectNone(URL_GET_USER_PHOTO + 'abc123?size=48');
  }));

  it('tracks a deferred request subscribed after its entry was evicted', fakeAsync(() => {
    const photo$ = service.getPhoto('deferred1', 48);
    overflowCache();
    photo$.subscribe();
    httpMock.expectOne(URL_GET_USER_PHOTO + 'deferred1?size=48').flush(new Blob(['late'], { type: 'image/jpeg' }));
    service.ngOnDestroy();
    expect(revokedUrls).toEqual([createdUrls[0]]);
    tick(5_001);
    expect(revokedUrls.length).toBe(1);
  }));

  it('silently downgrades 5xx and schedules negative cache eviction', (done) => {
    jasmine.clock().install();
    const warnSpy = spyOn(console, 'warn');
    service.getPhoto('abc123').subscribe((url) => {
      expect(url).toBeNull();
      expect(warnSpy).toHaveBeenCalled();
      jasmine.clock().tick(30_001);

      service.getPhoto('abc123').subscribe();
      const reqs2 = httpMock.match(URL_GET_USER_PHOTO + 'abc123');
      expect(reqs2.length).toBe(1);
      reqs2[0].flush(null, { status: 204, statusText: 'No Content' });
      jasmine.clock().uninstall();
      done();
    });
    const req = httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123');
    req.flush(new Blob(['boom'], { type: 'image/jpeg' }), { status: 503, statusText: 'Service Unavailable' });
  });
});
