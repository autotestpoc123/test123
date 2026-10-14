import { TestBed } from '@angular/core/testing';
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
    req.flush(new Blob(['not found']), { status: 404, statusText: 'Not Found' });
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
    reqs[0].flush(new Blob(['x']), { status: 200, statusText: 'OK' });
  });

  it('invalidate removes entry and schedules revoke', (done) => {
    jasmine.clock().install();
    service.getPhoto('abc123').subscribe(() => {
      service.invalidate('abc123');
      expect(revokedUrls.length).toBe(0);
      jasmine.clock().tick(5_001);
      expect(revokedUrls.length).toBe(1);

      service.getPhoto('abc123').subscribe();
      const reqs2 = httpMock.match(URL_GET_USER_PHOTO + 'abc123');
      expect(reqs2.length).toBe(1);
      reqs2[0].flush(new Blob(['y']), { status: 200, statusText: 'OK' });
      jasmine.clock().uninstall();
      done();
    });
    const req = httpMock.expectOne(URL_GET_USER_PHOTO + 'abc123');
    req.flush(new Blob(['x']), { status: 200, statusText: 'OK' });
  });

  it('LRU: hot entry survives cache overflow', () => {
    // Prime 200 distinct msids, each flushed to 204 (null) so no blob created
    for (let i = 0; i < 200; i++) {
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
    r0001.flush(null, { status: 204, statusText: 'No Content' });
  });

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
    req.flush(new Blob(['boom']), { status: 503, statusText: 'Service Unavailable' });
  });
});
