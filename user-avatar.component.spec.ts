import { CommonModule } from '@angular/common';
import { NgZone } from '@angular/core';
import { ComponentFixture, TestBed } from '@angular/core/testing';
import { DomSanitizer, SafeUrl } from '@angular/platform-browser';
import { of, Subject } from 'rxjs';
import { UserAvatarComponent } from './user-avatar.component';
import { UserPhotoService } from '../../../service/user-photo.service';

const TEST_PHOTO = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';

class MockPhotoService {
  getPhoto = jasmine.createSpy('getPhoto').and.returnValue(of(null));
  retainPhoto = jasmine.createSpy('retainPhoto');
  releasePhoto = jasmine.createSpy('releasePhoto');
}

describe('UserAvatarComponent', () => {
  let fixture: ComponentFixture<UserAvatarComponent>;
  let component: UserAvatarComponent;
  let photoSvc: MockPhotoService;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CommonModule],
      declarations: [UserAvatarComponent],
      providers: [{ provide: UserPhotoService, useClass: MockPhotoService }],
    }).compileComponents();

    fixture = TestBed.createComponent(UserAvatarComponent);
    component = fixture.componentInstance;
    photoSvc = TestBed.inject(UserPhotoService) as unknown as MockPhotoService;
  });

  it('renders default fa-user icon when loaded with no photo', () => {
    const original = (window as any).IntersectionObserver;
    (window as any).IntersectionObserver = undefined;
    try {
      component.msid = 'abc123';
      fixture.detectChanges();
      const icon = fixture.nativeElement.querySelector('.avatar-default .fa-user');
      expect(icon).not.toBeNull();
      expect(fixture.nativeElement.querySelector('img')).toBeNull();
    } finally {
      (window as any).IntersectionObserver = original;
    }
  });

  it('renders placeholder before load completes', () => {
    fixture.detectChanges();
    const placeholder = fixture.nativeElement.querySelector('.avatar-placeholder');
    expect(placeholder).not.toBeNull();
  });

  it('renders a cached photo after an observer callback outside Angular without a click', async () => {
    const original = window.IntersectionObserver;
    let callback!: IntersectionObserverCallback;
    const disconnect = jasmine.createSpy('disconnect');
    window.IntersectionObserver = class {
      constructor(cb: IntersectionObserverCallback) { callback = cb; }
      observe() {}
      disconnect = disconnect;
    } as unknown as typeof IntersectionObserver;
    try {
      const url = TestBed.inject(DomSanitizer).bypassSecurityTrustUrl(TEST_PHOTO);
      photoSvc.getPhoto.and.returnValue(of(url));
      component.msid = 'abc123';
      fixture.autoDetectChanges();
      expect(fixture.nativeElement.querySelector('img')).toBeNull();

      TestBed.inject(NgZone).runOutsideAngular(() => {
        callback([{ isIntersecting: true } as IntersectionObserverEntry], {} as IntersectionObserver);
      });
      await fixture.whenStable();

      expect(fixture.nativeElement.querySelector('img')).not.toBeNull();
      expect(disconnect).toHaveBeenCalled();
    } finally {
      window.IntersectionObserver = original;
    }
  });

  it('renders a photo emitted outside Angular without another user interaction', async () => {
    const original = window.IntersectionObserver;
    (window as any).IntersectionObserver = undefined;
    try {
      const photos = new Subject<any>();
      photoSvc.getPhoto.and.returnValue(photos);
      component.msid = 'abc123';
      fixture.autoDetectChanges();
      expect(fixture.nativeElement.querySelector('.avatar-placeholder')).not.toBeNull();

      const url = TestBed.inject(DomSanitizer).bypassSecurityTrustUrl(TEST_PHOTO);
      TestBed.inject(NgZone).runOutsideAngular(() => photos.next(url));
      await fixture.whenStable();

      expect(fixture.nativeElement.querySelector('img')).not.toBeNull();
      expect(fixture.nativeElement.querySelector('.avatar-placeholder')).toBeNull();
    } finally {
      window.IntersectionObserver = original;
    }
  });

  it('keeps the 100px margin, leaves queued work on exit and resumes on re-entry', () => {
    const original = window.IntersectionObserver;
    let callback!: IntersectionObserverCallback;
    let options: IntersectionObserverInit | undefined;
    const disconnect = jasmine.createSpy('disconnect');
    window.IntersectionObserver = class {
      constructor(cb: IntersectionObserverCallback, init?: IntersectionObserverInit) {
        callback = cb;
        options = init;
      }
      observe() {}
      disconnect = disconnect;
    } as unknown as typeof IntersectionObserver;
    try {
      const first = new Subject<SafeUrl | null>();
      const second = new Subject<SafeUrl | null>();
      photoSvc.getPhoto.and.returnValues(first, second);
      component.msid = 'abc123';
      fixture.detectChanges();
      expect(options?.rootMargin).toBe('100px');
      expect(photoSvc.getPhoto).not.toHaveBeenCalled();
      callback([{ isIntersecting: true } as IntersectionObserverEntry], {} as IntersectionObserver);
      expect(first.observed).toBeTrue();
      expect(disconnect).not.toHaveBeenCalled();
      callback([{ isIntersecting: false } as IntersectionObserverEntry], {} as IntersectionObserver);
      expect(first.observed).toBeFalse();
      first.next(TEST_PHOTO);
      expect(component.photoUrl).toBeNull();
      callback([{ isIntersecting: true } as IntersectionObserverEntry], {} as IntersectionObserver);
      expect(second.observed).toBeTrue();
      expect(photoSvc.getPhoto).toHaveBeenCalledTimes(2);
      second.next(TEST_PHOTO);
      expect(component.photoUrl).toBe(TEST_PHOTO);
      expect(disconnect).toHaveBeenCalled();
    } finally {
      window.IntersectionObserver = original;
    }
  });

  it('supplies live viewport priority and unsubscribes when destroyed', () => {
    const pending = new Subject<SafeUrl | null>();
    photoSvc.getPhoto.and.returnValue(pending);
    const original = window.IntersectionObserver;
    let callback!: IntersectionObserverCallback;
    window.IntersectionObserver = class {
      constructor(cb: IntersectionObserverCallback) { callback = cb; }
      observe() {}
      disconnect() {}
    } as unknown as typeof IntersectionObserver;
    try {
      const bounds = spyOn(fixture.nativeElement, 'getBoundingClientRect')
        .and.returnValue(new DOMRect(0, window.innerHeight + 50, 24, 24));
      component.msid = 'abc123';
      fixture.detectChanges();
      callback([{ isIntersecting: true } as IntersectionObserverEntry], {} as IntersectionObserver);
      const priority = photoSvc.getPhoto.calls.mostRecent().args[2];
      expect(priority()).toBe('prefetch');
      bounds.and.returnValue(new DOMRect(0, 10, 24, 24));
      expect(priority()).toBe('visible');
      fixture.destroy();
      expect(pending.observed).toBeFalse();
    } finally {
      window.IntersectionObserver = original;
    }
  });

  it('loads photo when IntersectionObserver is unavailable', () => {
    const original = (window as any).IntersectionObserver;
    (window as any).IntersectionObserver = undefined;
    try {
      component.msid = 'abc123';
      fixture.detectChanges();
      expect(photoSvc.getPhoto).toHaveBeenCalledWith('abc123', 48, jasmine.any(Function));
    } finally {
      (window as any).IntersectionObserver = original;
    }
  });

  it('requests an eager header photo without waiting for intersection', () => {
    component.msid = 'abc123';
    component.size = 'xl';
    component.loading = 'eager';
    fixture.detectChanges();
    expect(photoSvc.getPhoto).toHaveBeenCalledOnceWith('abc123', 512, jasmine.any(Function));
  });

  it('switches a pending request to the new user and ignores the old result', () => {
    const oldPhoto = new Subject<SafeUrl | null>();
    const newPhoto = new Subject<SafeUrl | null>();
    photoSvc.getPhoto.and.callFake((msid: string) => msid === 'abc123' ? oldPhoto : newPhoto);
    component.loading = 'eager';
    fixture.componentRef.setInput('msid', 'abc123');
    fixture.detectChanges();
    expect(component.loaded).toBeFalse();

    fixture.componentRef.setInput('msid', 'xyz789');
    fixture.detectChanges();
    expect(photoSvc.getPhoto).toHaveBeenCalledWith('xyz789', 48, jasmine.any(Function));

    oldPhoto.next(TEST_PHOTO);
    expect(component.loaded).toBeFalse();
    expect(component.photoUrl).toBeNull();

    newPhoto.next(TEST_PHOTO);
    fixture.detectChanges();
    expect(component.loaded).toBeTrue();
    expect(fixture.nativeElement.querySelector('img').getAttribute('src')).toBe(TEST_PHOTO);
  });

  it('renders the fallback if the downloaded image cannot be decoded', () => {
    component.loading = 'eager';
    component.msid = 'abc123';
    photoSvc.getPhoto.and.returnValue(of(TEST_PHOTO));
    fixture.detectChanges();
    fixture.nativeElement.querySelector('img').dispatchEvent(new Event('error'));
    fixture.detectChanges();
    expect(fixture.nativeElement.querySelector('img')).toBeNull();
    expect(fixture.nativeElement.querySelector('.avatar-default')).not.toBeNull();
  });

  it('reloads on msid change after first load', () => {
    const original = (window as any).IntersectionObserver;
    (window as any).IntersectionObserver = undefined;
    try {
      component.msid = 'abc123';
      fixture.detectChanges();
      expect(component.loaded).toBeTrue();
      photoSvc.getPhoto.calls.reset();

      component.msid = 'xyz789';
      component.ngOnChanges({
        msid: { previousValue: 'abc123', currentValue: 'xyz789', firstChange: false, isFirstChange: () => false },
      });
      expect(photoSvc.getPhoto).toHaveBeenCalledWith('xyz789', 48, jasmine.any(Function));
    } finally {
      (window as any).IntersectionObserver = original;
    }
  });

  it('applies the xl size class', () => {
    component.size = 'xl';
    fixture.detectChanges();
    const root = fixture.nativeElement.querySelector('.user-avatar');
    expect(root.classList.contains('xl')).toBeTrue();
  });

  it('retains its photo until destruction', () => {
    component.loading = 'eager';
    component.msid = 'abc123';
    const url = TestBed.inject(DomSanitizer).bypassSecurityTrustUrl(TEST_PHOTO);
    photoSvc.getPhoto.and.returnValue(of(url));
    fixture.detectChanges();
    expect(photoSvc.retainPhoto).toHaveBeenCalledOnceWith(url);
    fixture.destroy();
    expect(photoSvc.releasePhoto).toHaveBeenCalledWith(url);
  });

  it('releases its photo if the image fails to load', () => {
    component.loading = 'eager';
    component.msid = 'abc123';
    const url = TestBed.inject(DomSanitizer).bypassSecurityTrustUrl(TEST_PHOTO);
    photoSvc.getPhoto.and.returnValue(of(url));
    fixture.detectChanges();
    component.onImageError();
    expect(photoSvc.releasePhoto).toHaveBeenCalledWith(url);
    expect(component.photoUrl).toBeNull();
  });

  it('does not bind clickable affordances by default', () => {
    component.msid = 'abc123';
    fixture.detectChanges();
    const root = fixture.nativeElement.querySelector('.user-avatar');
    expect(root.classList.contains('clickable')).toBeFalse();
    expect(root.getAttribute('role')).toBeNull();
    expect(root.getAttribute('tabindex')).toBeNull();
  });

  const givePhoto = () => {
    (component as any).photoUrl = 'blob:fake-url';
    fixture.detectChanges();
  };

  it('exposes button role and tabindex when clickable and a photo is loaded', () => {
    component.clickable = true;
    component.fullName = 'Jane Doe';
    fixture.detectChanges();
    givePhoto();
    const root = fixture.nativeElement.querySelector('.user-avatar');
    expect(root.classList.contains('clickable')).toBeTrue();
    expect(root.getAttribute('role')).toBe('button');
    expect(root.getAttribute('tabindex')).toBe('0');
    expect(root.getAttribute('aria-label')).toBe('View photo of Jane Doe');
  });

  it('does not expose clickable affordances when clickable but no photo loaded', () => {
    component.clickable = true;
    component.fullName = 'Jane Doe';
    fixture.detectChanges();
    const root = fixture.nativeElement.querySelector('.user-avatar');
    expect(root.classList.contains('clickable')).toBeFalse();
    expect(root.getAttribute('role')).toBeNull();
    expect(root.getAttribute('tabindex')).toBeNull();
    expect(root.getAttribute('aria-label')).toBeNull();
  });

  it('emits avatarClick with the MouseEvent when clicked, clickable, and photo loaded', () => {
    component.clickable = true;
    fixture.detectChanges();
    givePhoto();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    fixture.nativeElement.querySelector('.user-avatar').click();
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.calls.mostRecent().args[0]).toEqual(jasmine.any(MouseEvent));
  });

  it('does not emit avatarClick when not clickable', () => {
    component.clickable = false;
    fixture.detectChanges();
    givePhoto();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    fixture.nativeElement.querySelector('.user-avatar').click();
    expect(spy).not.toHaveBeenCalled();
  });

  it('does not emit avatarClick when clickable but no photo loaded', () => {
    component.clickable = true;
    fixture.detectChanges();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    fixture.nativeElement.querySelector('.user-avatar').click();
    const enter = new KeyboardEvent('keydown', { key: 'Enter' });
    fixture.nativeElement.querySelector('.user-avatar').dispatchEvent(enter);
    expect(spy).not.toHaveBeenCalled();
  });

  it('emits avatarClick with the KeyboardEvent on Enter when clickable and photo loaded', () => {
    component.clickable = true;
    fixture.detectChanges();
    givePhoto();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    const event = new KeyboardEvent('keydown', { key: 'Enter' });
    fixture.nativeElement.querySelector('.user-avatar').dispatchEvent(event);
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.calls.mostRecent().args[0]).toBe(event);
  });

  it('emits avatarClick on Space key when clickable and photo loaded', () => {
    component.clickable = true;
    fixture.detectChanges();
    givePhoto();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    const event = new KeyboardEvent('keydown', { key: ' ' });
    fixture.nativeElement.querySelector('.user-avatar').dispatchEvent(event);
    expect(spy).toHaveBeenCalledTimes(1);
  });

  it('ignores other keys', () => {
    component.clickable = true;
    fixture.detectChanges();
    givePhoto();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    const event = new KeyboardEvent('keydown', { key: 'Tab' });
    fixture.nativeElement.querySelector('.user-avatar').dispatchEvent(event);
    expect(spy).not.toHaveBeenCalled();
  });
});
