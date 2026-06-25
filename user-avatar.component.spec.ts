import { ComponentFixture, TestBed } from '@angular/core/testing';
import { of } from 'rxjs';
import { UserAvatarComponent } from './user-avatar.component';
import { UserPhotoService } from '../../../service/user-photo.service';

class MockPhotoService {
  getPhoto = jasmine.createSpy('getPhoto').and.returnValue(of(null));
}

describe('UserAvatarComponent', () => {
  let fixture: ComponentFixture<UserAvatarComponent>;
  let component: UserAvatarComponent;
  let photoSvc: MockPhotoService;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [UserAvatarComponent],
      providers: [{ provide: UserPhotoService, useClass: MockPhotoService }],
    }).compileComponents();

    fixture = TestBed.createComponent(UserAvatarComponent);
    component = fixture.componentInstance;
    photoSvc = TestBed.inject(UserPhotoService) as unknown as MockPhotoService;
  });

  it('computes initials from full name', () => {
    component.fullName = 'Jane Doe';
    expect(component.initials).toBe('JD');
  });

  it('falls back to ? when fullName missing', () => {
    expect(component.initials).toBe('?');
  });

  it('handles single-word name', () => {
    component.fullName = 'Cher';
    expect(component.initials).toBe('C');
  });

  it('produces stable HSL bgColor from seed', () => {
    component.msid = 'abc123';
    const first = component.bgColor;
    const second = component.bgColor;
    expect(first).toBe(second);
    expect(first).toMatch(/^hsl\(\d+, 45%, 50%\)$/);
  });

  it('loads photo when IntersectionObserver is unavailable', () => {
    const original = (window as any).IntersectionObserver;
    (window as any).IntersectionObserver = undefined;
    try {
      component.msid = 'abc123';
      fixture.detectChanges();
      expect(photoSvc.getPhoto).toHaveBeenCalledWith('abc123');
    } finally {
      (window as any).IntersectionObserver = original;
    }
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
      expect(photoSvc.getPhoto).toHaveBeenCalledWith('xyz789');
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

  it('does not bind clickable affordances by default', () => {
    component.msid = 'abc123';
    fixture.detectChanges();
    const root = fixture.nativeElement.querySelector('.user-avatar');
    expect(root.classList.contains('clickable')).toBeFalse();
    expect(root.getAttribute('role')).toBeNull();
    expect(root.getAttribute('tabindex')).toBeNull();
  });

  it('exposes button role and tabindex when clickable', () => {
    component.clickable = true;
    component.fullName = 'Jane Doe';
    fixture.detectChanges();
    const root = fixture.nativeElement.querySelector('.user-avatar');
    expect(root.classList.contains('clickable')).toBeTrue();
    expect(root.getAttribute('role')).toBe('button');
    expect(root.getAttribute('tabindex')).toBe('0');
    expect(root.getAttribute('aria-label')).toBe('View photo of Jane Doe');
  });

  it('emits avatarClick with the MouseEvent when clicked and clickable', () => {
    component.clickable = true;
    fixture.detectChanges();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    fixture.nativeElement.querySelector('.user-avatar').click();
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.calls.mostRecent().args[0]).toEqual(jasmine.any(MouseEvent));
  });

  it('does not emit avatarClick when not clickable', () => {
    component.clickable = false;
    fixture.detectChanges();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    fixture.nativeElement.querySelector('.user-avatar').click();
    expect(spy).not.toHaveBeenCalled();
  });

  it('emits avatarClick with the KeyboardEvent on Enter when clickable', () => {
    component.clickable = true;
    fixture.detectChanges();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    const event = new KeyboardEvent('keydown', { key: 'Enter' });
    fixture.nativeElement.querySelector('.user-avatar').dispatchEvent(event);
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.calls.mostRecent().args[0]).toBe(event);
  });

  it('emits avatarClick on Space key when clickable', () => {
    component.clickable = true;
    fixture.detectChanges();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    const event = new KeyboardEvent('keydown', { key: ' ' });
    fixture.nativeElement.querySelector('.user-avatar').dispatchEvent(event);
    expect(spy).toHaveBeenCalledTimes(1);
  });

  it('ignores other keys', () => {
    component.clickable = true;
    fixture.detectChanges();
    const spy = jasmine.createSpy('avatarClick');
    component.avatarClick.subscribe(spy);
    const event = new KeyboardEvent('keydown', { key: 'Tab' });
    fixture.nativeElement.querySelector('.user-avatar').dispatchEvent(event);
    expect(spy).not.toHaveBeenCalled();
  });
});
