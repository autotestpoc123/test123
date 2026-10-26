import {
  booleanAttribute,
  Component,
  ChangeDetectorRef,
  ElementRef,
  EventEmitter,
  Input,
  NgZone,
  OnChanges,
  OnDestroy,
  OnInit,
  Output,
  SimpleChanges,
} from '@angular/core';
import { SafeUrl } from '@angular/platform-browser';
import { Subscription } from 'rxjs';
import { PhotoRequestPriority, PhotoWidth, UserPhotoService } from '../../../service/user-photo.service';
import { tracePhoto } from '../../../service/photo-diagnostics';

type AvatarSize = 'sm' | 'md' | 'lg' | 'xl';
type AvatarShape = 'circle' | 'square';
let nextAvatarId = 0;

const AVATAR_WIDTH: Record<AvatarSize, PhotoWidth> = {
  sm: 48,
  md: 48,
  lg: 48,
  xl: 512,
};

@Component({
  selector: 'user-avatar',
  templateUrl: './user-avatar.component.html',
  styleUrl: './user-avatar.component.css',
  standalone: false,
})
export class UserAvatarComponent implements OnInit, OnChanges, OnDestroy {
  @Input() msid?: string;
  @Input() fullName?: string;
  @Input() size: AvatarSize = 'sm';
  @Input() shape: AvatarShape = 'circle';
  @Input() variant = '';
  @Input() loading: 'lazy' | 'eager' = 'lazy';
  @Input({ transform: booleanAttribute }) clickable = false;
  @Output() avatarClick = new EventEmitter<Event>();

  photoUrl: SafeUrl | null = null;
  loaded = false;

  private sub?: Subscription;
  private observer?: IntersectionObserver;
  private loadStarted = false;
  private destroyed = false;
  private readonly avatarId = ++nextAvatarId;

  constructor(
    private host: ElementRef<HTMLElement>,
    private photoSvc: UserPhotoService,
    private zone: NgZone,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    this.trace('init', { loading: this.loading });
    if (this.loading === 'eager' || typeof IntersectionObserver === 'undefined') {
      this.load();
      return;
    }
    this.observer = new IntersectionObserver(
      (entries) => {
        if (this.destroyed || this.loaded) return;
        const latest = entries[entries.length - 1];
        if (!latest) return;
        this.zone.run(() => {
          if (latest.isIntersecting && !this.loadStarted) {
            this.trace('visible');
            this.load();
          } else if (!latest.isIntersecting && this.loadStarted) {
            // Remove this view from queued work when it scrolls away. An active
            // HTTP request can finish in the service and be reused on re-entry.
            this.sub?.unsubscribe();
            this.sub = undefined;
            this.loadStarted = false;
          }
        });
      },
      { rootMargin: '100px' }
    );
    this.observer.observe(this.host.nativeElement);
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (!changes['msid'] || changes['msid'].firstChange) {
      return;
    }
    // A reused component must switch users even while the old request is pending.
    if (this.loadStarted) {
      this.load();
    }
  }

  private load(): void {
    this.sub?.unsubscribe();
    this.loadStarted = true;
    this.loaded = false;
    this.setPhoto(null);
    this.trace('subscribe');
    this.sub = this.photoSvc.getPhoto(this.msid, AVATAR_WIDTH[this.size], () => this.photoPriority()).subscribe({
      next: (url) => {
        this.updatePhoto(url);
      },
      error: () => {
        this.updatePhoto(null);
      },
    });
  }

  private photoPriority(): PhotoRequestPriority {
    if (this.loading === 'eager' || typeof window === 'undefined') return 'visible';
    const bounds = this.host.nativeElement.getBoundingClientRect();
    return bounds.bottom > 0 && bounds.top < window.innerHeight
      && bounds.right > 0 && bounds.left < window.innerWidth ? 'visible' : 'prefetch';
  }

  private updatePhoto(url: SafeUrl | null): void {
    this.trace('result', { hasPhoto: url !== null });
    this.zone.run(() => {
      this.setPhoto(url);
      this.loaded = true;
      this.observer?.disconnect();
      this.observer = undefined;
      this.cdr.markForCheck();
    });
  }

  private setPhoto(url: SafeUrl | null): void {
    if (this.photoUrl === url) return;
    this.photoSvc.retainPhoto(url);
    const previous = this.photoUrl;
    this.photoUrl = url;
    this.photoSvc.releasePhoto(previous);
  }

  onImageLoad(event: Event): void {
    const img = event.target as HTMLImageElement;
    this.trace('image-loaded', { width: img.naturalWidth, height: img.naturalHeight });
  }

  onImageError(): void {
    this.trace('image-error');
    this.updatePhoto(null);
  }

  private trace(stage: string, details: Record<string, unknown> = {}): void {
    tracePhoto(this.msid, stage, { avatarId: this.avatarId, size: this.size, ...details });
  }

  get isInteractive(): boolean {
    return this.clickable && this.photoUrl !== null;
  }

  onClick(event: Event): void {
    if (this.isInteractive) {
      this.avatarClick.emit(event);
    }
  }

  onKeydown(event: KeyboardEvent): void {
    if (!this.isInteractive) {
      return;
    }
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      this.avatarClick.emit(event);
    }
  }

  ngOnDestroy(): void {
    this.destroyed = true;
    this.trace('destroy');
    this.observer?.disconnect();
    this.sub?.unsubscribe();
    this.setPhoto(null);
  }
}
