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
import { UserPhotoService } from '../../../service/user-photo.service';
import { tracePhoto } from '../../../service/photo-diagnostics';

type AvatarSize = 'sm' | 'md' | 'lg' | 'xl';
type AvatarShape = 'circle' | 'square';
let nextAvatarId = 0;

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
        if (!this.destroyed && !this.loadStarted && entries.some((e) => e.isIntersecting)) {
          this.trace('visible');
          // Native observer callbacks can run outside Angular. Enter the zone
          // even when getPhoto synchronously replays an already cached result.
          this.zone.run(() => this.load());
          this.observer?.disconnect();
          this.observer = undefined;
        }
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
    this.photoUrl = null;
    this.trace('subscribe');
    this.sub = this.photoSvc.getPhoto(this.msid).subscribe({
      next: (url) => {
        this.updatePhoto(url);
      },
      error: () => {
        this.updatePhoto(null);
      },
    });
  }

  private updatePhoto(url: SafeUrl | null): void {
    this.trace('result', { hasPhoto: url !== null });
    this.zone.run(() => {
      this.photoUrl = url;
      this.loaded = true;
      this.cdr.markForCheck();
    });
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
  }
}
