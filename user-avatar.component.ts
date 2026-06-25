import {
  Component,
  ElementRef,
  EventEmitter,
  Input,
  OnChanges,
  OnDestroy,
  OnInit,
  Output,
  SimpleChanges,
} from '@angular/core';
import { SafeUrl } from '@angular/platform-browser';
import { Subscription } from 'rxjs';
import { UserPhotoService } from '../../../service/user-photo.service';

type AvatarSize = 'sm' | 'md' | 'lg' | 'xl';
type AvatarShape = 'circle' | 'square';

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
  @Input() clickable = false;
  @Output() avatarClick = new EventEmitter<Event>();

  photoUrl: SafeUrl | null = null;
  loaded = false;

  private sub?: Subscription;
  private observer?: IntersectionObserver;

  constructor(
    private host: ElementRef<HTMLElement>,
    private photoSvc: UserPhotoService
  ) {}

  ngOnInit(): void {
    if (typeof IntersectionObserver === 'undefined') {
      this.load();
      return;
    }
    this.observer = new IntersectionObserver(
      (entries) => {
        if (entries.some((e) => e.isIntersecting)) {
          this.load();
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
    if (this.loaded) {
      this.loaded = false;
      this.photoUrl = null;
      this.load();
    }
  }

  private load(): void {
    this.sub?.unsubscribe();
    this.sub = this.photoSvc.getPhoto(this.msid).subscribe({
      next: (url) => {
        this.photoUrl = url;
        this.loaded = true;
      },
      error: () => {
        this.photoUrl = null;
        this.loaded = true;
      },
    });
  }

  get initials(): string {
    if (!this.fullName) {
      return '?';
    }
    const parts = this.fullName.trim().split(/\s+/).filter(Boolean);
    if (parts.length === 0) {
      return '?';
    }
    const first = parts[0][0] ?? '';
    const last = parts.length > 1 ? parts[parts.length - 1][0] : '';
    return (first + last).toUpperCase() || '?';
  }

  onClick(event: Event): void {
    if (this.clickable) {
      this.avatarClick.emit(event);
    }
  }

  onKeydown(event: KeyboardEvent): void {
    if (!this.clickable) {
      return;
    }
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      this.avatarClick.emit(event);
    }
  }

  get bgColor(): string {
    const seed = this.msid || this.fullName || '';
    let hash = 0;
    for (let i = 0; i < seed.length; i++) {
      hash = (hash << 5) - hash + seed.charCodeAt(i);
      hash |= 0;
    }
    const hue = Math.abs(hash) % 360;
    return `hsl(${hue}, 45%, 50%)`;
  }

  ngOnDestroy(): void {
    this.observer?.disconnect();
    this.sub?.unsubscribe();
  }
}
