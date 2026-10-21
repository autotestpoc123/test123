import * as _ from 'underscore';
import { Component, HostListener, OnInit } from '@angular/core';
import { ActivatedRoute, NavigationEnd, NavigationStart, ResolveEnd, Router, Scroll } from '@angular/router';
import { UserService } from '../../service/user.service';
import { Observable, Subject, Subscription, catchError, filter, takeUntil, throwError } from 'rxjs';
import { ConfigService } from '../../service/config.service';
import { FullUser } from '../../data/user-full';
import { EMPLYEE_TYPE } from '../../data/constants';
import { OidcSecurityService } from 'angular-auth-oidc-client';
import { NotificationService } from '../../service/notification.service';
import { HttpErrorResponse } from '@angular/common/http';
import { MailGroupCodFull } from '../../data/mail-group-cod-full';
import { MailGroupGlobalFull } from '../../data/mail-group-global-full';
import { SafeUrl } from '@angular/platform-browser';
import { UserPhotoService } from '../../service/user-photo.service';

@Component({
    selector: 'user-view',
    templateUrl: './user-view.component.html',
    styleUrl: './user-view.component.css',
    standalone: false
})
export class UserViewComponent implements OnInit {
    private readonly _destroying$ = new Subject<void>();
    emplyeeType = EMPLYEE_TYPE;
    userId?: string | null;
    user?: FullUser;
    myGroups?: (MailGroupCodFull | MailGroupGlobalFull)[];
    userFound?: boolean | null = null
    expandDirects: boolean = true;
    expandCoDirects: boolean = true;
    expandContigents: boolean = true;
    isSelfProfile: boolean = false;
    photoModalOpen: boolean = false;
    photoModalUrl: SafeUrl | null = null;
    private navSubscription = new Subscription();
    private checkAuthSubscription = new Subscription();
    private activeRouterSubscription = new Subscription();
    private photoSubscription = new Subscription();
    private photoRequestVersion = 0;

    constructor(private activeRouter: ActivatedRoute, private router: Router, private oidcSecurityService: OidcSecurityService,
        private userService: UserService, private configService: ConfigService, private notificationService: NotificationService,
        private photoSvc: UserPhotoService) {

    }

    ngOnInit(): void {
        this.navSubscription.unsubscribe();
        this.navSubscription = this.router.events.pipe(takeUntil(this._destroying$))
            .subscribe((e: any) => {
                if (e instanceof NavigationStart) this.closePhotoModal();
                // console.log(e)
                this.checkAuthSubscription.unsubscribe();
                this.activeRouterSubscription.unsubscribe();
                if (e.routerEvent instanceof NavigationEnd) {
                    this.configService.activeTabSubject.next(this.activeRouter.snapshot?.url[0].path || 'home');
                    this.loadData();
                }
            });
    }

    loadData() {
        this.activeRouterSubscription = this.activeRouter.params
            .pipe(takeUntil(this._destroying$))
            .subscribe(params => {
                if (this.activeRouter.snapshot.paramMap.get('id')) {
                    this.userId = this.activeRouter.snapshot.paramMap.get('id');
                    if (this.userId) {
                    this.handleUserInfoResponse(this.userService.getUserByEmailOrId(this.userId, true));
                    }
                }
            });

        if (!this.userId) {
            this.checkAuthSubscription = this.oidcSecurityService
                .checkAuth()
                .pipe(takeUntil(this._destroying$))
                .subscribe(({ isAuthenticated, userData }) => {
                    if (isAuthenticated) {
                        let email = userData.email || userData.preferred_username;
                        this.handleUserInfoResponse(this.userService.getLoginUserInfo(email, true));
                    }
                });
        }

    }

    handleUserInfoResponse(userObservable: Observable<FullUser>) {
        userObservable.pipe(
            takeUntil(this._destroying$),
            catchError((error: HttpErrorResponse) => {
                this.closePhotoModal();
                this.user = undefined;
                this.userFound = false;
                return throwError(() => error);
            })
        ).subscribe((user) => {
            if(user){
                this.populateUser(user);
                this.populateMyGroups(user);
            }
        });
    }

    populateUser(user: FullUser) {
        if (this.user?.UserAccount?.MSID?.toLowerCase() !== user.UserAccount.MSID?.toLowerCase()) {
            this.closePhotoModal();
        }
        this.user = user;
        this.userFound = true
        this.isSelfProfile = !!this.userService.loginUser
            && this.userService.loginUser.UserAccount['LoginID'] == this.user.UserAccount['LoginID'];
    }

    populateMyGroups(user: FullUser) {
        this.myGroups = [];
        if (user.MemberOf && user.MemberOf.length > 0) {
            for (let group of user.MemberOf) {
                this.myGroups.push(group)
            }
        }

        if (user.OwnerOf && user.OwnerOf.length > 0) {
            for (let group of user.OwnerOf) {
                let exsitintVal = _.filter(this.myGroups, item => item.MailGroup.Name == group.MailGroup.Name)
                if (!exsitintVal || exsitintVal.length == 0) {
                    this.myGroups.push(group);
                }
            }
        }
        if (user.ModeratorOf && user.ModeratorOf.length > 0) {
            for (let group of user.ModeratorOf) {
                let exsitintVal = _.filter(this.myGroups, item => item.MailGroup.Name == group.MailGroup.Name)
                if (!exsitintVal || exsitintVal.length == 0) {
                    this.myGroups.push(group);
                }
            }
        }

    }

    ngOnDestroy() {
        this._destroying$.next(undefined);
        this._destroying$.complete();
        this.navSubscription.unsubscribe();
        this.checkAuthSubscription.unsubscribe();
        this.activeRouterSubscription.unsubscribe();
        this.closePhotoModal();
    }

    onEditPhotoClick(): void {
        this.notificationService.setNotificationData(
            'alert-info',
            'Edit photo will be available in a later release.',
            ''
        );
    }

    onViewPhotoClick(event?: Event): void {
        event?.preventDefault();
        const msid = this.user?.UserAccount?.MSID;
        if (!msid) {
            return;
        }
        this.photoSubscription.unsubscribe();
        const requestVersion = ++this.photoRequestVersion;
        this.photoSubscription = this.photoSvc.getPhoto(msid, 512)
            .pipe(takeUntil(this._destroying$))
            .subscribe({
                next: (url) => {
                    if (!url || requestVersion !== this.photoRequestVersion
                        || this.user?.UserAccount?.MSID?.toLowerCase() !== msid.toLowerCase()) {
                        return;
                    }
                    if (url !== this.photoModalUrl) {
                        this.photoSvc.retainPhoto(url);
                        this.photoSvc.releasePhoto(this.photoModalUrl);
                        this.photoModalUrl = url;
                    }
                    this.photoModalOpen = true;
                },
                error: () => {
                    // photo errors are already negative-cached / silenced by the service
                }
            });
    }

    closePhotoModal(): void {
        this.photoRequestVersion++;
        this.photoSubscription.unsubscribe();
        this.photoModalOpen = false;
        const previous = this.photoModalUrl;
        this.photoModalUrl = null;
        this.photoSvc.releasePhoto(previous);
    }

    @HostListener('document:keydown.escape')
    onEscapeKey(): void {
        if (this.photoModalOpen) {
            this.closePhotoModal();
        }
    }
}
