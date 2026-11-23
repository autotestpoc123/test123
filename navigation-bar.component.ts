import { Component, OnInit } from '@angular/core';
import { Profile, profile } from '../../../data/profile';
import { switchMap, catchError, takeUntil, tap } from 'rxjs/operators';
import * as _ from 'underscore';
// import * as jQuery from 'jquery';
import { map, Observable, Observer, of, Subject } from 'rxjs';

// import { MsalService, MsalBroadcastService } from '@azure/msal-angular';
import { UserService } from '../../../service/user.service';
import { Router } from '@angular/router';
import { SearchService } from '../../../service/search.service';
import { UserCod } from '../../../data/user-cod';
import { UserGenpop } from '../../../data/user-genpop';
import { OidcSecurityService } from 'angular-auth-oidc-client';
import { ConfigService } from '../../../service/config.service';
import { MailGroupCodFull } from '../../../data/mail-group-cod-full';
import { MailGroupGlobalFull } from '../../../data/mail-group-global-full';

@Component({
  selector: 'navigation-bar',
  templateUrl: './navigation-bar.component.html',
  styleUrl: './navigation-bar.component.scss',
  standalone: false
})
export class NavigationBarComponent implements OnInit {
  searchResultLoading: boolean = false;
  meta!: Profile;
  searchStr: string = '';
  // searchStrCopy:string = '';
  searchResult?: Observable<(UserCod | UserGenpop)[]>;
  searchExResult?: Observable<(UserCod | UserGenpop)[]>;
  userArray: (UserCod | UserGenpop)[] = [];
  groupArray: (MailGroupCodFull | MailGroupGlobalFull)[] = [];
  isQuickSearch: boolean = true;
  noResultsFound: boolean = false;
  selectedIndex: number | undefined = undefined;
  selectedUser?: UserCod | UserGenpop;

  private readonly _destroying$ = new Subject<void>();

  constructor(private router: Router, private userService: UserService, private searchService: SearchService, private configService: ConfigService, private oidcSecurityService: OidcSecurityService) { }

  ngOnInit(): void {
    this.searchStr = '';
    this.meta = profile;
    this.oidcSecurityService
      .checkAuth()
      .pipe(takeUntil(this._destroying$))
      .subscribe(({ isAuthenticated, userData }) => {
        if (isAuthenticated) {
          let mail = userData.email || userData.preferred_username;
          this.userService.getLoginUserInfo(mail, false).subscribe(user => this.meta.loginUser = user);
        }
      });


    this.searchResult = new Observable((observer: Observer<string>) => {
      observer.next(this.searchStr)
    }).pipe(
      switchMap((query: string) => {
        this.userArray.length = 0;
        this.groupArray.length = 0;
        // if (query && query.length > 2) {
        return this.searchService.quickSearchEx(query, true, 'relevance')
          .pipe(
            tap(data => {
              this.userArray = data.Users ?? [];

              if ('MailGroups' in data && data.MailGroups) {
                this.groupArray = [];
                for (let group of data.MailGroups) {
                  this.groupArray.push(group);
                }
              } else { // this is must to prevent result is too large and response processing time to long and page hang
                this.groupArray.length = 0;
              }
            }),
            map((data: any) => {
              let result: any[] = [];
              result = 'Users' in data && data.Users ? result.concat(data.Users) : result
              result = 'MailGroups' in data && data.MailGroups ? result.concat(data.MailGroups) : result

              return result;
            }),
            catchError(err => {
              console.error('search failed:' + err.message);
              return of([]);
            })
          );
        // }
        // return of([]);
      })
    );

    this.configService.activeTabSubject
      .pipe(takeUntil(this._destroying$))
      .subscribe((activeTab) => {
        this.selectedUser = undefined;
        if (activeTab == 'user') {
          this.searchStr = '';
          this.selectedIndex = undefined;
          this.noResultsFound = false;
        }
      });
  }

  ngOnDestroy(): void {
    this._destroying$.next(undefined);
    this._destroying$.complete();
  }

  swithSearchMode() {
    this.isQuickSearch = !this.isQuickSearch;
  }

  noResultsFoundEvt(event: boolean) {
    this.noResultsFound = this.searchStr.length >= 3 ? event : false;
  }

  onSearchKeyDown(event: Event) {
    let keyEvent = event as KeyboardEvent;
    // if(keyEvent.key != 'ArrowDown'){
    //   this.searchStrCopy = jQuery('#searchBox').val();
    // }
    if (keyEvent.key == 'Enter' || keyEvent.key == 'ArrowUp' || keyEvent.key == 'ArrowDown') {
      event.stopPropagation();
      event.preventDefault();
    }
  }

  onSearchKeyUp(event: Event): void {
    let keyEvent = event as KeyboardEvent;
    switch (keyEvent.key) {
      case 'ArrowUp':
        event.stopPropagation();
        event.preventDefault();
        if (this.selectedIndex == undefined) {
          this.selectedIndex = 0;
        } else {
          this.selectedIndex--;
        }
        if (this.selectedIndex < 0) {
          this.selectedIndex = this.userArray.length > 10 ? 9 : (this.userArray.length - 1);
        }
        break;

      case 'ArrowDown':
        event.stopPropagation();
        event.preventDefault();
        if (this.selectedIndex == undefined) {
          this.selectedIndex = 0
        } else {
          this.selectedIndex = (this.selectedIndex + 1) % 10;
        }
        break;

      case 'Tab':
        this.selectedIndex = 0;
        break;

      case 'Enter':
        event.stopPropagation();
        event.preventDefault();
        if (this.selectedUser) {
          let selectedOption = this.selectedUser;
          this.selectedUser = undefined;
          this.selectedIndex = undefined;
          this.router.navigate(["/", "user", selectedOption.Mail ? selectedOption.Mail : selectedOption.MSID],
            {
              queryParams:
                { showPeople: false, showOthers: true }
            });
        } else {
          let filter = this.searchStr;
          this.searchStr = '';
          if (filter) {
            this.router.navigate(["/", "search", filter]);
          }
        }
        break;
      default:
        {
          this.selectedUser = undefined;
          this.selectedIndex = undefined;
        }
    }

  }

  onSelectUser(user: UserCod | UserGenpop | undefined): void {
    this.selectedUser = user;
  }

  onSearchResultLoading(event: boolean): void {
    this.searchResultLoading = event;
  }
}
