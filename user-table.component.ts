import { Component, EventEmitter, Input, Output, QueryList, ViewChildren, inject, input } from '@angular/core';
import { UserCod } from '../../../data/user-cod';
import { UserGenpop } from '../../../data/user-genpop';
import { Router } from '@angular/router';
import { TypeaheadMatch } from 'ngx-bootstrap/typeahead';
import { FormControl } from '@angular/forms';
import { NgbHighlight } from '@ng-bootstrap/ng-bootstrap';
import { Observable, Subject, Subscription, finalize, takeUntil, tap } from 'rxjs';
import * as _ from 'underscore';
import { BLANK_REPLACEMENT, EmployeeTypeStrMap, TITLE_MAP } from '../../../data/constants';
import { LocalStorageService } from '../../../service/local-storage.service';
import { FilterUsersPipe } from '../../../pipes/filter-users.pipe';
import { FullUser } from '../../../data/user-full';
import { DownloadExcelService } from '../../../service/download-excel.service';
import { GroupExportContext, GroupMembersExportService } from '../../../service/group-members-export.service';
import { NotificationService } from '../../../service/notification.service';

const pickedColumns = ['LoginID', 'FullName', 'Mail', 'Title', 'Department', 'SubDepartment', 'Division', 'City',
  'Building', 'Floor', 'MSID', 'MSFWID', 'CostCenterNumber', 'WorkPhone', 'PersonalMobilePhone'];
const exportHeaders = ['LoginID', 'FullName', 'Mail', 'Title', 'Department', 'SubDepartment', 'Division', 'City',
  'Building', 'Floor', 'MSID', 'MSFWID', 'EmployeeType', 'CostCenterNumber', 'WorkPhone', 'CorpMobilePhone', 'PersonalMobilePhone'];

@Component({
  selector: 'user-table',
  templateUrl: './user-table.component.html',
  styleUrl: './user-table.component.css',
  standalone: false
})
export class UserTableComponent {
  private static nextActionMenuId = 0;
  readonly actionMenuId = `user-table-actions-${UserTableComponent.nextActionMenuId++}`;
  @Input() groupExport?: GroupExportContext;
  exporting = false;
  private readonly groupMembersExportService = inject(GroupMembersExportService);
  private readonly notificationService = inject(NotificationService);

  get canExport(): boolean {
    return this.groupExport
      ? this.groupExport.membersCount <= 200
      : this.filterUsersPipe.transform(this.userArray, this.filters).length <= 200;
  }

  get hasExportRows(): boolean {
    return this.filterUsersPipe.transform(this.userArray, this.filters).length > 0;
  }

  private readonly _destroying$ = new Subject<void>();
  private navSubscription = new Subscription();
  BLANK_REPLACEMENT: string = BLANK_REPLACEMENT
  TITLE_MAP: { name: string; value: string; }[] = TITLE_MAP;
  page = 1;
  pageSize = 15;
  userList!: (UserCod | UserGenpop)[] | undefined;
  @Output() selectedUser: EventEmitter<UserCod | UserGenpop | undefined> = new EventEmitter<UserCod | UserGenpop | undefined>();
  _userArray!: (UserCod | UserGenpop)[];

  _selectedIndex?: number | undefined = undefined;
  selectedRow?: number | undefined = undefined;
  // how many item will be displayed. in search drap down, will hide item over display count
  _displayCount?: number;
  cityList: Set<string> = new Set();;
  buildingList: Set<string> = new Set();

  filterName = new FormControl('', { nonNullable: true });
  filters: {
    [key: string]: any;
    FullName: string,
    LoginID?: string,
    Title: string,
    Division: string,
    City: string,
    Building: string,
    MSID: string,
    CostCenterNumber: string
  } = {
    FullName: '',
    Title: '',
    Division: '',
    City: '',
    Building: '',
    MSID: '',
    CostCenterNumber: '',
  };
  @Input()
  exportFileName?: string;
  @Input()
  tableTitle!: string;
  @Input()
  showCountNum?: boolean = false;
  @Input()
  showCountNumWithUnit?: boolean = false;
  @Input()
  showDownloadIcon?: boolean = false;
  @Input()
  showDownloadLabel?: boolean = false;
  @Input()
  showActionmenu?: boolean = false;
  @Input()
  showHeaders?: boolean = true;
  @Input()
  showFilter?: boolean = false;
  @Input()
  viewType?: string;
  @Input()
  showPagination?: boolean = false

  @Input()
  hideColumns?: { [key: string]: boolean };
  @Input()
  set selectedIndex(value: number | undefined) {
    if (value != undefined) {
      if (this.userArray && this.userArray.length > 0) {
        let len = this.userArray?.length;
        this._selectedIndex = value >= 0 ? (value < len ? value : value - len) : value + len;
        this.selectedUser.emit(this.userArray[this._selectedIndex]);
      }
    } else {
      this.selectedUser.emit(undefined);
    }
  }
  get selectedIndex(): number | undefined {
    return this._selectedIndex;
  }
  @Input()
  set userArray(value: (UserCod | UserGenpop)[]) {
    this._userArray = value;
  }

  get userArray(): (UserCod | UserGenpop)[] {
    return this._userArray;
  }

  @Input()
  set displayCount(value: number) {
    this._displayCount = value;
  }

  get displayCount(): number | undefined {
    return this._displayCount;
  }

  constructor(private router: Router, private localStorageService: LocalStorageService, private downloadExcelService: DownloadExcelService, private filterUsersPipe: FilterUsersPipe) { }

  ngOnInit(): void {
    this.resetFilters();
    this.initFilterOptions();
    if (this.selectedIndex && this.userArray && this.userArray.length > 0) {
      this.selectedUser.emit(this.userArray[this.selectedIndex]);
    }

    this.navSubscription = this.router.events.subscribe((e: any) => {
      this.resetFilters();
      this.initFilterOptions();
    });
  };

  resetFilters() {
    this.filters = {
      FullName: '',
      Title: '',
      Division: '',
      City: '',
      Building: '',
      MSID: '',
      CostCenterNumber: '',
    };
  }

  initFilterOptions(users?: (UserCod | UserGenpop)[]) {
    if (!users) {
      users = this._userArray;
    }
    if (users) {
      _.each(users, user => {
        if (user.City) {
          this.cityList.add(user.City)
        }
        if (user.Building) {
          this.buildingList.add(user.Building)
        }
      });
    }
  }

  onClickUserRow(email: string | undefined, MSID: string | undefined) {
    if (email) {
      this.router.navigate(['/', 'user', email]);
    } else if (MSID) {
      this.router.navigate(['/', 'user', MSID]);
    }
  }

  clearRecentlyViewd(key: string) {
    localStorage.removeItem(key);
    this.localStorageService.clearSubject.next(key)
  }

  onFilterChange() {
    this.page = 1;
    let filteredUsers = this.filterUsersPipe.transform(this.userArray, this.filters);
    if (!filteredUsers) {
      this.cityList.clear()
      this.buildingList.clear()
    } else {
      this.initFilterOptions(filteredUsers)
    }
  }
  ngOnDestroy() {
    this._destroying$.next(undefined);
    this._destroying$.complete();
    this.navSubscription.unsubscribe();
  }

  // PersonalMobilePhone may be a single value or a list; normalize to a string for the tel: link.
  personalPhoneText(user: UserCod | UserGenpop): string {
    const v = user.PersonalMobilePhone;
    return Array.isArray(v) ? v.join(', ') : (v ?? '');
  }

  exportUser() {
    if (!this.canExport || this.exporting) { return; }
    if (this.groupExport) {
      this.exportGroupUsers();
      return;
    }
    let filteredUsers = this.filterUsersPipe.transform(this.userArray, this.filters);
    if (filteredUsers && filteredUsers.length > 0) {
      this.downloadExcelService.exportToExcel(this.exportRows(filteredUsers), this.exportFileName ? this.exportFileName : 'usersResult', exportHeaders);
    }
  }

  private exportRows(filteredUsers: (UserCod | UserGenpop)[]): any[] {
    return filteredUsers.map(user => {
      const row: Record<string, unknown> = {
        ..._.pick(user, ...pickedColumns),
        CorpMobilePhone: user.CorpMobilePhone?.join(', ') || '',
        EmployeeType: user.EmployeeType ? (EmployeeTypeStrMap[user.EmployeeType] || user.EmployeeType) : ''
      };
      // Copy array values now, not after the asynchronous metadata request.
      for (const key of Object.keys(row)) {
        if (Array.isArray(row[key])) { row[key] = (row[key] as unknown[]).join(', '); }
      }
      return row;
    });
  }

  private exportGroupUsers(): void {
    const context = this.groupExport!;
    const users = this.filterUsersPipe.transform(this.userArray, this.filters);
    if (!users.length) { return; }
    const hasFilters = Object.values(this.filters).some(value => !!value);
    // Snapshot rows at click time so edits during the request cannot change this export.
    const rows = this.exportRows(users);
    const fileName = (this.exportFileName || 'group_users').replace(/[\\/:*?"<>|\x00-\x1f]/g, '_');
    this.exporting = true;
    const request: Observable<unknown> = hasFilters
      ? this.groupMembersExportService.getLabelMetadata().pipe(tap(label => {
          this.downloadExcelService.exportToExcel(rows, fileName, exportHeaders, label);
        }))
      : this.groupMembersExportService.export(context).pipe(tap(blob => {
          if (!blob.size) { throw new Error('Empty export file'); }
          this.downloadExcelService.saveAsExcelFile(blob, fileName);
        }));
    request.pipe(
      takeUntil(this._destroying$),
      finalize(() => this.exporting = false)
    ).subscribe({
      // This request opts out of global error notifications to avoid duplicate alerts.
      // Never fall back to an unlabelled file, including on authorization failures.
      error: () => this.notificationService.setNotificationData('alert-danger', 'Excel export failed. Please retry or contact support.', '')
    });
  }
}
