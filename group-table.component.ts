import { Component, EventEmitter, Input, Output, inject } from '@angular/core';
import { CODMailGroup } from '../../../data/mail-group-cod';
import { GlobalMailGroup } from '../../../data/mail-group-global';

import { FormControl } from '@angular/forms';
import * as _ from 'underscore';

import { FilterUsersPipe } from '../../../pipes/filter-users.pipe';
import { LocalStorageService } from '../../../service/local-storage.service';
import { Router } from '@angular/router';
import { Subject, Subscription } from 'rxjs';
import { MailGroupCodFull } from '../../../data/mail-group-cod-full';
import { MailGroupGlobalFull } from '../../../data/mail-group-global-full';
import { BLANK_REPLACEMENT } from '../../../data/constants';
import { DownloadExcelService } from '../../../service/download-excel.service';
const pickedColumns = ['Name', 'Description', 'MembersCount', 'MembershipAdmin', 'InfoSensitivity'];

const exportHeaders = ['Name', 'Description', 'OwnersCount', 'MembersCount', 'Restricted', 'ACL', 'MembershipAdmin', 'InfoSensitivity', 'IsInternetAccessible', 'IsModerated'];

@Component({
    selector: 'group-table',
    templateUrl: './group-table.component.html',
    styleUrl: './group-table.component.css',
    standalone: false
})
export class GroupTableComponent {
  private static nextActionMenuId = 0;
  readonly actionMenuId = `group-table-actions-${GroupTableComponent.nextActionMenuId++}`;
  private readonly _destroying$ = new Subject<void>();
  private navSubscription = new Subscription();
  BLANK_REPLACEMENT: string = BLANK_REPLACEMENT
  page = 1;
  pageSize = 15;
  @Output() selectedGroup: EventEmitter<MailGroupCodFull | MailGroupGlobalFull | undefined> = new EventEmitter<MailGroupCodFull | MailGroupGlobalFull | undefined>();
  _groupArray!: (MailGroupCodFull | MailGroupGlobalFull)[];

  _selectedIndex?: number | undefined = undefined;
  selectedRow?: number | undefined = undefined;
  _dispalyCount?: number;
  filterName = new FormControl('', { nonNullable: true });
  filters: {
    [key: string]: any;
    Name: string,
    Description: string,
    MyRole: string,
    MembershipAdmin: string,
    Restricted: string
  } = {
    Name: '',
    Description: '',
    MyRole: '',
    MembershipAdmin: '',
    Restricted: ''
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
  showGroupDetailInNewPage?: boolean = false

  @Input()
  hideColumns?: { [key: string]: boolean };
  @Input()
  set selectedIndex(value: number | undefined) {
    if (value != undefined) {
      if (this.groupArray && this.groupArray.length > 0) {
        let len = this.groupArray?.length;
        this._selectedIndex = value >= 0 ? (value < len ? value : value - len) : value + len;
        this.selectedGroup.emit(this.groupArray[this._selectedIndex]);
      }
    } else {
      this.selectedGroup.emit(undefined);
    }
  }

  get selectedIndex(): number | undefined {
    return this._selectedIndex;
  }
  @Input()
  set groupArray(value: (MailGroupCodFull | MailGroupGlobalFull)[]) {
    this._groupArray = value;
  }

  get groupArray(): (MailGroupCodFull | MailGroupGlobalFull)[] {
    return this._groupArray;
  }

  get filteredGroups(): (MailGroupCodFull | MailGroupGlobalFull)[] {
    return this._groupArray || [];
  }

  @Input()
  set dispalyCount(value: number) {
    this._dispalyCount = value;
  }

  get dispalyCount(): number | undefined {
    return this._dispalyCount;
  }

  ngOnInit(): void {
    this.resetFilters();

    this.navSubscription = this.router.events.subscribe((e: any) => {
      this.resetFilters();
    });
  }

  constructor(private router: Router, private localStorageService: LocalStorageService, private downloadExcelService: DownloadExcelService) { }

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
  }

  ngOnDestroy() {
    this._destroying$.next(undefined);
    this._destroying$.complete();
    this.navSubscription.unsubscribe();
  }

  resetFilters() {
    this.filters = {
      Name: '',
      Description: '',
      MyRole: '',
      MembershipAdmin: '',
      Restricted: ''
    };
  }

  // const exportHeaders= ['Name', 'Description', 'OwnersCount', 'MembersCount', 'Restricted', 'IsACLEnabled', 'MembershipAdmin'];

  exportGroup() {
    let filteredGroups = this.groupArray; // this.filterUsersPipe.transform(this.groupArray, this.filters);
    if (filteredGroups && filteredGroups.length > 0) {
      let exportGroups: any[] = []
      for (let group of filteredGroups) {
        let groupToExport = _.pick(group.MailGroup, ...pickedColumns);
        groupToExport = _.extend(groupToExport,
          {
            OwnersCount: group.MailGroup.OwnersDN ? group.MailGroup.OwnersDN.length : 0,
            Restricted: group.MailGroup.IsRestricted ? 'Y' : 'N',
            ACL: group.MailGroup.IsACLEnabled ? 'ACL' : 'N',
            IsInternetAccessible: group.MailGroup.IsInternetAccessible ? 'Y' : 'N',
            IsModerated: group.MailGroup.IsModerated ? 'Y' : 'N'
          })
        exportGroups.push(groupToExport);
      }
      this.downloadExcelService.exportToExcel(exportGroups, this.exportFileName ? this.exportFileName : 'groupsResult', exportHeaders);
    }
  }

}
