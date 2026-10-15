import { TestBed } from '@angular/core/testing';
import { Subject, of } from 'rxjs';
import { UserTableComponent } from './user-table.component';
import { FilterUsersPipe } from '../../../pipes/filter-users.pipe';
import { GroupMembersExportService } from '../../../service/group-members-export.service';
import { NotificationService } from '../../../service/notification.service';

describe('Group table server export', () => {
  let component: UserTableComponent;
  let api: jasmine.SpyObj<GroupMembersExportService>;
  let download: any;
  beforeEach(() => {
    api = jasmine.createSpyObj('GroupMembersExportService', ['export']);
    download = jasmine.createSpyObj('DownloadExcelService', ['exportToExcel', 'saveAsExcelFile']);
    TestBed.configureTestingModule({ providers: [
      { provide: GroupMembersExportService, useValue: api },
      { provide: NotificationService, useValue: { setNotificationData: jasmine.createSpy() } }
    ] });
    component = TestBed.runInInjectionContext(() => new UserTableComponent({} as any, {} as any, download, new FilterUsersPipe()));
    component.userArray = Array.from({ length: 30 }, (_, i) => ({ MSID: `u${i}`, FullName: `User ${i}`, City: i % 2 ? 'London' : 'Paris' }));
    component.groupExport = { groupId: 'test@mock.local', role: 'member', membersCount: 200 };
    api.export.and.returnValue(of(new Blob(['server bytes'])));
  });

  it('exports all via server and saves its bytes', () => {
    component.exportUser();
    expect(api.export).toHaveBeenCalledWith(component.groupExport!, { scope: 'all' });
    expect(download.saveAsExcelFile).toHaveBeenCalled();
    expect(download.exportToExcel).not.toHaveBeenCalled();
  });

  it('exports all matching IDs across pages', () => {
    component.filters.City = 'Paris';
    component.page = 2;
    component.pageSize = 5;
    component.exportUser();
    expect(api.export).toHaveBeenCalledWith(component.groupExport!, {
      scope: 'selected', userIds: Array.from({ length: 15 }, (_, i) => `msid:u${i * 2}`)
    });
    expect(download.exportToExcel).not.toHaveBeenCalled();
  });

  it('blocks groups above 200 even when filtered below 200', () => {
    component.groupExport!.membersCount = 201;
    component.filters.FullName = 'User 1';
    expect(component.canExport).toBeFalse();
    component.exportUser();
    expect(api.export).not.toHaveBeenCalled();
  });

  it('does not treat an empty selection as all', () => {
    component.filters.City = 'Missing';
    component.exportUser();
    expect(api.export).not.toHaveBeenCalled();
  });

  it('prevents duplicate requests and never falls back on failure', () => {
    const pending = new Subject<Blob>();
    api.export.and.returnValue(pending);
    component.exportUser();
    component.exportUser();
    expect(api.export).toHaveBeenCalledTimes(1);
    pending.error(new Error('Failed'));
    expect(component.exporting).toBeFalse();
    expect(download.exportToExcel).not.toHaveBeenCalled();
    expect(download.saveAsExcelFile).not.toHaveBeenCalled();
  });

  it('rejects empty server files', () => {
    api.export.and.returnValue(of(new Blob([])));
    component.exportUser();
    expect(download.saveAsExcelFile).not.toHaveBeenCalled();
    expect(TestBed.inject(NotificationService).setNotificationData).toHaveBeenCalled();
  });

  it('rejects filtered users without an identifier', () => {
    component.userArray = [{ FullName: 'No identifier' }];
    component.filters.FullName = 'No identifier';
    component.exportUser();
    expect(api.export).not.toHaveBeenCalled();
  });

  it('preserves local export for non-group lists', () => {
    component.groupExport = undefined;
    component.exportUser();
    expect(api.export).not.toHaveBeenCalled();
    expect(download.exportToExcel).toHaveBeenCalled();
  });
});
