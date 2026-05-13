import { Component, Input, input } from '@angular/core';
import { NgbPopover } from '@ng-bootstrap/ng-bootstrap';
import { FullUser } from '../../data/user-full';
import { EMPLYEE_TYPE } from '../../data/constants';
import { NotificationService } from '../../service/notification.service';
import { UserService } from '../../service/user.service';

@Component({
    selector: 'profile-view',
    templateUrl: './profile-view.component.html',
    styleUrl: './profile-view.component.css',
    standalone: false
})
export class ProfileViewComponent {
    emplyeeType = EMPLYEE_TYPE;
    @Input()
    user!: FullUser;
    isLoginUser?:boolean = false;
    expandDirects: boolean = true;
    expandCoDirects: boolean = true;
    expandContigents: boolean = true;
    private noTextingCloseTimer?: ReturnType<typeof setTimeout>;

    constructor(private notificationService: NotificationService) {

    }

    ngOnInit(): void {
        if (this.user) {
            this.expandDirects = (this.user.DirectReports && this.user.DirectReports.length > 10) ? false : true;
            this.expandCoDirects = (this.user.CoDirectReports && this.user.CoDirectReports.length > 10) ? false : true;
            this.expandContigents = (this.user.Contingents && this.user.Contingents.length > 10) ? false : true;
        }
    }

    clickExpandDirects() {
        this.expandDirects = !this.expandDirects;
    }

    clickExpandCoDirects() {
        this.expandCoDirects = !this.expandCoDirects;
    }

    clickExpandContigents() {
        this.expandContigents = !this.expandContigents;
    }

    openNoTextingPopover(popover: NgbPopover): void {
        this.cancelCloseNoTextingPopover();
        popover.open();
    }

    scheduleCloseNoTextingPopover(popover: NgbPopover): void {
        this.cancelCloseNoTextingPopover();
        this.noTextingCloseTimer = setTimeout(() => popover.close(), 200);
    }

    cancelCloseNoTextingPopover(): void {
        if (this.noTextingCloseTimer) {
            clearTimeout(this.noTextingCloseTimer);
            this.noTextingCloseTimer = undefined;
        }
    }

    closeNoTextingPopover(popover: NgbPopover): void {
        this.cancelCloseNoTextingPopover();
        popover.close();
    }

    public onClipboardCopy(successful: boolean): void {
        this.notificationService.setNotificationData('alert-success', 'Copied to clipboard ' + (successful ? 'successfully' : 'failed'), '');
    }

    ngOnDestroy(): void {
        this.cancelCloseNoTextingPopover();
    }
}
