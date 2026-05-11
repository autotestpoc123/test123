import { FullUser } from '../../data/user-full';
import { EMPLYEE_TYPE } from '../../data/constants';
import { NotificationService } from '../../service/notification.service';
import { UserService } from '../../service/user.service';



export class FullUser {
    [key: string]: any;
    UserAccount!: UserAccount;
    MemberOf?: (MailGroupCodFull | MailGroupGlobalFull)[];
    OwnerOf?: (MailGroupCodFull | MailGroupGlobalFull)[];
    ModeratorOf?: (MailGroupCodFull | MailGroupGlobalFull)[];
    Manager?: UserAccount;
    DirectReports?: UserAccount[];
    CoDirectReports?: UserAccount[];
    Contingents?: UserAccount[];
}
