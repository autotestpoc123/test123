import { jsonObject } from "typedjson";
import { MailGroupCodFull } from "./mail-group-cod-full";
import { MailGroupGlobalFull } from "./mail-group-global-full";
import { UserCod } from "./user-cod";
import { UserGenpop } from "./user-genpop";

@jsonObject
export class QuickSearchResult {
    [key: string]: unknown;
    Users?: (UserCod | UserGenpop)[];
    MailGroups?: (MailGroupCodFull | MailGroupGlobalFull)[];
}
