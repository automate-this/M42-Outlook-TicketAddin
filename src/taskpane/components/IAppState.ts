export interface IAppState {
    [key: string]: IAppState[keyof IAppState];
    sender: string;
    subject: string;
    body: string;
    attachments: any[];
    mailboxItem: any;
    mailboxItemProps: any;
    initiatorUserId: string;
    creatorUserId: string;
    baseurl: string;
    apitoken: string;
    accesstoken: string;
    ticketid: string;
    ticketnumber: string;
    creatingTicket: boolean;
    tickettype: string;
    ticketCategory: string;
    ticketTransform: boolean;
    ticketTransformTo: string;
    ticketRealType: string;
    assignToMe: boolean;
}