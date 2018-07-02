import { IUser, INote } from '.';

export interface IDocument
{
    displayName: string;
    id: string;
    proposalDocument: IContent;
    customer: ICustomer;
    teamMembers: IUser[];
    notes: INote[];
}

export interface IContent
{
   content: {
          "proposalSectionList": ISectionItem[]
        };
}

export interface ISectionItem
{
    typeName: string;
    id: string;
    displayName: string;
    owner: IUser;
    sectionStatus: number;
    task: string;
    assignedTo: IUser;
    lastModifiedDateTime: string;
    subSectionId: string;
}

export interface ICustomer
{
    id: string;
    reference: string;
    displayName: string;
}