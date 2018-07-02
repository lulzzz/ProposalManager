import { IUser } from '.';

export interface INote
{
    id: string;
    noteBody: string;
    createdDateTime: Date;
    createdBy: IUser;
}