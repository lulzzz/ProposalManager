export interface IUser
{
    typeName: string;
    displayName: string;
    id: string;
    fields: {
        mail: string;
        userPrincipalName: string;
        title: string;
        userRoles: {}[]
    }
}