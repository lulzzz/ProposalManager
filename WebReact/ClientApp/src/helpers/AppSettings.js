/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

// General settings
//export const appUri = 'https://localhost:44385';
export const appUri = 'https://webreact20180403042343.azurewebsites.net'; //Azure onterawe

// Authentication settings
export const clientId = 'b8d26a6b-3bfe-4800-871e-4b2d0a5a157b';
export const redirectUri = appUri + "/";
export const instanceId = 'https://login.microsoftonline.com/';
export const graphScopes = ["offline_access", "profile", "User.ReadBasic.All", "mail.send"];
export const graphScopesAdmin = ["offline_access", "profile", "User.Read.All", "mail.send", "Sites.ReadWrite.All", "Files.ReadWrite.All", "Group.ReadWrite.All"];
export const webApiScopes = ["api://b8d26a6b-3bfe-4800-871e-4b2d0a5a157b/access_as_user"];
export const clientSecret = 'yruUD4880jbryEKIBQ6{){+';
//export const authority = 'https://login.microsoftonline.com/onterawe.onmicrosoft.com'; // Use to override common login and specify authority (tenant) to use eg. https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/authori
export const authority = null; // Null for login as common (multi-tenant also) eg. https://login.microsoftonline.com/common/oauth2/v2.0/authorize
export const teamsAppInstanceId = "04ecf1ad-20fc-4e43-9e7c-176b92be0e44";