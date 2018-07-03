/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

// General settings
export const appUri = '<to be updated>';

// Authentication settings
export const clientId = '<to be updated>';
export const redirectUri = appUri + "/";
export const instanceId = 'https://login.microsoftonline.com/';
export const graphScopes = ["offline_access", "profile", "User.ReadBasic.All", "mail.send"];
export const graphScopesAdmin = ["offline_access", "profile", "User.Read.All", "mail.send", "Sites.ReadWrite.All", "Files.ReadWrite.All", "Group.ReadWrite.All"];
export const webApiScopes = ["api://<to be updated>"];
export const clientSecret = '<to be updated>';
//export const authority = 'https://login.microsoftonline.com/onterawe.onmicrosoft.com'; // Use to override common login and specify authority (tenant) to use eg. https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/authori
export const authority = null; // Null for login as common (multi-tenant also) eg. https://login.microsoftonline.com/common/oauth2/v2.0/authorize
export const teamsAppInstanceId = "<to be updated>";