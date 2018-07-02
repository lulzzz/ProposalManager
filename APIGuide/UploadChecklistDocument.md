

#
# Uploading Checklist Document

Uploading document in one of the checklists for an opportunity.

## Permissions

The following permission is required to call this API.

- User should have the role of &#39;Loan Officer&#39; in UserRoles list in Sharepoint and hence member of the AD group associated with this role.

## HTTP request

> PUT \{applicationUrl}/api/document/UploadFile/[UrlEncode]\{OpportunityName}//ChecklistDocument=\{teamsChannelName},\{checkListItemId}

| **Key** | **Value** |
| --- | --- |
| Authorization | Bearer {token}. Required. |

### Request body

| **Option** | **Key** | **Value** |
| --- | --- | --- |
| form-data | file | {choose file to be uploaded} |

### Response

If successful, this method returns 200 OK response code.

### Example

##### Request

Here is an example of the request.

> PUT \{applicationUrl}/api/document/UploadFile/[UrlEncode]\{OpportunityName}/ChecklistDocument=\{teamsChannelName},\{checkListItemId}

##### Response

If successful, this method returns 200 OK response code.

##### Screenshot from Postman
![alt text](UploadChecklistDocument.png)
