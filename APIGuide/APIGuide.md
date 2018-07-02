

## Guide for using the Opportunity API

This guide is created for consuming the rest API using the Postman tool.

### Sections:
[Create Opportunity](CreateOpportunity.md)
<br>
[Get Opportunity](GetOpportunity.md)
<br>
[Update Opportunity](UpdateOpportunity.md)
<br>
[Upload Checklist Document](UploadChecklistDocument.md)
<br>
[Upload Proposal Document](UploadProposalDocument.md)
<br>
[Delete Opportunity](DeleteOpportunity.md)


### Important Points:

 Postman is installed on the machine

- Download link: [https://www.getpostman.com/apps](https://www.getpostman.com/apps)

- Sending API requests from postman: [https://www.getpostman.com/docs/v6/postman/sending\_api\_requests/requests](https://www.getpostman.com/docs/v6/postman/sending_api_requests/requests)



You have to login in to the proposal manager application with the valid user who is authorized to carry out the required transaction. This is required for grabbing the &quot;webApi token&quot; to be used while invoking the API. Token validity may be for a short duration. If you get a 401 unauthorized response, please try to sign-in to the application again and fetch a fresh token.  Below screenshots are for fetching the token from different browsers:

 Edge

![alt text](tokenfromedge.png)


Chrome

 ![alt text](tokenfromchrome.png)

Firefox

 ![alt text](tokenfromfirefox.png)

- Your machine allows calling the api(hosted on azure) , no firewall or proxy blocking the same.
- Url used is  \{hostedAppUrl}/api/opportunity. This same url will be used for all CRUD operations.
[Sample opportunity json object](SampleOpportunityJson.txt)


