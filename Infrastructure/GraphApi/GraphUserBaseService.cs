// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using Infrastructure.Services;
using ApplicationCore;
using ApplicationCore.Interfaces;
using ApplicationCore.Entities.GraphServices;
using ApplicationCore.Helpers;
using ApplicationCore.Helpers.Exceptions;

namespace Infrastructure.GraphApi
{
    public abstract class GraphUserBaseService : BaseService<GraphUserBaseService>
    {
        protected readonly IGraphClientContext _graphClientContext;
        protected readonly IHostingEnvironment _hostingEnvironment;

        public GraphUserBaseService(
            ILogger<GraphUserBaseService> logger,
            IOptions<AppOptions> appOptions,
            IGraphClientContext graphClientContext,
            IHostingEnvironment hostingEnvironment) : base(logger, appOptions)
        {
            Guard.Against.Null(graphClientContext, nameof(graphClientContext));
            Guard.Against.Null(hostingEnvironment, nameof(hostingEnvironment));
            _graphClientContext = graphClientContext;
            _hostingEnvironment = hostingEnvironment;
        }

        /// <summary>
        /// Graph Service client
        /// </summary>
        public GraphServiceClient GraphClient => _graphClientContext?.GraphClient;

        public async Task<JObject> GetUserAsync(string Upn, bool memberOf, string requestId = "")
        {
            // GET: https://graph.microsoft.com/v1.0/users/{id | userPrincipalName} 
            // EXAMPLE: https://graph.microsoft.com/v1.0/users/akira@onterawe.onmicrosoft.com

            _logger.LogInformation($"RequestId: {requestId} - GetUserAsync called.");
            try
            {
                Guard.Against.NullOrEmpty(Upn, nameof(Upn), requestId);

                var memberOfOption = String.Empty;
                if (memberOf)
                {
                    memberOfOption = "/memberOf";
                }

                var requestUrl = $"{_appOptions.GraphRequestUrl}users/{Upn}{memberOfOption}";

                // Create the request message and add the content.
                HttpRequestMessage hrm = new HttpRequestMessage(HttpMethod.Get, requestUrl);

                var response = new HttpResponseMessage();

                // Authenticate (add access token) our HttpRequestMessage
                await GraphClient.AuthenticationProvider.AuthenticateRequestAsync(hrm);

                _logger.LogInformation($"RequestId: {requestId} - GetUserAsync call to graph: " + requestUrl);

                // Send the request and get the response.
                response = await GraphClient.HttpProvider.SendAsync(hrm);

                // Get the status response and throw if is not 200.
                Guard.Against.NotStatus200OK(response.StatusCode, nameof(this.GetUserAsync), requestId);

                var content = await response.Content.ReadAsStringAsync();
                JObject responseJObject = JObject.Parse(await response.Content.ReadAsStringAsync());

                _logger.LogInformation($"RequestId: {requestId} - GetUserAsync end.");
                return responseJObject;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - GetUserAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - GetUserAsync Service Exception: {ex}");
            }
        }

        public async Task<JObject> GetGroupAsync(IList<QueryParam> queryOptions, string expand = "", string requestId = "")
        {
            // GET: https://graph.microsoft.com/v1.0/groups/{id} 
            // EXAMPLE: https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName,'Test')

            _logger.LogInformation($"RequestId: {requestId} - GetGroupAsync called.");
            try
            {
                Guard.Against.Null(queryOptions, nameof(queryOptions), requestId);

                if (!String.IsNullOrEmpty(expand))
                {
                    queryOptions.Add(new QueryParam("expand", expand));
                }

                var requestOptions = string.Empty;
                foreach (var item in queryOptions)
                {
                    if (String.IsNullOrEmpty(requestOptions))
                    {
                        requestOptions = $"?${item.Name}={item.Value}";
                    }
                    else
                    {
                        requestOptions = requestOptions + $"&{item.Name}={item.Value}";
                    }
                }

                var requestUrl = $"{_appOptions.GraphRequestUrl}groups{requestOptions}";
                _logger.LogInformation($"RequestId: {requestId} - GetGroupAsync requestUrl: {requestUrl}");

                // Create the request message and add the content.
                HttpRequestMessage hrm = new HttpRequestMessage(HttpMethod.Get, requestUrl);

                var response = new HttpResponseMessage();

                // Authenticate (add access token) our HttpRequestMessage
                await GraphClient.AuthenticationProvider.AuthenticateRequestAsync(hrm);

                // Send the request and get the response.
                response = await GraphClient.HttpProvider.SendAsync(hrm);

                // Get the status response and throw if is not 200.
                Guard.Against.NotStatus200OK(response.StatusCode, nameof(this.GetGroupAsync), requestId);

                var content = await response.Content.ReadAsStringAsync();
                JObject responseJObject = JObject.Parse(await response.Content.ReadAsStringAsync());

                _logger.LogInformation($"RequestId: {requestId} - GetGroupAsync end.");
                return responseJObject;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - GetGroupAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - GetGroupAsync Service Exception: {ex}");
            }
        }

        public async Task<JObject> GetGroupMembersAsync(string groupId, string requestId = "")
        {
            // GET /groups/{id}/members 
            // EXAMPLE: https://graph.microsoft.com/v1.0/groups/b5e658d8-1373-4700-b989-8aa8ec375253/members

            _logger.LogInformation($"RequestId: {requestId} - GetGroupMembersAsync called.");
            try
            {
                Guard.Against.NullOrEmpty(groupId, nameof(groupId), requestId);

                var requestUrl = $"{_appOptions.GraphRequestUrl}groups/{groupId}/members";

                // Create the request message and add the content.
                HttpRequestMessage hrm = new HttpRequestMessage(HttpMethod.Get, requestUrl);

                var response = new HttpResponseMessage();

                // Authenticate (add access token) our HttpRequestMessage
                await GraphClient.AuthenticationProvider.AuthenticateRequestAsync(hrm);

                _logger.LogInformation($"RequestId: {requestId} - GetGroupMembersAsync call to graph: " + requestUrl);

                // Send the request and get the response.
                response = await GraphClient.HttpProvider.SendAsync(hrm);

                // Get the status response and throw if is not 200.
                Guard.Against.NotStatus200OK(response.StatusCode, $"GetGroupMembersAsync_response: {requestUrl}", requestId);

                var content = await response.Content.ReadAsStringAsync();
                JObject responseJObject = JObject.Parse(await response.Content.ReadAsStringAsync());

                _logger.LogInformation($"RequestId: {requestId} - GetGroupMembersAsync end.");
                return responseJObject;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - GetGroupMembersAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - GetGroupMembersAsync Service Exception: {ex}");
            }
        }

        public async Task<JObject> AddGroupMemberAsync(string userId, string groupId, string requestId = "")
        {
            // POST https://graph.microsoft.com/v1.0/groups/{id}/members/$ref 
            // EXAMPLE: https://graph.microsoft.com/v1.0/groups/b5e658d8-1373-4700-b989-8aa8ec375253/members/$ref
            // EXAMPLE BODY: {'@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/03f9aaaa-ffe3-4e32-b1d9-d37ed8f5ccd2'}

            _logger.LogInformation($"RequestId: {requestId} - AddGroupMemberAsync called.");
            try
            {
                Guard.Against.NullOrEmpty(userId, nameof(userId), requestId);
                Guard.Against.NullOrEmpty(groupId, nameof(groupId), requestId);

                // Create Json object for request body
                var requestBody = "{'@odata.id': 'https://graph.microsoft.com/beta/directoryObjects/" + userId + "'}";

                var requestUrl = $"{_appOptions.GraphBetaRequestUrl}groups/{groupId}/members/$ref";


                // Create the request message and add the content.
                HttpRequestMessage hrm = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                hrm.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");

                var response = new HttpResponseMessage();

                // Authenticate (add access token) our HttpRequestMessage
                await GraphClient.AuthenticationProvider.AuthenticateRequestAsync(hrm);

                _logger.LogInformation($"RequestId: {requestId} - AddGroupMemberAsync call to graph: " + requestUrl);

                // Send the request and get the response.
                response = await GraphClient.HttpProvider.SendAsync(hrm);

                // Get the status response and throw if is not 200.
                Guard.Against.NotStatus204NoContent(response.StatusCode, nameof(this.AddGroupMemberAsync), requestId);

                // 204 returns empty, using 204 as return value
                JObject responseJObject = JObject.FromObject(StatusCodes.Status204NoContent);

                _logger.LogInformation($"RequestId: {requestId} - AddGroupMemberAsync end.");
                return responseJObject;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - AddGroupMemberAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - AddGroupMemberAsync Service Exception: {ex}");
            }
        }

        public async Task<JObject> AddGroupOwnerAsync(string userId, string groupId, string requestId = "")
        {
            // POST https://graph.microsoft.com/v1.0/groups/{id}/owners/$ref 
            // EXAMPLE: https://graph.microsoft.com/v1.0/groups/b5e658d8-1373-4700-b989-8aa8ec375253/owners/$ref
            // EXAMPLE BODY: {'@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/03f9aaaa-ffe3-4e32-b1d9-d37ed8f5ccd2'}

            _logger.LogInformation($"RequestId: {requestId} - AddGroupOwnerAsync called.");
            try
            {
                Guard.Against.NullOrEmpty(userId, nameof(userId), requestId);
                Guard.Against.NullOrEmpty(groupId, nameof(groupId), requestId);

                // Create Json object for request body
                var requestBody = "{'@odata.id': 'https://graph.microsoft.com/beta/users/" + userId + "'}";

                var requestUrl = $"{_appOptions.GraphBetaRequestUrl}groups/{groupId}/owners/$ref";


                // Create the request message and add the content.
                HttpRequestMessage hrm = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                hrm.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");

                var response = new HttpResponseMessage();

                // Authenticate (add access token) our HttpRequestMessage
                await GraphClient.AuthenticationProvider.AuthenticateRequestAsync(hrm);

                _logger.LogInformation($"RequestId: {requestId} - AddGroupOwnerAsync call to graph: " + requestUrl);

                // Send the request and get the response.
                response = await GraphClient.HttpProvider.SendAsync(hrm);

                // Get the status response and throw if is not 200.
                Guard.Against.NotStatus204NoContent(response.StatusCode, nameof(this.AddGroupMemberAsync), requestId);

                // 204 returns empty, using 204 as return value
                JObject responseJObject = JObject.FromObject(StatusCodes.Status204NoContent);

                _logger.LogInformation($"RequestId: {requestId} - AddGroupOwnerAsync end.");
                return responseJObject;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - AddGroupOwnerAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - AddGroupOwnerAsync Service Exception: {ex}");
            }
        }


        // email and misc related to user
        public async Task<JObject> SendEmail(string userUpn, string messageJson, string requestId = "")
        {
            // POST: https://graph.microsoft.com/v1.0/users/{ id | userPrincipalName}/sendMail
            // EXAMPLE: 

            _logger.LogInformation($"RequestId: {requestId} - SendEmail called.");
            
            try
            {
                Guard.Against.NullOrEmpty(messageJson, "SendEmail_messageJson null-empty", requestId);
                Guard.Against.NullOrEmpty(userUpn, "SendEmail_userUpn null-empty", requestId);

                var requestUrl = $"{_appOptions.GraphRequestUrl}users/{userUpn}/sendMail";

                // Create the request message and add the content.
                HttpRequestMessage hrm = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                hrm.Content = new StringContent(messageJson, Encoding.UTF8, "application/json");

                var response = new HttpResponseMessage();

                // Authenticate (add access token) our HttpRequestMessage
                await GraphClient.AuthenticationProvider.AuthenticateRequestAsync(hrm);

                // Send the request and get the response.
                response = await GraphClient.HttpProvider.SendAsync(hrm);

                // Get the status response and throw if is not 202.
                Guard.Against.NotStatus202Accepted(response.StatusCode, "SendEmail-not202", requestId);

                JObject responseJObject = JObject.FromObject(response);

                _logger.LogInformation($"RequestId: {requestId} - SendEmail end.");
                return responseJObject;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - SendEmail Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - SendEmail Service Exception: {ex}");
            }
        }


        // TODO: Reference methods to be deprecated
        public async Task<JObject> GetMyUserInfoAsync()
        {
            try
            {
                // Call to Graph API to get the current user information.
                var graphResponse = await GraphClient.Me.Request().GetAsync() as User;

                // TODO: See if with 1 call we can get everything like picture, manager, etc.

                if (graphResponse == null) throw new ServiceException(new Error { Code = ErrorConstants.Codes.ItemNotFound });

                var responseJObject = JObject.FromObject(graphResponse);

                return responseJObject;
            }
            catch (ServiceException ex)
            {
                switch (ex.Error.Code)
                {
                    case "Request_ResourceNotFound":
                    case "ResourceNotFound":
                    case "ErrorItemNotFound":
                    case "itemNotFound":
                        return new JObject();
                    case "TokenNotFound":
                        //await HttpContext.ChallengeAsync();
                        throw;
                    default:
                        throw;
                }
            }
        }

        public async Task<JObject> GetUserBasicAsync(string userObjectIdentifier)
        {
            try
            {
                // Call to Graph API to get the current user direct reports information.
                var graphResponse = new User();
                graphResponse = await GraphClient.Users[userObjectIdentifier].Request().GetAsync();

                if (graphResponse == null) throw new ServiceException(new Error { Code = ErrorConstants.Codes.ItemNotFound });

                var responseJObject = JObject.FromObject(graphResponse);

                return responseJObject;
            }
            catch (ServiceException ex)
            {
                switch (ex.Error.Code)
                {
                    case "Request_ResourceNotFound":
                    case "ResourceNotFound":
                    case "ErrorItemNotFound":
                    case "itemNotFound":
                        return new JObject();
                    case "TokenNotFound":
                        //await HttpContext.ChallengeAsync();
                        throw;
                    default:
                        throw;
                }
            }
        }

        public async Task<string> GetPictureBase64Async(string userMail)
        {
            try
            {
                if (!String.IsNullOrEmpty(userMail))
                {
                    // Load user's profile picture.
                    var pictureStream = await GraphClient.Users[userMail].Photo.Content.Request().GetAsync();

                    // Copy stream to MemoryStream object so that it can be converted to byte array.
                    var pictureMemoryStream = new MemoryStream();
                    await pictureStream.CopyToAsync(pictureMemoryStream);

                    // Convert stream to byte array.
                    var pictureByteArray = pictureMemoryStream.ToArray();

                    // Convert byte array to base64 string.
                    var pictureBase64 = Convert.ToBase64String(pictureByteArray);

                    return "data:image/jpeg;base64," + pictureBase64;
                }

                return "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==";
            }
            catch (ServiceException ex)
            {
                switch (ex.Error.Code)
                {
                    case "Request_ResourceNotFound":
                    case "ResourceNotFound":
                    case "ErrorItemNotFound":
                    case "itemNotFound":
                    case "ErrorInvalidUser":
                        // If picture not found, return the default image.
                        return "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==";
                    case "TokenNotFound":
                        //await HttpContext.ChallengeAsync();
                        return String.Empty;
                    default:
                        return "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==";
                }
            }
        }

        public async Task SendEmailAsync(string sendTo, string subject, string emailTemplate)
        {
            if (sendTo == null) return;

            var attachments = new MessageAttachmentsCollectionPage();

            try
            {
                // Load user's profile picture.
                var pictureStream = await GraphClient.Me.Photo.Content.Request().GetAsync();

                // Copy stream to MemoryStream object so that it can be converted to byte array.
                var pictureMemoryStream = new MemoryStream();
                await pictureStream.CopyToAsync(pictureMemoryStream);

                // Convert stream to byte array and add as attachment.
                attachments.Add(new FileAttachment
                {
                    ODataType = "#microsoft.graph.fileAttachment",
                    ContentBytes = pictureMemoryStream.ToArray(),
                    ContentType = "image/png",
                    Name = "me.png"
                });
            }
            catch (ServiceException ex)
            {
                switch (ex.Error.Code)
                {
                    case "Request_ResourceNotFound":
                    case "ResourceNotFound":
                    case "ErrorItemNotFound":
                    case "itemNotFound":
                        throw;
                    case "TokenNotFound":
                        //await HttpContext.ChallengeAsync();
                        throw;
                    default:
                        throw;
                }
            }

            // Prepare the recipient list.
            var splitRecipientsString = sendTo.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
            var recipientList = splitRecipientsString.Select(recipient => new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient.Trim()
                }
            }).ToList();

            // Build the email message.
            var email = new Message
            {
                Body = new ItemBody
                {
                    Content = System.IO.File.ReadAllText(_hostingEnvironment.WebRootPath + emailTemplate),
                    ContentType = BodyType.Html,
                },
                Subject = subject,
                ToRecipients = recipientList,
                Attachments = attachments
            };

            await GraphClient.Me.SendMail(email, true).Request().PostAsync();
        }
    }
}
