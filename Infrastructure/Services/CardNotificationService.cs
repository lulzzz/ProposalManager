// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using ApplicationCore;
using ApplicationCore.Artifacts;
using ApplicationCore.Entities;
using ApplicationCore.Helpers;
using ApplicationCore.Interfaces;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using Microsoft.Bot.Connector.Teams.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Protocols;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Infrastructure.Services
{
    public class CardNotificationService : BaseService<CardNotificationService>
    {
        private readonly MicrosoftAppCredentials _microsoftAppCredentials;
        private readonly IUserContext _userContext;
        private ConnectorClient _connectorClient;

        public CardNotificationService(
            ILogger<CardNotificationService> logger,
            IOptions<AppOptions> appOptions,
            IUserContext userContext) : base(logger, appOptions)
        {
            Guard.Against.Null(userContext, nameof(userContext));

            _userContext = userContext;

            MicrosoftAppCredentials.TrustServiceUrl(_appOptions.BotServiceUrl, DateTime.MaxValue);

            _microsoftAppCredentials = new MicrosoftAppCredentials(
                _appOptions.MicrosoftAppId,
               _appOptions.MicrosoftAppPassword);

            _connectorClient = new ConnectorClient(
                new Uri(_appOptions.BotServiceUrl), _microsoftAppCredentials);
        }
        public async Task<StatusCodes> sendNotificationCardAsync(string opportunityName, string channelId, IList<UserProfile> sendToList, string messageText, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_sendNotificationCardAsync main called.");

            Guard.Against.NullOrEmpty(opportunityName, "sendNotificationCardAsync opportunityName is null or empty", requestId);
            Guard.Against.NullOrEmpty(channelId, "sendNotificationCardAsync channelId is null or empty", requestId);

            var sendList = sendToList.ToList();

            var botAccount = new ChannelAccount(_appOptions.BotId, _appOptions.BotName);
            var userAccount = new ChannelAccount();
            if (sendList.Count > 0)
            {
                userAccount = new ChannelAccount(sendList[0].Fields.UserPrincipalName, sendList[0].DisplayName);
            }
            else
            {
                userAccount = botAccount;
            }

            var channelData = new TeamsChannelData
            {
                Channel = new ChannelInfo(channelId),
                Notification = new NotificationInfo(true)
            };

            IMessageActivity newMessage = Activity.CreateMessageActivity();
            newMessage.Type = ActivityTypes.Message;
            newMessage.Text = messageText;
            newMessage.Summary = $"Opportunity: {opportunityName}";

            //var heroCard = new HeroCard
            //{
            //    Title = "Test title",
            //    Subtitle = "Test Subtitle",
            //    Text = $"Some text"
            //    //Buttons = new List<CardAction> { new CardAction(ActionTypes.OpenUrl, title: Constants.cardButtonTitle, value: notificationRequest.DeepLink) }
            //};

            //newMessage.Attachments.Add(heroCard.ToAttachment());


            var recepientsList = new List<ChannelAccount>();
            foreach (var item in sendList)
            {
                if (!String.IsNullOrEmpty(item.DisplayName))
                {
                    newMessage.AddMentionToText(new ChannelAccount(item.Id, item.DisplayName), MentionTextLocation.AppendText, $"@{item.DisplayName}");
                    recepientsList.Add(new ChannelAccount(item.Fields.UserPrincipalName, item.DisplayName));
                }
            }
            
            if (recepientsList.Count == 0)
            {
                recepientsList = null;
            }
            else
            {
                if (String.IsNullOrEmpty(recepientsList[0].Id))
                {
                    recepientsList = null;
                }
            }

            ConversationParameters conversationParams = new ConversationParameters(
                isGroup: true,
                bot: botAccount,
                //members: null,
                members: recepientsList,
                topicName: "Proposal Manager",
                activity: (Activity)newMessage,
                channelData: channelData);

            var result = await _connectorClient.Conversations.CreateConversationAsync(conversationParams);

            return StatusCodes.Status200OK;
        }

        public async Task<StatusCodes> sendNotificationCardAsync(string opportunityName, string channelId, UserProfile sendTo, string messageText, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_sendNotificationCardAsync name-user called.");

            var sendToList = new List<UserProfile>();
            sendToList.Add(sendTo);

            return await sendNotificationCardAsync(opportunityName, channelId, sendToList, requestId);
        }

        public async Task<StatusCodes> sendNotificationCardAsync(Opportunity opportunity, IList<UserProfile> sendToList, string messageText, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_sendNotificationCardAsync opp-list called.");

            return await sendNotificationCardAsync(opportunity.DisplayName, opportunity.Metadata.OpportunityChannelId, sendToList, messageText, requestId);
        }

        public async Task<StatusCodes> sendNotificationCardAsync(Opportunity opportunity, UserProfile sendTo, string messageText, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_sendNotificationCardAsync opp-user called.");

            var sendToList = new List<UserProfile>();
            sendToList.Add(sendTo);

            return await sendNotificationCardAsync(opportunity.DisplayName, opportunity.Metadata.OpportunityChannelId, sendToList, messageText, requestId);
        }



        /// <summary>
        /// Handles incoming Bot Framework messages.
        /// </summary>
        /// <param name="activity">Incoming request from Bot Framework.</param>
        /// <param name="connectorClient">Connector client instance for posting to Bot Framework.</param>
        /// <returns>HTTP response message.</returns>
        public async Task<HttpResponseMessage> HandleIncomingRequestAsync(Activity activity, ConnectorClient connectorClient, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_HandleIncomingRequestAsync called.");

            try
            {
                var activityJson = JObject.FromObject(activity);
                // return our reply to the user
                if (activity.Type == ActivityTypes.ConversationUpdate)
                {
                    var reply = activity.CreateReply($"Notifications service bot has been enabled for this channel.");
                    await connectorClient.Conversations.ReplyToActivityAsync(reply);
                }
                
                return new HttpResponseMessage(HttpStatusCode.OK);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - CardNotificationService_HandleIncomingRequestAsync error: {ex}");
                throw;
            }

            
            //switch (activity.GetActivityType())
            //{
            //    case ActivityTypes.Message:
            //        await HandleTextMessagesAsync(activity, connectorClient);
            //        break;

            //    case ActivityTypes.ConversationUpdate:
            //        await HandleConversationUpdatesAsync(activity, connectorClient);
            //        break;

            //    case ActivityTypes.Invoke:
            //        return await HandleInvokeAsync(activity, connectorClient);

            //    case ActivityTypes.ContactRelationUpdate:
            //    case ActivityTypes.Typing:
            //    case ActivityTypes.DeleteUserData:
            //    case ActivityTypes.Ping:
            //    default:
            //        break;
            //}

            //return new HttpResponseMessage(HttpStatusCode.OK);
        }



        // TODO: REVIEW
        public async Task<StatusCodes> sendCard2Async(JObject message, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_sendCard2Async called.");


            var channelId = "19:5325c0c7b9724852bb58d2cf6a6038bf@thread.skype";
            var sendTo = UserProfile.Empty;
            sendTo.Fields.UserPrincipalName = "admin@onterawe.onmicrosoft.com";
            sendTo.DisplayName = "Terawe Solutions";

            var channelData = new TeamsChannelData
            {
                Channel = new ChannelInfo(channelId),
                Notification = new NotificationInfo(true)
            };

            //var botAccount = new ChannelAccount(name: "Proposal Manager onterawe", id: "f1336ec9-ee95-4e1a-83cf-43b3571f37e7");
            // PMOnterawe@AqZKOKau4w0
            var botAccount = new ChannelAccount("28:f1336ec9-ee95-4e1a-83cf-43b3571f37e7", "Proposal Manager onterawe");
            var userAccount = new ChannelAccount(sendTo.Fields.UserPrincipalName, sendTo.DisplayName);

            IMessageActivity newMessage = Activity.CreateMessageActivity();
            //newMessage.Type = ActivityTypes.Message;
            newMessage.From = botAccount;
            newMessage.Recipient = userAccount;
            newMessage.Text = $"Bot called with activity: {message}";

            var recepientsList = new List<ChannelAccount>();
            recepientsList.Add(userAccount);

            ConversationParameters conversationParams = new ConversationParameters(
                isGroup: true,
                bot: botAccount,
                members: recepientsList,
                topicName: "Proposal Manager",
                activity: (Activity)newMessage,
                channelData: channelData);

            var result = await _connectorClient.Conversations.CreateConversationAsync(conversationParams);

            return StatusCodes.Status200OK;
        }

        public async Task<StatusCodes> sendCardAsync(string channelId, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_sendCardAsync called.");


            var userAccount = new ChannelAccount(name: "Terawe Solutions", id: "@018ce842-7af3-4453-afb2-3b86468b9de7");
            //var connector = new ConnectorClient(new Uri("https://smba.trafficmanager.net/amer-client-ss.msg/"));
            var connector = _connectorClient;
            var botAccount = new ChannelAccount(name: "PMOnterawe", id: "@f1336ec9-ee95-4e1a-83cf-43b3571f37e7");
            var conversationId = await connector.Conversations.CreateDirectConversationAsync(botAccount, userAccount);

            IMessageActivity message = Activity.CreateMessageActivity();
            message.From = botAccount;
            message.Recipient = userAccount;
            message.Conversation = new ConversationAccount(id: conversationId.Id);
            message.Text = "Hello, Larry!";
            message.Locale = "en-Us";
            await connector.Conversations.SendToConversationAsync((Activity)message);

            return StatusCodes.Status200OK;
        }

        public async Task<StatusCodes> sendDirectCardAsync(Opportunity opportunity, UserProfile sendTo, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_sendDirectCardAsync called.");

            var channelId = "19:5325c0c7b9724852bb58d2cf6a6038bf@thread.skype";

            var channelData = new TeamsChannelData
            {
                Channel = new ChannelInfo(channelId)
                //Notification = new NotificationInfo(true)
            };

            var botAccount = new ChannelAccount("28:f1336ec9-ee95-4e1a-83cf-43b3571f37e7", "Proposal Manager onterawe");
            var userAccount = new ChannelAccount(sendTo.Fields.UserPrincipalName, sendTo.DisplayName);

            //var conversationId = (await _connectorClient.Conversations.CreateDirectConversationAsync(botAccount, userAccount)).Id;
            var conversationId = "None yet";
            var responseCreate = new JObject();
            try
            {
                var respCreateConversation = await _connectorClient.Conversations.CreateDirectConversationAsync(botAccount, userAccount);
                //var respCreateConversation = _connectorClient.Conversations.CreateOrGetDirectConversation(botAccount, userAccount, "9ec16324-e534-4d84-81de-59a03f343e20");
                conversationId = respCreateConversation.Id;
                responseCreate = JObject.FromObject(respCreateConversation);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - CardNotificationService_sendDirectCardAsync CreateDirectConversationAsync error: {ex}");
            }
            // Aki@onterawe.onmicrosoft.com

            var recepientsList = new List<ChannelAccount>();
            recepientsList.Add(userAccount);

            IMessageActivity newMessage = Activity.CreateMessageActivity();
            newMessage.Type = ActivityTypes.Message;
            newMessage.From = botAccount;
            newMessage.Recipient = userAccount;
            newMessage.ChannelId = "msteams";
            newMessage.Text = $"New opportunity via sendDirectCardAsync CreateConversationAsync conversationId: {conversationId}";
            newMessage.AddMentionToText(userAccount, MentionTextLocation.AppendText, "@Proposal Manager onterawe");


            ConversationParameters conversationParams = new ConversationParameters(
                isGroup: true,
                bot: botAccount,
                members: recepientsList,
                topicName: "Proposal Manager",
                channelData: channelData,
                activity: (Activity)newMessage);

            var result = await _connectorClient.Conversations.CreateConversationAsync(conversationParams);

            conversationParams.Activity.Text = "Before GetConversationMembersAsync";
            result = await _connectorClient.Conversations.CreateConversationAsync(conversationParams);

            var members = await _connectorClient.Conversations.GetConversationMembersAsync(result.ActivityId);

            //newMessage.Conversation = new ConversationAccount(id: result.ActivityId);
            newMessage.Text = $"New opportunity via sendDirectCardAsync SendToConversationAsync {members}";
            conversationParams.Activity = (Activity)newMessage;
            result = await _connectorClient.Conversations.CreateConversationAsync(conversationParams);

            //var result2 = await _connectorClient.Conversations.SendToConversationAsync((Activity)newMessage);

            //var result = await _connectorClient.Conversations.CreateDirectConversationAsync(botAccount, userAccount, newMessage);

            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_sendDirectCardAsync result.Id {result.Id}");
            _logger.LogInformation($"RequestId: {requestId} - CardNotificationService_sendDirectCardAsync result.ServiceUrl {result.ServiceUrl}");

            return StatusCodes.Status200OK;
        }

        /// <summary>
        /// Handles text message input sent by user.
        /// </summary>
        /// <param name="activity">Incoming request from Bot Framework.</param>
        /// <param name="connectorClient">Connector client instance for posting to Bot Framework.</param>
        /// <returns>Task tracking operation.</returns>
        private static async Task HandleTextMessagesAsync(Activity activity, ConnectorClient connectorClient)
        {
            if (activity.Text.Contains("GetChannels"))
            {
                Activity replyActivity = activity.CreateReply();
                replyActivity.AddMentionToText(activity.From, MentionTextLocation.PrependText);

                ConversationList channels = connectorClient.GetTeamsConnectorClient().Teams.FetchChannelList(activity.GetChannelData<TeamsChannelData>().Team.Id);

                // Adding to existing text to ensure @Mention text is not replaced.
                replyActivity.Text = replyActivity.Text + " <p>" + string.Join("</p><p>", channels.Conversations.ToList().Select(info => info.Name + " --> " + info.Id));
                await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
            }
            else if (activity.Text.Contains("GetTenantId"))
            {
                Activity replyActivity = activity.CreateReply();
                replyActivity = replyActivity.AddMentionToText(activity.From, MentionTextLocation.PrependText);

                if (!activity.Conversation.IsGroup.GetValueOrDefault())
                {
                    replyActivity = replyActivity.NotifyUser();
                }

                replyActivity.Text += " Tenant ID - " + activity.GetTenantId();
                await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
            }
            else if (activity.Text.Contains("Create1on1"))
            {
                var response = connectorClient.Conversations.CreateOrGetDirectConversation(activity.Recipient, activity.From, activity.GetTenantId());
                Activity newActivity = new Activity()
                {
                    Text = "Hello",
                    Type = ActivityTypes.Message,
                    Conversation = new ConversationAccount
                    {
                        Id = response.Id
                    },
                };

                await connectorClient.Conversations.SendToConversationAsync(response.Id, newActivity);
            }
            else if (activity.Text.Contains("GetMembers"))
            {
                var response = (await connectorClient.Conversations.GetConversationMembersAsync(activity.Conversation.Id)).AsTeamsChannelAccounts();
                StringBuilder stringBuilder = new StringBuilder();
                Activity replyActivity = activity.CreateReply();
                replyActivity.Text = string.Join("</p><p>", response.ToList().Select(info => info.GivenName + " " + info.Surname + " --> " + info.ObjectId));
                await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
            }
            else if (activity.Text.Contains("TestRetry"))
            {
                for (int i = 0; i < 15; i++)
                {
                    Activity replyActivity = activity.CreateReply();
                    replyActivity.Text = "Message Count " + i;
                    await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
                }
            }
            else if (activity.Text.Contains("O365Card"))
            {
                O365ConnectorCard card = CreateSampleO365ConnectorCard();
                Activity replyActivity = activity.CreateReply();
                replyActivity.Attachments = new List<Attachment>();
                Attachment plAttachment = card.ToAttachment();
                replyActivity.Attachments.Add(plAttachment);
                await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
            }
            else if (activity.Text.Contains("Signin"))
            {
                var userId = activity.From.Id;
                //if (userIdFacebookTokenCache.ContainsKey(userId))
                //{
                //    // Use cached token
                //    var token = userIdFacebookTokenCache[userId];
                //    try
                //    {
                //        // Send a thumbnail card with user's FB profile
                //        var card = await CreateFBProfileCard(token);
                //        Activity replyActivity = activity.CreateReply();
                //        replyActivity.Text = "Cached credential is found. Use cached token to fetch info.";
                //        replyActivity.Attachments.Add(card);
                //        await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
                //    }
                //    catch (Exception)
                //    {
                //        await SendSigninCardAsync(activity, connectorClient);
                //    }
                //}
                //else
                //{
                //    // No token cached: issue a new Signin card
                //    await SendSigninCardAsync(activity, connectorClient);
                //}
            }
            else if (activity.Text.Contains("Signout"))
            {
                var userId = activity.From.Id;
                //userIdFacebookTokenCache.Remove(userId);
                Activity replyActivity = activity.CreateReply();
                replyActivity.Text = "Your credential has been removed.";
                await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
            }
            else if (activity.Text.Contains("GetTeamDetails"))
            {
                if (string.IsNullOrEmpty(activity.GetChannelData<TeamsChannelData>()?.Team?.Id))
                {
                    Activity replyActivity = activity.CreateReply();
                    replyActivity.Text = "This call can only be made from a Team.";
                    await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
                }
                else
                {
                    var teamDetails = await connectorClient.GetTeamsConnectorClient().Teams.FetchTeamDetailsAsync(activity.GetChannelData<TeamsChannelData>().Team.Id);
                    Activity replyActivity = activity.CreateReply();
                    replyActivity.Text = "<p>Team Id " + teamDetails.Id + " </p>" +
                        "<p>Team Name " + teamDetails.Name + " </p>" +
                        "<p>Team AAD Group Id " + teamDetails.AadGroupId + " </p>";
                    await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
                }
            }
            else
            {
                var accountList = connectorClient.Conversations.GetConversationMembers(activity.Conversation.Id);

                Activity replyActivity = activity.CreateReply();
                replyActivity.Text = "Help " +
                    "<p>Type GetChannels to get List of Channels. </p>" +
                    "<p>Type GetTenantId to get Tenant Id </p>" +
                    "<p>Type Create1on1 to create one on one conversation. </p>" +
                    "<p>Type GetMembers to get list of members in a conversation (team or direct conversation). </p>" +
                    "<p>Type TestRetry to get multiple messages from Bot in throttled and retried mechanism. </p>" +
                    "<p>Type O365Card to get a O365 actionable connector card. </p>" +
                    "<p>Type Signin to issue a Signin card to sign in a Facebook app. </p>" +
                    "<p>Type Signout to logout Facebook app and clear cached credentials. </p>" +
                    "<p>Type GetTeamDetails to get details for the current team. </p>";
                replyActivity = replyActivity.AddMentionToText(activity.From);
                await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
            }
        }

        /// <summary>
        /// Issue a new Signin card.
        /// </summary>
        /// <param name="activity">Incoming request from Bot Framework.</param>
        /// <param name="connectorClient">Connector client instance for posting to Bot Framework.</param>
        /// <returns>The returned ResourceResponse</returns>
        private static async Task<ResourceResponse> SendSigninCardAsync(Activity activity, ConnectorClient connectorClient)
        {
            var userId = activity.From.Id;
            //var authUrl = ConfigurationManager.AppSettings["SigninBaseUrl"] + "/auth/start/" + userId;
            var authUrl = "";
            SigninCard card = new SigninCard();
            card.Text = "Sign in Facebook app";
            card.Buttons = new List<CardAction>() { new CardAction("signin", "Login", null, authUrl) };
            Activity replyActivity = activity.CreateReply();
            replyActivity.Attachments = new List<Attachment>();
            Attachment plAttachment = card.ToAttachment();
            replyActivity.Attachments.Add(plAttachment);
            return await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
        }

        /// <summary>
        /// Create a sample O365 connector card.
        /// </summary>
        /// <returns>The result card with actions.</returns>
        private static O365ConnectorCard CreateSampleO365ConnectorCard()
        {
            var actionCard1 = new O365ConnectorCardActionCard(
                O365ConnectorCardActionCard.Type,
                "Multiple Choice",
                "card-1",
                new List<O365ConnectorCardInputBase>
                {
                    new O365ConnectorCardMultichoiceInput(
                        O365ConnectorCardMultichoiceInput.Type,
                        "list-1",
                        true,
                        "Pick multiple options",
                        null,
                        new List<O365ConnectorCardMultichoiceInputChoice>
                        {
                            new O365ConnectorCardMultichoiceInputChoice("Choice 1", "1"),
                            new O365ConnectorCardMultichoiceInputChoice("Choice 2", "2"),
                            new O365ConnectorCardMultichoiceInputChoice("Choice 3", "3")
                        },
                        "expanded",
                        true),
                    new O365ConnectorCardMultichoiceInput(
                        O365ConnectorCardMultichoiceInput.Type,
                        "list-2",
                        true,
                        "Pick multiple options",
                        null,
                        new List<O365ConnectorCardMultichoiceInputChoice>
                        {
                            new O365ConnectorCardMultichoiceInputChoice("Choice 4", "4"),
                            new O365ConnectorCardMultichoiceInputChoice("Choice 5", "5"),
                            new O365ConnectorCardMultichoiceInputChoice("Choice 6", "6")
                        },
                        "compact",
                        true),
                    new O365ConnectorCardMultichoiceInput(
                        O365ConnectorCardMultichoiceInput.Type,
                        "list-3",
                        false,
                        "Pick an option",
                        null,
                        new List<O365ConnectorCardMultichoiceInputChoice>
                        {
                            new O365ConnectorCardMultichoiceInputChoice("Choice a", "a"),
                            new O365ConnectorCardMultichoiceInputChoice("Choice b", "b"),
                            new O365ConnectorCardMultichoiceInputChoice("Choice c", "c")
                        },
                        "expanded",
                        false),
                    new O365ConnectorCardMultichoiceInput(
                        O365ConnectorCardMultichoiceInput.Type,
                        "list-4",
                        false,
                        "Pick an option",
                        null,
                        new List<O365ConnectorCardMultichoiceInputChoice>
                        {
                            new O365ConnectorCardMultichoiceInputChoice("Choice x", "x"),
                            new O365ConnectorCardMultichoiceInputChoice("Choice y", "y"),
                            new O365ConnectorCardMultichoiceInputChoice("Choice z", "z")
                        },
                        "compact",
                        false)
    },
                new List<O365ConnectorCardActionBase>
                {
                    new O365ConnectorCardHttpPOST(
                        O365ConnectorCardHttpPOST.Type,
                        "Send",
                        "card-1-btn-1",
                        @"{""list1"":""{{list-1.value}}"", ""list2"":""{{list-2.value}}"", ""list3"":""{{list-3.value}}"", ""list4"":""{{list-4.value}}""}")
                });

            var actionCard2 = new O365ConnectorCardActionCard(
                O365ConnectorCardActionCard.Type,
                "Text Input",
                "card-2",
                new List<O365ConnectorCardInputBase>
                {
                    new O365ConnectorCardTextInput(
                        O365ConnectorCardTextInput.Type,
                        "text-1",
                        false,
                        "multiline, no maxLength",
                        null,
                        true,
                        null),
                    new O365ConnectorCardTextInput(
                        O365ConnectorCardTextInput.Type,
                        "text-2",
                        false,
                        "single line, no maxLength",
                        null,
                        false,
                        null),
                    new O365ConnectorCardTextInput(
                        O365ConnectorCardTextInput.Type,
                        "text-3",
                        true,
                        "multiline, max len = 10, isRequired",
                        null,
                        true,
                        10),
                    new O365ConnectorCardTextInput(
                        O365ConnectorCardTextInput.Type,
                        "text-4",
                        true,
                        "single line, max len = 10, isRequired",
                        null,
                        false,
                        10)
                },
                new List<O365ConnectorCardActionBase>
                {
                    new O365ConnectorCardHttpPOST(
                        O365ConnectorCardHttpPOST.Type,
                        "Send",
                        "card-2-btn-1",
                        @"{""text1"":""{{text-1.value}}"", ""text2"":""{{text-2.value}}"", ""text3"":""{{text-3.value}}"", ""text4"":""{{text-4.value}}""}")
                });

            var actionCard3 = new O365ConnectorCardActionCard(
                O365ConnectorCardActionCard.Type,
                "Date Input",
                "card-3",
                new List<O365ConnectorCardInputBase>
                {
                    new O365ConnectorCardDateInput(
                        O365ConnectorCardDateInput.Type,
                        "date-1",
                        true,
                        "date with time",
                        null,
                        true),
                    new O365ConnectorCardDateInput(
                        O365ConnectorCardDateInput.Type,
                        "date-2",
                        false,
                        "date only",
                        null,
                        false)
                },
                new List<O365ConnectorCardActionBase>
                {
                    new O365ConnectorCardHttpPOST(
                        O365ConnectorCardHttpPOST.Type,
                        "Send",
                        "card-3-btn-1",
                        @"{""date1"":""{{date-1.value}}"", ""date2"":""{{date-2.value}}""}")
                });

            var section = new O365ConnectorCardSection(
                "**section title**",
                "section text",
                "activity title",
                "activity subtitle",
                "activity text",
                "http://connectorsdemo.azurewebsites.net/images/MSC12_Oscar_002.jpg",
                "avatar",
                true,
                new List<O365ConnectorCardFact>
                {
                    new O365ConnectorCardFact("Fact name 1", "Fact value 1"),
                    new O365ConnectorCardFact("Fact name 2", "Fact value 2"),
                },
                new List<O365ConnectorCardImage>
                {
                    new O365ConnectorCardImage
                    {
                        Image = "http://connectorsdemo.azurewebsites.net/images/MicrosoftSurface_024_Cafe_OH-06315_VS_R1c.jpg",
                        Title = "image 1"
                    },
                    new O365ConnectorCardImage
                    {
                        Image = "http://connectorsdemo.azurewebsites.net/images/WIN12_Scene_01.jpg",
                        Title = "image 2"
                    },
                    new O365ConnectorCardImage
                    {
                        Image = "http://connectorsdemo.azurewebsites.net/images/WIN12_Anthony_02.jpg",
                        Title = "image 3"
                    }
                });

            O365ConnectorCard card = new O365ConnectorCard()
            {
                Summary = "O365 card summary",
                ThemeColor = "#E67A9E",
                Title = "card title",
                Text = "card text",
                Sections = new List<O365ConnectorCardSection> { section },
                PotentialAction = new List<O365ConnectorCardActionBase>
                    {
                        actionCard1,
                        actionCard2,
                        actionCard3,
                        new O365ConnectorCardViewAction(
                            O365ConnectorCardViewAction.Type,
                            "View Action",
                            null,
                            new List<string>
                            {
                                "http://microsoft.com"
                            }),
                        new O365ConnectorCardOpenUri(
                            O365ConnectorCardOpenUri.Type,
                            "Open Uri",
                            "open-uri",
                            new List<O365ConnectorCardOpenUriTarget>
                            {
                                new O365ConnectorCardOpenUriTarget
                                {
                                    Os = "default",
                                    Uri = "http://microsoft.com"
                                },
                                new O365ConnectorCardOpenUriTarget
                                {
                                    Os = "iOS",
                                    Uri = "http://microsoft.com"
                                },
                                new O365ConnectorCardOpenUriTarget
                                {
                                    Os = "android",
                                    Uri = "http://microsoft.com"
                                },
                                new O365ConnectorCardOpenUriTarget
                                {
                                    Os = "windows",
                                    Uri = "http://microsoft.com"
                                }
                            })
                    }
            };

            return card;
        }

        /// <summary>
        /// Handles conversational updates.
        /// </summary>
        /// <param name="activity">Incoming request from Bot Framework.</param>
        /// <param name="connectorClient">Connector client instance for posting to Bot Framework.</param>
        /// <returns>Task tracking operation.</returns>
        private static async Task HandleConversationUpdatesAsync(Activity activity, ConnectorClient connectorClient)
        {
            TeamEventBase eventData = activity.GetConversationUpdateData();

            switch (eventData.EventType)
            {
                case TeamEventType.ChannelCreated:
                    {
                        ChannelCreatedEvent channelCreatedEvent = eventData as ChannelCreatedEvent;

                        Activity newActivity = new Activity
                        {
                            Type = ActivityTypes.Message,
                            ChannelId = "msteams",
                            ServiceUrl = activity.ServiceUrl,
                            From = activity.Recipient,
                            Text = channelCreatedEvent.Channel.Name + " Channel creation complete",
                            ChannelData = new TeamsChannelData
                            {
                                Channel = channelCreatedEvent.Channel,
                                Team = channelCreatedEvent.Team,
                                Tenant = channelCreatedEvent.Tenant
                            },
                        };

                        await connectorClient.Conversations.SendToConversationWithRetriesAsync(newActivity, channelCreatedEvent.Channel.Id);
                        break;
                    }

                case TeamEventType.ChannelDeleted:
                    {
                        ChannelDeletedEvent channelDeletedEvent = eventData as ChannelDeletedEvent;

                        Activity newActivity = activity.CreateReplyToGeneralChannel(channelDeletedEvent.Channel.Name + " Channel deletion complete");

                        await connectorClient.Conversations.SendToConversationWithRetriesAsync(newActivity, activity.GetGeneralChannel().Id);
                        break;
                    }

                case TeamEventType.MembersAdded:
                    {
                        MembersAddedEvent memberAddedEvent = eventData as MembersAddedEvent;

                        Activity newActivity = activity.CreateReplyToGeneralChannel("Members added to team.");

                        await connectorClient.Conversations.SendToConversationWithRetriesAsync(newActivity, activity.GetGeneralChannel().Id);
                        break;
                    }

                case TeamEventType.MembersRemoved:
                    {
                        MembersRemovedEvent memberRemovedEvent = eventData as MembersRemovedEvent;

                        Activity newActivity = activity.CreateReplyToGeneralChannel("Members removed from the team.");

                        await connectorClient.Conversations.SendToConversationWithRetriesAsync(newActivity, activity.GetGeneralChannel().Id);
                        break;
                    }
            }
        }

        /// <summary>
        /// Handles invoke messages.
        /// </summary>
        /// <param name="activity">Incoming request from Bot Framework.</param>
        /// <param name="connectorClient">Connector client instance for posting to Bot Framework.</param>
        /// <returns>Task tracking operation.</returns>
        private static async Task<HttpResponseMessage> HandleInvokeAsync(Activity activity, ConnectorClient connectorClient)
        {
            // Check if the Activity if of type compose extension.
            if (activity.IsComposeExtensionQuery())
            {
                return await HandleComposeExtensionQueryAsync(activity, connectorClient);
            }
            else if (activity.IsO365ConnectorCardActionQuery())
            {
                return await HandleO365ConnectorCardActionQueryAsync(activity, connectorClient);
            }
            else if (activity.IsSigninStateVerificationQuery())
            {
                return await HandleStateVerificationQueryAsync(activity, connectorClient);
            }
            else
            {
                return new HttpResponseMessage(HttpStatusCode.OK);
            }
        }

        /// <summary>
        /// Handles O365 connector card action queries.
        /// </summary>
        /// <param name="activity">Incoming request from Bot Framework.</param>
        /// <param name="connectorClient">Connector client instance for posting to Bot Framework.</param>
        /// <returns>Task tracking operation.</returns>
        private static async Task<HttpResponseMessage> HandleO365ConnectorCardActionQueryAsync(Activity activity, ConnectorClient connectorClient)
        {
            // Get O365 connector card query data.
            O365ConnectorCardActionQuery o365CardQuery = activity.GetO365ConnectorCardActionQueryData();

            Activity replyActivity = activity.CreateReply();
            replyActivity.TextFormat = "xml";
            replyActivity.Text = $@"
                <h2>Thanks, {activity.From.Name}</h2><br/>
                <h3>Your input action ID:</h3><br/>
                <pre>{o365CardQuery.ActionId}</pre><br/>
                <h3>Your input body:</h3><br/>
                <pre>{o365CardQuery.Body}</pre>
            ";
            await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
            return new HttpResponseMessage(HttpStatusCode.OK);
        }

        /// <summary>
        /// Handles state verification query for signin auth flow.
        /// </summary>
        /// <param name="activity">Incoming request from Bot Framework.</param>
        /// <param name="connectorClient">Connector client instance for posting to Bot Framework.</param>
        /// <returns>Task tracking operation.</returns>
        private static async Task<HttpResponseMessage> HandleStateVerificationQueryAsync(Activity activity, ConnectorClient connectorClient)
        {
            SigninStateVerificationQuery stateVerifyQuery = activity.GetSigninStateVerificationQueryData();
            var state = stateVerifyQuery.State;

            // Decrypt state string to get code and original userId
            var botSecret = ""; // ConfigurationManager.AppSettings[MicrosoftAppCredentials.MicrosoftAppPasswordKey];
            var decryptedState = ""; // CipherHelper.Decrypt(state, botSecret);
            var stateObj = JsonConvert.DeserializeObject<JObject>(decryptedState);
            var code = stateObj.GetValue("accessCode").Value<string>();
            var userId = stateObj.GetValue("userId").Value<string>();

            // Verify userId
            var trustableUserId = activity.From.Id;
            if (userId != trustableUserId)
            {
                // Remove invalid user's cached credential (if any)
                //userIdFacebookTokenCache.Remove(userId);

                // Issue a unauthorized message to clients
                Activity replyError = activity.CreateReply();
                replyError.Text = "Unauthorized: User ID verification failed. Please try again.";
                await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyError);
                return new HttpResponseMessage(HttpStatusCode.Unauthorized);
            }
            else
            {
                // Prepare FB OAuth request
                var fbAppId = ""; // ConfigurationManager.AppSettings["SigninFbClientId"];
                var fbOAuthRedirectUrl = ""; // ConfigurationManager.AppSettings["SigninBaseUrl"] + "/auth/callback";
                var fbAppSecret = ""; // ConfigurationManager.AppSettings["SigninFbClientSecret"];
                var fbOAuthTokenUrl = "/v2.10/oauth/access_token";
                var fbOAuthTokenParams = $"?client_id={fbAppId}&redirect_uri={fbOAuthRedirectUrl}&client_secret={fbAppSecret}&code={code}";

                // Use access code to exchange FB token
                HttpResponseMessage fbResponse = await FBGraphHttpClient.Instance.GetAsync(fbOAuthTokenUrl + fbOAuthTokenParams);
                var tokenObj = await fbResponse.Content.ReadAsAsync<JObject>();
                var token = tokenObj.GetValue("access_token").Value<string>();

                // Update cache
                //userIdFacebookTokenCache[userId] = token;

                // Send a thumbnail card with user's FB profile
                var card = await CreateFBProfileCard(token);
                Activity replyActivity = activity.CreateReply();
                replyActivity.Attachments.Add(card);
                await connectorClient.Conversations.ReplyToActivityWithRetriesAsync(replyActivity);
                return new HttpResponseMessage(HttpStatusCode.OK);
            }
        }

        /// <summary>
        /// Perform Facebook graph API to create a thumbnail card with user profile.
        /// </summary>
        /// <param name="token">Access token.</param>
        /// <returns>Attachment of a thumbnail card.</returns>
        private static async Task<Attachment> CreateFBProfileCard(string token)
        {
            // Use FB token to perform graph API to fetch user FB information
            var fbResponse = await PerformFBGraphApi(token, "me", "fields=name,id,email");
            if (fbResponse.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception("Performing FB graph API failed");
            }

            var fbUser = await fbResponse.Content.ReadAsAsync<JObject>();
            var fbUserId = fbUser.GetValue("id").Value<string>();
            var fbUserPic = await PerformFBGraphApi(token, $"{fbUserId}/picture", "height=100");
            var fbUserPicUrl = fbUserPic.RequestMessage.RequestUri.AbsoluteUri;

            // Send a thumbnail card with user's FB profile
            var card = new ThumbnailCard()
            {
                Title = fbUser.GetValue("name").Value<string>(),
                Subtitle = fbUser.GetValue("email").Value<string>(),
                Images = new List<CardImage>() { new CardImage(fbUserPicUrl, fbUserPicUrl, null) }
            };
            return card.ToAttachment();
        }

        /// <summary>
        /// Perform Facebook graph API.
        /// </summary>
        /// <param name="token">Access token.</param>
        /// <param name="endPoint">Endpoint of graph API.</param>
        /// <param name="parameters">Parameter string.</param>
        /// <returns>Json object returned by FB graph.</returns>
        private static async Task<HttpResponseMessage> PerformFBGraphApi(string token, string endPoint, string parameters)
        {
            var fbGraphParams = $"?access_token={token}&" + parameters;
            var fbResponse = await FBGraphHttpClient.Instance.GetAsync(endPoint + fbGraphParams);
            return fbResponse;
        }

        /// <summary>
        /// Handles compose extension queries.
        /// </summary>
        /// <param name="activity">Incoming request from Bot Framework.</param>
        /// <param name="connectorClient">Connector client instance for posting to Bot Framework.</param>
        /// <returns>Task tracking operation.</returns>
        private static async Task<HttpResponseMessage> HandleComposeExtensionQueryAsync(Activity activity, ConnectorClient connectorClient)
        {
            // Get Compose extension query data.
            ComposeExtensionQuery composeExtensionQuery = activity.GetComposeExtensionQueryData();

            // Process data and return the response.
            ComposeExtensionResponse response = new ComposeExtensionResponse
            {
                ComposeExtension = new ComposeExtensionResult
                {
                    Attachments = new List<ComposeExtensionAttachment>
                    {
                        new HeroCard
                        {
                            Buttons = new List<CardAction>
                            {
                                new CardAction
                                {
                                        Image = "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c7/Bing_logo_%282016%29.svg/160px-Bing_logo_%282016%29.svg.png",
                                        Type = ActionTypes.OpenUrl,
                                        Title = "Bing",
                                        Value = "https://www.bing.com"
                                },
                            },
                            Title = "SampleHeroCard",
                            Subtitle = "BingHeroCard",
                            Text = "Bing.com"
                        }.ToAttachment().ToComposeExtensionAttachment()
                    },
                    Type = "result",
                    AttachmentLayout = "list"
                }
            };

            StringContent stringContent = new StringContent(JsonConvert.SerializeObject(response));
            HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK);
            httpResponseMessage.Content = stringContent;
            return httpResponseMessage;
        }

        /// <summary>
        /// Reusable Facebook graph HTTP client
        /// </summary>
        public class FBGraphHttpClient
        {
            /// <summary>
            /// Private instance of singleton
            /// </summary>
            private static HttpClient httpClient;

            /// <summary>
            /// Gets reusable singleton of Facebook graph HTTP client
            /// </summary>
            public static HttpClient Instance
            {
                get
                {
                    if (httpClient == null)
                    {
                        httpClient = new HttpClient();
                        httpClient.BaseAddress = new Uri("https://graph.facebook.com");
                        httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("image/jpeg"));
                    }

                    return httpClient;
                }
            }
        }
    }
}
