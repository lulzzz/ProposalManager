// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using Microsoft.Bot.Connector.Teams.Models;
using Microsoft.AspNetCore.Mvc;
using WebReact.Interfaces;
using ApplicationCore.Helpers;
using ApplicationCore.Artifacts;
using Newtonsoft.Json.Linq;
using WebReact.ViewModels;
using Microsoft.AspNetCore.Authorization;
using System.IO;
using ApplicationCore.Interfaces;
using Microsoft.Rest;
using System.Net.Http;
using Infrastructure.Services;
using System.Net;

namespace WebReact.Api
{
    /// <summary>
    /// Messaging controller.
    /// </summary>
    [Route("api/[controller]")]
    //[TenantFilter]
    public class MessagesController : BaseApiController<MessagesController>
    {
        private readonly MicrosoftAppCredentials _microsoftAppCredentials;
        private ConnectorClient _connectorClient;
        private readonly CardNotificationService _cardNotificationService;
        private readonly IOpportunityService _opportunityService;

        public MessagesController(
            ILogger<MessagesController> logger,
            IOptions<AppOptions> appOptions,
            CardNotificationService cardNotificationService,
            IOpportunityService opportunityService) : base(logger, appOptions)
        {
            Guard.Against.Null(cardNotificationService, "MessagesController_Constructor cardNotificationService is null");
            Guard.Against.Null(opportunityService, "MessagesController_Constructor opportunityService is null");

            _cardNotificationService = cardNotificationService;
            _opportunityService = opportunityService;

            _microsoftAppCredentials = new MicrosoftAppCredentials(
                _appOptions.MicrosoftAppId,
               _appOptions.MicrosoftAppPassword);

            _connectorClient = new ConnectorClient(
                new Uri(_appOptions.BotServiceUrl), _microsoftAppCredentials);
        }

        /// <summary>
        /// Processes Botframework incoming activities.
        /// </summary>
        /// <param name="activity">Bot framework incoming request.</param>
        /// <returns>Ok result.</returns>
        //[Authorize(Roles = "Bot")]
        [HttpPost]
        public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
        {
            var requestId = $"bot_{Guid.NewGuid().ToString()}";
            _logger.LogInformation($"RequestID:{requestId} - MessagesController_Post called.");

            try
            {
                _logger.LogInformation($"RequestID:{requestId} - MessagesController_Post activity.ServiceUrl: {activity.ServiceUrl}");
                _logger.LogInformation($"RequestID:{requestId} - MessagesController_Post activity.ChannelId: {activity.ChannelId}");

                var channelDataJson = JObject.FromObject(activity.ChannelData);
                var activityJson = JObject.FromObject(activity);
                _logger.LogInformation($"RequestID:{requestId} - MessagesController_Post activity.channelData: {activityJson}");
                //var teamsChannelData = (TeamsChannelData)activity.ChannelData;
                _logger.LogInformation($"RequestID:{requestId} - MessagesController_Post activity.channelData: {channelDataJson["team"]["name"].ToString()} - {channelDataJson["team"]["id"].ToString()}");

                var connector = new ConnectorClient(new Uri(activity.ServiceUrl), _microsoftAppCredentials);
                var response = await _cardNotificationService.HandleIncomingRequestAsync(activity, _connectorClient);

                // Update the opportunity with the channelId
                var opportunity = await _opportunityService.GetItemByNameAsync(channelDataJson["team"]["name"].ToString(), false, requestId);
                if (opportunity != null)
                {
                    opportunity.OpportunityChannelId = channelDataJson["team"]["id"].ToString();
                    var updateOpp = await _opportunityService.UpdateItemAsync(opportunity, requestId);
                }

                var resp = new HttpResponseMessage(HttpStatusCode.OK);
                resp.Content = new StringContent($"<html><body>Message received.</body></html>", System.Text.Encoding.UTF8, @"text/html");

                return resp; 
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} MessagesController_Post error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"MessagesController_Post error: {ex} ", requestId);

                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest);
                resp.Content = new StringContent($"<html><body>Error: {errorResponse} </body></html>", System.Text.Encoding.UTF8, @"text/html");

                return resp;
            }
        }
    }
}
