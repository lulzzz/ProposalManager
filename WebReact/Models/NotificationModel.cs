// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore.Entities;
using Newtonsoft.Json;

namespace WebReact.Models
{
    public class NotificationModel
    {
        /// <summary>
        /// Notification identifier
        /// </summary>
        /// <value>Unique ID to identify the model data</value>
        [JsonProperty("id", Order = 1)]
        public string Id { get; set; }

        /// <summary>
        /// Notification title
        /// </summary>
        [JsonProperty("title", Order = 2)]
        public string Title { get; set; }

        /// <summary>
        /// Sent to address
        /// </summary>
        [JsonProperty("sentTo", Order = 3)]
        public string SentTo { get; set; }

        /// <summary>
        /// Sent from address
        /// </summary>
        [JsonProperty("sentFrom", Order = 4)]
        public string SentFrom { get; set; }

        /// <summary>
        /// Sent date
        /// </summary>
        [JsonProperty("sentDate", Order = 5)]
        public DateTimeOffset SentDate { get; set; }

        /// <summary>
        /// Sets the read status
        /// </summary>
        [JsonProperty("isRead", Order = 6)]
        public bool IsRead { get; set; }

        /// <summary>
        /// Message body of the notification
        /// </summary>
        [JsonProperty("message", Order = 7)]
        public string Message { get; set; }
    }
}
