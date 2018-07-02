// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace ApplicationCore.Entities
{
    public class Notification : BaseEntity<Notification>
    {
        /// <summary>
        /// Notification title
        /// </summary>
        [JsonProperty("title", Order = 2)]
        public string Title { get; set; }

        /// <summary>
        /// The values for the Notification fields
        /// </summary>
        [JsonProperty("fields", Order = 3)]
        public NotificationFields Fields { get; set; }

        /// <summary>
        /// Represents the empty client. This field is read-only.
        /// </summary>
        public static Notification Empty
        {
            get => new Notification
            {
                Id = String.Empty,
                Title = String.Empty,
                Fields = NotificationFields.Empty
            };
        }
    }

    public class NotificationFields
    {
        /// <summary>
        /// Sent to address
        /// </summary>
        [JsonProperty("sentTo", Order = 1)]
        public string SentTo { get; set; }

        /// <summary>
        /// Sent from address
        /// </summary>
        [JsonProperty("sentFrom", Order = 2)]
        public string SentFrom { get; set; }

        /// <summary>
        /// Sent date
        /// </summary>
        [JsonProperty("sentDate", Order = 3)]
        public DateTimeOffset SentDate { get; set; }

        /// <summary>
        /// Sets the read status
        /// </summary>
        [JsonProperty("isRead", Order = 4)]
        public bool IsRead { get; set; }

        /// <summary>
        /// Message body of the notification
        /// </summary>
        [JsonProperty("message", Order = 5)]
        public string Message { get; set; }

        /// <summary>
        /// Represents the empty client. This field is read-only.
        /// </summary>
        public static NotificationFields Empty
        {
            get => new NotificationFields
            {
                SentTo = String.Empty,
                SentFrom = String.Empty,
                SentDate = DateTimeOffset.MinValue,
                IsRead = false,
                Message = String.Empty
            };
        }
    }
}
