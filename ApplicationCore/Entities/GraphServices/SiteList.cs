// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace ApplicationCore.Entities.GraphServices
{
    /// <summary>
    /// SiteList object used in GraphApi services
    /// </summary>
    public class SiteList : BaseEntity<SiteList>
    {
        public SiteList()
        {
            ListId = String.Empty;
            SiteId = String.Empty;
        }

        /// <summary>
        /// List id of the SharePoint list
        /// </summary>
        [JsonProperty("listId")]
        public string ListId { get; set; }

        /// <summary>
        /// Site id where the SharePoint list is
        /// </summary>
        [JsonProperty("siteId")]
        public string SiteId { get; set; }

        public new string Id { get { return "sid_" + SiteId + "_lid_" + ListId; } }

        public static SiteList Empty =
            new SiteList() { ListId = String.Empty, SiteId = String.Empty };
    }
}
