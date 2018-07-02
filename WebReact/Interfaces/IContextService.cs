// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WebReact.Interfaces
{
    public interface IContextService
    {
        Task<JObject> GetTeamGroupDriveAsync(string teamGroupName);

        Task<JObject> GetSiteDriveAsync(string siteName);

        Task<JObject> GetSiteIdAsync(string siteName);
    }
}