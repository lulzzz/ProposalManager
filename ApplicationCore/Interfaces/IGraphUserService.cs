// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ApplicationCore.Interfaces
{
    public interface IGraphUserService
    {
        Task<JObject> GetMyManagerAsync();
        Task<JObject> GetMyUserInfoAsync();
        Task<string> GetPictureBase64Async(string userMail);
        Task<JObject> GetUserBasicAsync(string userObjectIdentifier);
        Task SendEmailAsync(string sendTo, string subject, string emailTemplate);
    }
}