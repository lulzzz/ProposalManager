// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;

using System.IO;
using System.Threading.Tasks;
using System.Net.Http;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;

using ApplicationCore;
using Infrastructure.GraphApi;
using ApplicationCore.Interfaces;
using ApplicationCore.Entities.GraphServices;

namespace Infrastructure.Services
{
    public class LogicAppsService : BaseService<LogicAppsService>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public LogicAppsService(
            ILogger<LogicAppsService> logger,
            IOptions<AppOptions> appOptions) : base(logger, appOptions)
        {
            //TBD
        }

        public void test()
        {
            //var testVar = _appOptions.SharePointSite;
        }
    }
}
