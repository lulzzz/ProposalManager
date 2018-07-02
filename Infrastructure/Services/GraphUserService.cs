// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ApplicationCore;
using ApplicationCore.Interfaces;
using Infrastructure.GraphApi;
using Microsoft.AspNetCore.Hosting;

namespace Infrastructure.Services
{
    public class GraphUserAppService : GraphUserBaseService
    {
        public GraphUserAppService(
            ILogger<GraphUserAppService> logger, 
            IOptions<AppOptions> appOptions,
            IGraphClientAppContext graphClientContext, 
            IHostingEnvironment hostingEnvironment) : base(logger, appOptions, graphClientContext, hostingEnvironment)
        {
        }
    }

    public class GraphUserUserService : GraphUserBaseService
    {
        public GraphUserUserService(
            ILogger<GraphUserUserService> logger,
            IOptions<AppOptions> appOptions,
            IGraphClientUserContext graphClientContext,
            IHostingEnvironment hostingEnvironment) : base(logger, appOptions, graphClientContext, hostingEnvironment)
        {
        }
    }
}
