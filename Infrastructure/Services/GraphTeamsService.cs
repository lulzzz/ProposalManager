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

namespace Infrastructure.Services
{
    public class GraphTeamsAppService : GraphTeamsBaseService
    {
        public GraphTeamsAppService(
            ILogger<GraphTeamsAppService> logger, 
            IOptions<AppOptions> appOptions,
            IGraphClientAppContext graphClientContext) : base(logger, appOptions, graphClientContext)
        {
        }
    }

    public class GraphTeamUserService : GraphTeamsBaseService
    {
        public GraphTeamUserService(
            ILogger<GraphTeamUserService> logger,
            IOptions<AppOptions> appOptions,
            IGraphClientUserContext graphClientContext) : base(logger, appOptions, graphClientContext)
        {
        }
    }
}
