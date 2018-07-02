// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore;
using ApplicationCore.Helpers;

namespace WebReact.Api
{
    [Produces("application/json")]
    [Route("api/[controller]")]
    //[Route("api/[controller]/[action]")]
    public class BaseApiController<T> : Controller
    {
        protected readonly ILogger _logger;
        protected readonly AppOptions _appOptions;

        protected BaseApiController(
            ILogger<T> logger,
            IOptions<AppOptions> appOptions)
        {
            Guard.Against.Null(logger, nameof(logger));
            Guard.Against.Null(appOptions, nameof(appOptions));

            _logger = logger;
            _appOptions = appOptions.Value;
        }
    }
}
