// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Text;
using ApplicationCore;
using ApplicationCore.Helpers;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Infrastructure.Services
{
    /// <summary>
    /// Base abstract class for services
    /// </summary>
    public abstract class BaseService<T>
    {
        protected readonly ILogger _logger;
        protected readonly AppOptions _appOptions;

        /// <summary>
        /// Constructor
        /// </summary>
        public BaseService(
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
