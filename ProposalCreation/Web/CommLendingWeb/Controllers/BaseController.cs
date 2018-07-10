// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using CommLendingWeb.Helpers;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

namespace CommLendingWeb.Controllers
{
	[Produces("application/json")]
    [Route("api/[controller]/[action]")]
    public abstract class BaseController : Controller
    {
		public BaseController(IConfiguration configuration, IGraphSdkHelper graphSdkHelper)
		{
			this.Configuration = configuration;
			this.GraphHelper = graphSdkHelper;
		}

		protected IConfiguration Configuration
		{
			get;
			private set;
		}

		protected IGraphSdkHelper GraphHelper
		{
			get;
			private set;
		}
	}
}