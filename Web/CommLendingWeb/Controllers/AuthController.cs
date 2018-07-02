// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using CommLendingWeb.Extensions;
using CommLendingWeb.Helpers;
using CommLendingWeb.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

namespace CommLendingWeb.Controllers
{
	public class AuthController : Controller
	{
		private readonly AzureAdOptions azureAdOptions;

		public AuthController(IConfiguration configuration)
		{
			// Get from config
			azureAdOptions = new AzureAdOptions();
			configuration.Bind("AzureAd", azureAdOptions);
			
		}
		public IActionResult Index()
		{
			var model = new AuthModel() { ApplicationId = azureAdOptions.ClientId };
			return View(model);
		}

		public IActionResult End()
		{
			return View();
		}

	}
}