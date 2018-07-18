// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using CommLendingWeb.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Localization;
using System.Linq;

namespace CommLendingWeb.Controllers
{
	public class HomeController : Controller
    {
		private readonly IStringLocalizer localizer;
		public HomeController(IStringLocalizer<Resource> localizer)
		{
			this.localizer = localizer;
		}


        public IActionResult Index()
		{
			var model = new HomeModel
			{
				Resources = localizer.GetAllStrings().Select(x => new ResourceItem() { Key = x.Name, Value = System.Web.HttpUtility.JavaScriptStringEncode(x.Value) })
			};

			return View(model);
        }
    }
}
