// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Microsoft.AspNetCore.Mvc;

namespace CommLendingWeb.Controllers
{
	public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
