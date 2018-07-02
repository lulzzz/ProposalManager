// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore;

namespace CommLendingWeb
{
    public class Program
    {
        public static void Main(string[] args)
        {
            BuildWebHost(args).Run();
        }

        public static IWebHost BuildWebHost(string[] args) =>
        WebHost.CreateDefaultBuilder(args)
        .UseStartup<Startup>()
        .Build();
    }
}
