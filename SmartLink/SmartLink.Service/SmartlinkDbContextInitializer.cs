// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using SmartLink.Entity;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Migrations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public class SmartlinkDbContextInitializer : CreateDatabaseIfNotExists<SmartlinkDbContext>
    {
        protected override void Seed(SmartlinkDbContext context)
        {
            base.Seed(context);

            context.SourcePointGroups.AddOrUpdate(x => x.Id,
                    new SourcePointGroup() { Id = 1, Name = "FY 2017" },
                    new SourcePointGroup() { Id = 2, Name = "FY 2016" },
                    new SourcePointGroup() { Id = 3, Name = "FY 2015" },
                    new SourcePointGroup() { Id = 4, Name = "2017 Annual Fiscal Year Report" },
                    new SourcePointGroup() { Id = 5, Name = "Q4 2016 Quarterly Enarings Report" },
                    new SourcePointGroup() { Id = 6, Name = "Q3 2016 Quarterly Enarings Report" }
        );

        }
    }
}
