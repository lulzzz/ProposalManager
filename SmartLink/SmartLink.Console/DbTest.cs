// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SmartLink.Service;

namespace SmartLink.Console
{
    public class DbTest
    {
        public void TestInsertSourceCategory()
        {
            var dbContext = new SmartlinkDbContext();
            dbContext.SourceCatalogs.Add(new Entity.SourceCatalog() { Name = "First One" });
            dbContext.SaveChanges();

            dbContext.SourceCatalogs.ToList().ForEach(o =>
            {
                System.Console.WriteLine("SourceCatalog Id:{0}\tName:{1}", o.Id, o.Name);
            });
        }
        
    }
}
