// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddRangeIdToDestination : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.DestinationPoints", "RangeId", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.DestinationPoints", "RangeId");
        }
    }
}
