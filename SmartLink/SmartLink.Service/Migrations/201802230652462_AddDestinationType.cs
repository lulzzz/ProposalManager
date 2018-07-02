// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;

    public partial class AddDestinationType : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.DestinationPoints", "DestinationType", c => c.Int(nullable: true, defaultValue: 0));
        }

        public override void Down()
        {
            DropColumn("dbo.DestinationPoints", "DestinationType");
        }
    }
}
