// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;

    public partial class AddSourcePointType : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.SourcePoints", "SourceType", c => c.Int(nullable: false, defaultValue: 1));
        }

        public override void Down()
        {
            DropColumn("dbo.SourcePoints", "SourceType");
        }
    }
}
