// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;

    public partial class AddCustomFormatGroups : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.CustomFormats", "OrderBy", c => c.Int(nullable: false, defaultValue: 1));
            AddColumn("dbo.CustomFormats", "IsDeleted", c => c.Boolean(nullable: false, defaultValue: false));
            AddColumn("dbo.CustomFormats", "GroupName", c => c.String());
            AddColumn("dbo.CustomFormats", "GroupOrderBy", c => c.Int(nullable: false, defaultValue: 1));
            AddColumn("dbo.DestinationPoints", "DecimalPlace", c => c.Int());
        }

        public override void Down()
        {
            DropColumn("dbo.DestinationPoints", "DecimalPlace");
            DropColumn("dbo.CustomFormats", "GroupOrderBy");
            DropColumn("dbo.CustomFormats", "GroupName");
            DropColumn("dbo.CustomFormats", "IsDeleted");
            DropColumn("dbo.CustomFormats", "OrderBy");
        }
    }
}
