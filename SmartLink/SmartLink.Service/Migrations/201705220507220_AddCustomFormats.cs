// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddCustomFormats : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.CustomFormats",
                c => new
                    {
                        Id = c.Int(nullable: false),
                        Name = c.String(),
                        DisplayName = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.DestinationPointCustomFormats",
                c => new
                    {
                        DestinationPoint_Id = c.Guid(nullable: false),
                        CustomFormat_Id = c.Int(nullable: false),
                    })
                .PrimaryKey(t => new { t.DestinationPoint_Id, t.CustomFormat_Id })
                .ForeignKey("dbo.DestinationPoints", t => t.DestinationPoint_Id, cascadeDelete: true)
                .ForeignKey("dbo.CustomFormats", t => t.CustomFormat_Id, cascadeDelete: true)
                .Index(t => t.DestinationPoint_Id)
                .Index(t => t.CustomFormat_Id);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.DestinationPointCustomFormats", "CustomFormat_Id", "dbo.CustomFormats");
            DropForeignKey("dbo.DestinationPointCustomFormats", "DestinationPoint_Id", "dbo.DestinationPoints");
            DropIndex("dbo.DestinationPointCustomFormats", new[] { "CustomFormat_Id" });
            DropIndex("dbo.DestinationPointCustomFormats", new[] { "DestinationPoint_Id" });
            DropTable("dbo.DestinationPointCustomFormats");
            DropTable("dbo.CustomFormats");
        }
    }
}
