// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddRecentFiles : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.RecentFiles",
                c => new
                    {
                        Id = c.Guid(nullable: false, identity: true),
                        User = c.String(maxLength: 255),
                        Date = c.DateTime(nullable: false),
                        CatalogId = c.Guid(nullable: false),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.SourceCatalogs", t => t.CatalogId, cascadeDelete: true)
                .Index(t => t.CatalogId);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.RecentFiles", "CatalogId", "dbo.SourceCatalogs");
            DropIndex("dbo.RecentFiles", new[] { "CatalogId" });
            DropTable("dbo.RecentFiles");
        }
    }
}
