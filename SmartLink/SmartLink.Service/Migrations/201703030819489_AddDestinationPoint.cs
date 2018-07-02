// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddDestinationPoint : DbMigration
    {
        public override void Up()
        {
            DropForeignKey("dbo.DestinationPoints", "SourcePoint_Id", "dbo.SourcePoints");
            DropIndex("dbo.DestinationPoints", new[] { "SourcePoint_Id" });
            RenameColumn(table: "dbo.DestinationPoints", name: "SourcePoint_Id", newName: "SourcePointId");
            CreateTable(
                "dbo.DestinationCatalogs",
                c => new
                    {
                        Id = c.Guid(nullable: false, identity: true),
                        Name = c.String(maxLength: 255),
                    })
                .PrimaryKey(t => t.Id);
            
            AddColumn("dbo.DestinationPoints", "CatalogId", c => c.Guid(nullable: false));
            AlterColumn("dbo.DestinationPoints", "SourcePointId", c => c.Guid(nullable: false));
            CreateIndex("dbo.DestinationPoints", "SourcePointId");
            CreateIndex("dbo.DestinationPoints", "CatalogId");
            AddForeignKey("dbo.DestinationPoints", "CatalogId", "dbo.DestinationCatalogs", "Id", cascadeDelete: true);
            AddForeignKey("dbo.DestinationPoints", "SourcePointId", "dbo.SourcePoints", "Id", cascadeDelete: true);
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.DestinationPoints", "SourcePointId", "dbo.SourcePoints");
            DropForeignKey("dbo.DestinationPoints", "CatalogId", "dbo.DestinationCatalogs");
            DropIndex("dbo.DestinationPoints", new[] { "CatalogId" });
            DropIndex("dbo.DestinationPoints", new[] { "SourcePointId" });
            AlterColumn("dbo.DestinationPoints", "SourcePointId", c => c.Guid());
            DropColumn("dbo.DestinationPoints", "CatalogId");
            DropTable("dbo.DestinationCatalogs");
            RenameColumn(table: "dbo.DestinationPoints", name: "SourcePointId", newName: "SourcePoint_Id");
            CreateIndex("dbo.DestinationPoints", "SourcePoint_Id");
            AddForeignKey("dbo.DestinationPoints", "SourcePoint_Id", "dbo.SourcePoints", "Id");
        }
    }
}
