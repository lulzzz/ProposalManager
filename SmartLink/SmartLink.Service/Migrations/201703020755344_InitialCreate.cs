// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialCreate : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.PublishedHistories",
                c => new
                    {
                        Id = c.Guid(nullable: false, identity: true),
                        SourcePointId = c.Guid(),
                        Name = c.String(maxLength: 255),
                        Position = c.String(),
                        Value = c.String(),
                        PublishedUser = c.String(maxLength: 255),
                        PublishedDate = c.DateTime(nullable: false),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.SourcePoints", t => t.SourcePointId)
                .Index(t => t.SourcePointId);
            
            CreateTable(
                "dbo.SourcePoints",
                c => new
                    {
                        Id = c.Guid(nullable: false, identity: true),
                        Name = c.String(maxLength: 255),
                        RangeId = c.String(),
                        Position = c.String(),
                        Value = c.String(),
                        Creator = c.String(maxLength: 255),
                        Created = c.DateTime(nullable: false),
                        Status = c.Int(nullable: false),
                        CatalogId = c.Guid(nullable: false),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.SourceCatalogs", t => t.CatalogId, cascadeDelete: true)
                .Index(t => t.CatalogId);
            
            CreateTable(
                "dbo.SourceCatalogs",
                c => new
                    {
                        Id = c.Guid(nullable: false, identity: true),
                        Name = c.String(maxLength: 255),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.DestinationPoints",
                c => new
                    {
                        Id = c.Guid(nullable: false, identity: true),
                        SourcePoint_Id = c.Guid(),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.SourcePoints", t => t.SourcePoint_Id)
                .Index(t => t.SourcePoint_Id);
            
            CreateTable(
                "dbo.SourcePointGroups",
                c => new
                    {
                        Id = c.Int(nullable: false),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.SourcePointGroupSourcePoints",
                c => new
                    {
                        SourcePointGroup_Id = c.Int(nullable: false),
                        SourcePoint_Id = c.Guid(nullable: false),
                    })
                .PrimaryKey(t => new { t.SourcePointGroup_Id, t.SourcePoint_Id })
                .ForeignKey("dbo.SourcePointGroups", t => t.SourcePointGroup_Id, cascadeDelete: true)
                .ForeignKey("dbo.SourcePoints", t => t.SourcePoint_Id, cascadeDelete: true)
                .Index(t => t.SourcePointGroup_Id)
                .Index(t => t.SourcePoint_Id);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.PublishedHistories", "SourcePointId", "dbo.SourcePoints");
            DropForeignKey("dbo.SourcePointGroupSourcePoints", "SourcePoint_Id", "dbo.SourcePoints");
            DropForeignKey("dbo.SourcePointGroupSourcePoints", "SourcePointGroup_Id", "dbo.SourcePointGroups");
            DropForeignKey("dbo.DestinationPoints", "SourcePoint_Id", "dbo.SourcePoints");
            DropForeignKey("dbo.SourcePoints", "CatalogId", "dbo.SourceCatalogs");
            DropIndex("dbo.SourcePointGroupSourcePoints", new[] { "SourcePoint_Id" });
            DropIndex("dbo.SourcePointGroupSourcePoints", new[] { "SourcePointGroup_Id" });
            DropIndex("dbo.DestinationPoints", new[] { "SourcePoint_Id" });
            DropIndex("dbo.SourcePoints", new[] { "CatalogId" });
            DropIndex("dbo.PublishedHistories", new[] { "SourcePointId" });
            DropTable("dbo.SourcePointGroupSourcePoints");
            DropTable("dbo.SourcePointGroups");
            DropTable("dbo.DestinationPoints");
            DropTable("dbo.SourceCatalogs");
            DropTable("dbo.SourcePoints");
            DropTable("dbo.PublishedHistories");
        }
    }
}
