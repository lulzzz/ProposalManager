// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class updateSourcePointIdNotNullable : DbMigration
    {
        public override void Up()
        {
            DropForeignKey("dbo.PublishedHistories", "SourcePointId", "dbo.SourcePoints");
            DropIndex("dbo.PublishedHistories", new[] { "SourcePointId" });
            AlterColumn("dbo.PublishedHistories", "SourcePointId", c => c.Guid(nullable: false));
            CreateIndex("dbo.PublishedHistories", "SourcePointId");
            AddForeignKey("dbo.PublishedHistories", "SourcePointId", "dbo.SourcePoints", "Id", cascadeDelete: true);
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.PublishedHistories", "SourcePointId", "dbo.SourcePoints");
            DropIndex("dbo.PublishedHistories", new[] { "SourcePointId" });
            AlterColumn("dbo.PublishedHistories", "SourcePointId", c => c.Guid());
            CreateIndex("dbo.PublishedHistories", "SourcePointId");
            AddForeignKey("dbo.PublishedHistories", "SourcePointId", "dbo.SourcePoints", "Id");
        }
    }
}
