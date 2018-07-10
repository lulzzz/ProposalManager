// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;

    public partial class AddNameRangeIdToSourcePoint : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.SourcePoints", "NamePosition", c => c.String());
            AlterColumn("dbo.SourcePoints", "NameRangeId", c => c.String(maxLength: 255));
        }

        public override void Down()
        {
            AlterColumn("dbo.SourcePoints", "NamePosition", c => c.String());
            AlterColumn("dbo.SourcePoints", "NameRangeId", c => c.String());
        }
    }
}
