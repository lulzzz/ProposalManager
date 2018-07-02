// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Entity
{
    public class RecentFile : BaseEntity
    {
        [StringLength(255)]
        public string User { get; set; }

        public DateTime Date { get; set; }

        public Guid CatalogId { get; set; }

        [ForeignKey("CatalogId")]
        public virtual SourceCatalog Catalog { get; set; }
    }
}
