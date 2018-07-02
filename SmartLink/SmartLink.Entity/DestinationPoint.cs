// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Entity
{
    public enum DestinationTypes
    {
        None = 0,
        Point = 1,
        TableCell = 2,
        TableImage = 3,
        Chart = 4
    }

    public class DestinationPoint : BaseEntity
    {
        [ForeignKey("SourcePointId")]
        public SourcePoint ReferencedSourcePoint { get; set; }
        public Guid SourcePointId { get; set; }
        [ForeignKey("CatalogId")]
        public DestinationCatalog Catalog { get; set; }
        [StringLength(255)]
        public string RangeId { get; set; }
        public Guid CatalogId { get; set; }
        [StringLength(255)]
        public string Creator { get; set; }
        public DateTime Created { get; set; }

        [NotMapped]
        [JsonIgnore]
        public bool SerializeCatalog { get; set; } = false;
        public bool ShouldSerializeCatalog()
        {
            return SerializeCatalog;
        }

        public virtual ICollection<CustomFormat> CustomFormats { get; set; }

        public int? DecimalPlace { get; set; } = null;

        public DestinationTypes DestinationType { get; set; }

        public DestinationPoint()
        {
            CustomFormats = new List<CustomFormat>();
        }
    }
}
