// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Newtonsoft.Json;

namespace SmartLink.Entity
{
    public enum SourcePointStatus
    {
        Created = 0,
        Deleted = 1
    }

    public enum SourceTypes
    {
        Point = 1,
        Table = 2,
        Chart = 3
    }

    public class SourcePoint : BaseEntity
    {
        [StringLength(255)]
        public string Name { get; set; }
        [StringLength(255)]
        public string RangeId { get; set; }

        public string Position { get; set; }
        public string Value { get; set; }
        [StringLength(255)]
        public string Creator { get; set; }
        public DateTime Created { get; set; }
        public SourcePointStatus Status { get; set; }
        
        public SourceTypes SourceType { get; set; }

        public string NamePosition { get; set; }
        [StringLength(255)]
        public string NameRangeId { get; set; }

        public Guid CatalogId { get; set; }
        [ForeignKey("CatalogId")]
        public virtual SourceCatalog Catalog { get; set; }
        public virtual ICollection<SourcePointGroup> Groups { get; set; }
        public virtual ICollection<PublishedHistory> PublishedHistories { get; set; }
        [JsonIgnore]
        public virtual ICollection<DestinationPoint> DestinationPoints { get; set; }

        public SourcePoint()
        {
            DestinationPoints = new List<DestinationPoint>();
            PublishedHistories = new List<PublishedHistory>();
            Groups = new List<SourcePointGroup>();
        }

        [NotMapped]
        [JsonIgnore]
        public bool SerializeCatalog { get; set; } = false;
        public bool ShouldSerializeCatalog()
        {
            return SerializeCatalog;
        }
    }
}
