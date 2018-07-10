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
    public class PublishedHistory : BaseEntity
    {
        public Guid SourcePointId { get; set; }
        [ForeignKey("SourcePointId")]
        [JsonIgnore]
        public SourcePoint SourcePoint { get; set; }
        [StringLength(255)]
        public string Name { get; set; }
        public string Position { get; set; }
        public string Value { get; set; }
        [StringLength(255)]
        public string PublishedUser { get; set; }
        public DateTime PublishedDate { get; set; }
    }
}
