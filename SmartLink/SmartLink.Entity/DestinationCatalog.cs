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
    public class DestinationCatalog : BaseEntity
    {
        [NotMapped]
        public string FileName
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(Name))
                    return Name.Substring(Name.LastIndexOf('/') + 1);
                else
                    return Name;
            }
        }

        public string Name { get; set; }

        public string DocumentId { get; set; }

        public bool IsDeleted { get; set; }

        public ICollection<DestinationPoint> DestinationPoints { get; set; }
        public DestinationCatalog()
        {
            DestinationPoints = new List<DestinationPoint>();
        }

        [NotMapped]
        [JsonIgnore]
        public bool SerializeDestinationPoints { get; set; } = true;
        public bool ShouldSerializeDestinationPoints()
        {
            return SerializeDestinationPoints;
        }
    }
}
