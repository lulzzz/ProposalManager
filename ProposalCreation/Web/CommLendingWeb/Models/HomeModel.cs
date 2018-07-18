using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CommLendingWeb.Models
{
	public class HomeModel
	{
		public IEnumerable<ResourceItem> Resources { get; set; }
    }

	public class ResourceItem
	{
		public string Key { get; set; }
		public string Value { get; set; }
	}
}
