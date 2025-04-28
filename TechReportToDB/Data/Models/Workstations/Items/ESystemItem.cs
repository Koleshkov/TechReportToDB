using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TechReportToDB.Data.Models.Workstations.Items
{
    public class ESystemItem
    {
        public string? Name { get; set; }
        public string? QuantityPlan { get; set; }
        public string? QuantityFact { get; set; }
        public string? Status { get; set; }
        public string? Comment { get; set; }
    }
}
