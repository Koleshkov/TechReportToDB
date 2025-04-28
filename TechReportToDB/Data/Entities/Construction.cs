using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TechReportToDB.Data.Entities
{
    internal class Construction : BaseEntity
    {
        public string? Section { get; set; }
        public double? BitOutDiam { get; set; }
        public double? CaseInsDiam { get; set; }
        public double? DepthProject { get; set; }
        public double? DepthFact { get; set; }
        public string? Telemetry { get; set; }
    }
}
