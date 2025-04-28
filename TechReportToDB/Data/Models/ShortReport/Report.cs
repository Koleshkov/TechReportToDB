using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TechReportToDB.Data.Models.ShortReport
{
    internal class Report
    {
        public string? FieldTeam { get; set; }
        public string? Field { get; set; }
        public string? Pad { get; set; }
        public string? Well { get; set; }
        public string? Type { get; set; }
        public string? Section { get; set; }
        public string? Distance { get; set; }
        public string? Depth { get; set; }
        public string? Comment { get; set; }
    }
}
