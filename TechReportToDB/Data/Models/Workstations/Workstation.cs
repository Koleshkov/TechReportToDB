using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechReportToDB.Data.Models.Workstations.Items;

namespace TechReportToDB.Data.Models.Workstations
{
    public class Workstation
    {
        public string? Field { get; set; }
        public string? Pad { get; set; }
        public string? Well { get; set; }
        public string? FieldTeam { get; set; }
        public string? RegNumber { get; set; }
        public string? SerialNubmer { get; set; }
        public string? Type { get; set; }
        public string? Date { get; set; }
        public string? ActiveId { get; set; }
        public string? GenEngineer { get; set; }
        public string? FilePath { get; set; }
        public List<CheckListItem> CheckList { get; set; } = new();
        public List<AvailabilityItem> AvailabilityList { get; set; } = new();
        public List<ESystemItem> ESystemList { get; set; } = new();
        public List<FireExtinguisherItem> FireExtinguisherList { get; set; } = new();
        public List<DocumentationItem> DocumentationList { get; set; } = new();

    }
}
