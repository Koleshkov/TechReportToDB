using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace TechReportToDB.Data.Entities
{
    internal class Job : BaseEntity
    {
        [JsonIgnore]
        public ICollection<Tool> Tools { get; set; } = new List<Tool>();
        [JsonIgnore]
        public ICollection<Kit> Kits { get; set; } = new List<Kit>();
        [JsonIgnore]
        public ICollection<DD> DDs { get; set; } = new List<DD>();
        [JsonIgnore]
        public ICollection<MWD> MWDs { get; set; } = new List<MWD>();
        [JsonIgnore]
        public ICollection<Construction> Constructions { get; set; } = new List<Construction>();

        [Column("Месторождение")]
        public string? Field { get; set; }
            
        [Column("Куст")]
        public string? Pad { get; set; }

        [Column("Скважина")]
        public string? Well { get; set; }

        [Column("Полевая партия")]
        public string? FieldTeam { get; set; }

        [Column("Номер телефона")]
        public string? Phone { get; set; }

        [Column("Тип скважины")]
        public string? Type { get; set; }

        [Column("Забой")]
        public string? Depth { get; set; }

        [Column("Буровой подрядчик")]
        public string? DrillingContractor { get; set; }
        [JsonIgnore]
        public string Label { get; set; } = "";
        [JsonIgnore]
        public double Latitude { get; set; }
        [JsonIgnore]
        public double Longitude { get; set; }
        [JsonIgnore]

        [Column("Путь к файлу")]
        public string? FilePath { get; set; }

        [JsonIgnore]
        [NotMapped]
        public string? FilterName
        {
            get
            {
                string n = "";
                foreach (var dd in DDs)
                {
                    n=n + " " + dd.Name;
                }

                foreach (var mwd in MWDs)
                {
                    n = n + " " + mwd.Name;
                }

                foreach (var tool in Tools.Where(t=>t.Status=="В КНБК"))
                {
                    n = n + " " + tool.Name;
                }

                foreach (var c in Constructions)
                {
                    n = n + " " + c.Telemetry;
                }

                return $"{Field} {Pad} {Well} ПП{FieldTeam} {Phone} {Type} {n}";
            }
        }

        [JsonIgnore]
        [NotMapped]
        public string? Name
        {
            get
            {
                return $"{Field} {Pad} {Well}";
            }
        }
    }
}
