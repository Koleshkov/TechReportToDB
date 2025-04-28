using System.ComponentModel.DataAnnotations.Schema;

namespace TechReportToDB.Data.Entities
{
    internal class Job : BaseEntity
    {
        public ICollection<Tool> Tools { get; set; } = new List<Tool>();

        public ICollection<Kit> Kits { get; set; } = new List<Kit>();

        public ICollection<DD> DDs { get; set; } = new List<DD>();

        public ICollection<MWD> MWDs { get; set; } = new List<MWD>();

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

        [Column("Буровой подрядчик")]
        public string? DrillingContractor { get; set; }

        public string Label { get; set; } = "";
        public double Latitude { get; set; }
        public double Longitude { get; set; }


        [Column("Путь к файлу")]
        public string? FilePath { get; set; }


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
