using System.ComponentModel.DataAnnotations.Schema;

namespace TechReportToDB.Data.Entities
{
    internal class Job : Base
    {
        public List<Tool> Tools { get; set; } = new List<Tool>();

        public List<Kit> Kits { get; set; } = new List<Kit>();

        public List<DD> DDs { get; set; } = new List<DD>();

        public List<MWD> MWDs { get; set; } = new List<MWD>();

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

        [Column("Широта")]
        public string? Lat { get; set; }
        [Column("Долгота")]
        public string? Long { get; set; }

        [Column("Путь к файлу")]
        public string? FilePath { get; set; }
        

        [NotMapped]
        public string? Name => $"{Field} {Pad} {Well} ПП{FieldTeam} {Phone} {Type}";
        
    }
}
