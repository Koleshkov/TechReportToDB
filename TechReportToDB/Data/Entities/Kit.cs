using System.ComponentModel.DataAnnotations.Schema;

namespace TechReportToDB.Data.Entities
{
    internal class Kit : EquipmentBase
    {
        public Guid JobId { get; set; }

        public Job Job { get; set; } = new();

        public List<KitTool> KitTools { get; set; } = new();

        [Column("Нименование")]
        public string? Name { get; set; }

        [Column("Акртикул")]
        public string? Art { get; set; }

        [Column("Серийный номер")]
        public string? SerialNumber { get; set; }

        [Column("Норма")]
        public string? QuantityNorm { get; set; }

        [Column("Факт")]
        public string? QuantityFact { get; set; }

        [Column("Статус")]
        public string? Status { get; set; }

        [Column("Комментарий")]
        public string? Comment { get; set; }

        [Column("Дата прибытия")]
        public string? ArrivalDate { get; set; }

        [Column("От куда")]
        public string? From { get; set; }

        [Column("Дата отправки")]
        public string? DepartureDate { get; set; }

        [Column("Куда")]
        public string? To { get; set; }

        [Column("Дата ТО")]
        public string? InspectionDate { get; set; }
    }
}
