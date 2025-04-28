using System.ComponentModel.DataAnnotations.Schema;

namespace TechReportToDB.Data.Entities
{
    internal class Tool : EquipmentBase
    {
        public Guid JobId { get; set; }
        public virtual Job Job { get; set; } = new Job();

        [Column("Типоразмер")]
        public string? Size { get; set; }

        [Column("Максимальный диаметр")]
        public string? MaxSize { get; set; }

        [Column("Класс")]
        public string? ToolClass { get; set; }

        [Column("Тип")]
        public string? ToolType { get; set; }

        [Column("Принадлежность")]
        public string? Owner { get; set; }

        [Column("Наименование")]
        public string? Name { get; set; }

        [Column("Серийный номер")]
        public string? SerialNumber { get; set; }

        [Column("Артикул")]
        public string? Art { get; set; }

        [Column("Паспорт")]
        public string? Pasport { get; set; }

        [Column("Наработка после ТО")]
        public double? CircTimeAfterInspection { get; set; }

        [Column("Общая наработка")]
        public double? CircTime { get; set; }

        [Column("Дата ТО")]
        public string? InspectionDate { get; set; }

        [Column("Дней до ТО")]
        public int? Days { get; set; }

        [Column("Статус")]
        public string? Status { get; set; }

        [Column("Дата прибытия")]
        public string? ArrivalDate { get; set; }

        [Column("От куда")]
        public string? From { get; set; }

        [Column("Дата отправки")]
        public string? DepartureDate { get; set; }

        [Column("Куда")]
        public string? To { get; set; }

        [Column("Цветовое поле")]
        public string? CollorField { get; set; }

        [Column("Коментарий")]
        public string? Comment { get; set; }

        [Column("Резьба верх")]
        public string? TopThread { get; set; }

        [Column("Резьба низ")]
        public string? BottomTread { get; set; }

        [Column("Сборка")]
        public string? Assembly { get; set; }   

        [Column("Остаток на стейве")]
        public Double? Battery { get; set; }

        [NotMapped]
        public string? FilterName => $"{Job.Field} {Job.Pad} {Job.Well} {ToolClass} {Name} {SerialNumber} {Art} {Status}";
    }
}
