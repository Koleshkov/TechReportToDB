using System.ComponentModel.DataAnnotations.Schema;

namespace TechReportToDB.Data.Entities
{
    internal class EquipmentBase : BaseEntity
    {
        [Column("Идентификатор актива")]
        public string? ActiveId { get; set; }
    }
}
