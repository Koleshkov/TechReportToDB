using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TechReportToDB.Data.Entities
{
    internal class EquipmentBase : BaseEntity
    {
        [Column("Идентификатор актива")]
        public string? ActiveId { get; set; }
    }
}
