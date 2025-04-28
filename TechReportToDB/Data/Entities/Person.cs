using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TechReportToDB.Data.Entities
{
    internal class Person : BaseEntity
    {
        public string? Name { get; set; }

        public string? Position { get; set; }

        public string? DateOfBirt { get; set; }

        public string? DateOfJob { get; set; }

        public string? Phone { get; set; }

        public Guid? JobId { get; set; }

        public virtual Job? Job { get; set; }

        [NotMapped]
        public string? FilterName => Name + " " + Position + " " + Job?.Field + " " + Job?.Pad + " " + Job?.Well + " " + Job?.FieldTeam;

    }
}
