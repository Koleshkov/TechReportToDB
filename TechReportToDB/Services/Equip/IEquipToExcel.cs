using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TechReportToDB.Services.Equip
{
    internal interface IEquipToExcel
    {
        Task ExportToExcel(string filePath);
    }
}
