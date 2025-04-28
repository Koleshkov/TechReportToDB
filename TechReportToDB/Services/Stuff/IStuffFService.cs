using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechReportToDB.Data.Entities;

namespace TechReportToDB.Services.Stuff
{
    internal interface IStuffFService
    {
        Task ExportToExcel(string filePath, IEnumerable<Person> persons);
    }
}
