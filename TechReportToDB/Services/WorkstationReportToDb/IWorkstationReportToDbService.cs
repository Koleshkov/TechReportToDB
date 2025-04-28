using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TechReportToDB.Services.WorkstationReportToDb
{
    interface IWorkstationReportToDbService
    {
        Task<string> SaveToDbAsync(IProgress<int> progress, string folderName);
        Task ExportToExcel(string filePath);
    }
}
