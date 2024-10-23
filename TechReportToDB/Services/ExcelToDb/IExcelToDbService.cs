

namespace TechReportToDB.Services.ExcelToDb
{
    internal interface IExcelToDbService
    {
        Task<string> SaveToolsToDbAsync(IProgress<int> progress, string folderName);

        Task ExportToExcel(string filePath);
    }
}
