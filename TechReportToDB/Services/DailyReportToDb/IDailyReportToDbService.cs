

namespace TechReportToDB.Services.DailyReportToDb
{
    internal interface IDailyReportToDbService
    {
        Task<string> SaveToolsToDbAsync(IProgress<int> progress, string folderName);
        Task ExportToExcel(string filePath);
    }
}
