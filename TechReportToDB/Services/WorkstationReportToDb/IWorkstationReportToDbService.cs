namespace TechReportToDB.Services.WorkstationReportToDb
{
    interface IWorkstationReportToDbService
    {
        Task<string> SaveToDbAsync(IProgress<int> progress, string folderName);
        Task ExportToExcel(string filePath);
    }
}
