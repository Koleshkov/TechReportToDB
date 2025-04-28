using TechReportToDB.Data.Models.ShortReport;

namespace TechReportToDB.Services.ShortReport
{
    internal interface IShortReportService
    {
        Task ExportToExcel(IProgress<int> progress, string filePath, DateTime selectedDate, int selectedTime);
    }
}
