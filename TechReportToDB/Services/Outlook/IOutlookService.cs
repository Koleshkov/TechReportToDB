using TechReportToDB.Data.Models.ShortReport;

namespace TechReportToDB.Services.Outlook
{
    internal interface IOutlookService
    {
        string DownloadDailyReports(IProgress<int> status, string folderName, DateTime selectedDate, int selectedStartTime, int selectedEndTime);

        IEnumerable<Report>? GetShortReportsData(IProgress<int> status, DateTime selectedDate, int selectedTime);
    }
}
