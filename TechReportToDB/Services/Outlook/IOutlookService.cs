using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechReportToDB.Data.Models.ShortReport;

namespace TechReportToDB.Services.Outlook
{
    internal interface IOutlookService
    {
        string DownloadDailyReports(IProgress<int> status, string folderName, DateTime selectedDate, int selectedStartTime, int selectedEndTime);

        IEnumerable<Report>? GetShortReportsData(IProgress<int> status, DateTime selectedDate, int selectedTime);
    }
}
