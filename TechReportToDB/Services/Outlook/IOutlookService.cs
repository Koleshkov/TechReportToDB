using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TechReportToDB.Services.Outlook
{
    internal interface IOutlookService
    {
        string DownloadDailyReports(IProgress<int> status, string folderName, DateTime selectedDate);
    }
}
