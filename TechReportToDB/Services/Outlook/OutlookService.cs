using Microsoft.IdentityModel.Tokens;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System.IO;
using TechReportToDB.Services.Navigation;


namespace TechReportToDB.Services.Outlook
{
    internal class OutlookService : IOutlookService
    {
        private readonly INavigationService navigationService;

        public OutlookService(INavigationService navigationService)
        {
            this.navigationService = navigationService;
        }

        public string DownloadDailyReports(IProgress<int> status, string folderName, DateTime selectedDate)
        {
            
            string result = "";
            try
            {
                Application outlookApp = new Application();
                NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
                MAPIFolder selectedFolder = outlookNamespace.PickFolder();
                var attachmentList = new List<Attachment>();
                if (selectedFolder != null)
                {
                    var items = selectedFolder.Items;

                    string filter = "[ReceivedTime] >= '" + selectedDate.AddDays(1).ToShortDateString() + " 0:01" + "' AND [ReceivedTime] <= '" + selectedDate.AddDays(1).ToShortDateString() + " 10:00'";

                    var filtredItems = items.Restrict(filter);

                    foreach (var item in filtredItems)
                    {
                        if (item is MailItem mailItem)
                        {
                            if (mailItem.Attachments.Count > 0)
                            {
                                foreach (Attachment attachment in mailItem.Attachments)
                                {
                                    var subject = mailItem.Subject;
                                    var name = attachment.FileName;

                                    if (subject.Contains("СР") && name.Contains("СР"))
                                    {
                                        result = $"{folderName}\\{selectedDate.ToShortDateString()}";
                                        if (!Directory.Exists(result)) Directory.CreateDirectory(result);

                                        
                                        attachmentList.Add(attachment);
                                    }
                                }
                            }
                        }
                    }
                    int i = 1;
                    Parallel.ForEach(attachmentList.Cast<Attachment>(), attachment =>
                    {
                        attachment.SaveAsFile($"{result}\\{attachment.FileName}");
                        status.Report(((i + 1) * 100) / attachmentList.Count);
                        i++;
                    });
  

                }
                return result.IsNullOrEmpty()?"":result;
            }
            catch (System.Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
