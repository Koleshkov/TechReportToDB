using Microsoft.IdentityModel.Tokens;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using TechReportToDB.Data.Models.ShortReport;
using TechReportToDB.Services.Navigation;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Exception = System.Exception;


namespace TechReportToDB.Services.Outlook
{
    internal class OutlookService : IOutlookService
    {
        private readonly INavigationService navigationService;

        public OutlookService(INavigationService navigationService)
        {
            this.navigationService = navigationService;
        }

        public string DownloadDailyReports(IProgress<int> status, string folderName, DateTime selectedDate, int selectedStartTime, int selectedEndTime)
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

                    string filter = "[ReceivedTime] >= '" + selectedDate.ToShortDateString() + $" {selectedStartTime}:01" + "' AND [ReceivedTime] <= '" + selectedDate.ToShortDateString() + $" {selectedEndTime}:00'";

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
                return result.IsNullOrEmpty() ? "" : result;
            }
            catch (System.Exception ex)
            {
                return ex.Message;
            }
        }

        public IEnumerable<Report>? GetShortReportsData(IProgress<int> status, DateTime selectedDate, int selectedTime)
        {
            List<Report> reports = new();
            try
            {
                Application outlookApp = new Application();
                NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
                MAPIFolder selectedFolder = outlookNamespace.PickFolder();

                var items = selectedFolder.Items;

                string filter = "";

                switch (selectedTime)
                {
                    case 0:
                        filter = "[ReceivedTime] >= '" + selectedDate.ToShortDateString() + $" 0:01" + "' AND [ReceivedTime] <= '" + selectedDate.ToShortDateString() + $" 9:00'";
                        break;
                    case 1:
                        filter = "[ReceivedTime] >= '" + selectedDate.ToShortDateString() + $" 9:01" + "' AND [ReceivedTime] <= '" + selectedDate.ToShortDateString() + $" 15:30'";
                        break;
                    case 2:
                        filter = "[ReceivedTime] >= '" + selectedDate.ToShortDateString() + $" 15:31" + "' AND [ReceivedTime] <= '" + selectedDate.ToShortDateString() + $" 23:59'";
                        break;

                }
                

                var filtredItems = items.Restrict(filter);

                foreach (MailItem mail in filtredItems)
                {
                    Report report = new();
                    if (mail.BodyFormat == OlBodyFormat.olFormatHTML) // Проверяем, что тело письма в формате HTML
                    {

                        report.FieldTeam = mail.SenderName.Replace("Полевая партия #","");
                        string htmlBody = mail.HTMLBody; // Получаем HTML-тело письма

                        // Парсим HTML и извлекаем таблицу
                        var tableData = ExtractTableFromHtml(htmlBody);
                        if (tableData != null)
                        {
                            var temp = tableData[0][1].Split(['_', '(']);

                            report.Field = temp[0];
                            report.Pad = temp[1].Replace(" ", "");
                            report.Well = temp[2].Replace(" ", "");
                            report.Type = temp[3].Replace("(", "").Replace(")", "");

                            report.Comment = tableData[1][1] + "\n" + tableData[2][1] + "\n" + tableData[3][1];

                            report.Depth = tableData[4][1];
                            report.Distance = tableData[4][3];

                            reports.Add(report);
                        }
                    }
                    
                }
                return reports;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private List<List<string>> ExtractTableFromHtml(string htmlBody)
        {
            var tableData = new List<List<string>>();

            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(htmlBody);

            var tables = htmlDoc.DocumentNode.SelectNodes("//table");
            if (tables != null)
            {
                foreach (var table in tables)
                {
                    var rows = table.SelectNodes(".//tr");
                    foreach (var row in rows)
                    {
                        var rowData = new List<string>();
                        var cells = row.SelectNodes(".//td");
                        if (cells != null)
                        {
                            foreach (var cell in cells)
                            {
                                rowData.Add(cell.InnerText.Trim());
                            }
                            tableData.Add(rowData);
                        }
                    }
                }
            }

            return tableData;
        }
    }
}
