using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.IO;
using TechReportToDB.Converters;
using TechReportToDB.Data.Entities;
using TechReportToDB.Data.Models.ShortReport;
using TechReportToDB.Services.Outlook;

namespace TechReportToDB.Services.ShortReport
{
    internal class ShortReportService : IShortReportService
    {
        private readonly IOutlookService outlookService;

        public ShortReportService(IOutlookService outlookService)
        {
            this.outlookService = outlookService;
        }

        public async Task ExportToExcel(IProgress<int> progress, string filePath, DateTime selectedDate, int selectedTime)
        {
            var shortReports = outlookService.GetShortReportsData(progress, selectedDate, selectedTime)?.ToList();

            string date = selectedDate.ToString("dd.MM.yy");

            if (File.Exists(filePath) && shortReports != null)
            {
                var tempFile = filePath.Replace(".xlsx", "") + "_New.xlsx";

                File.Copy(filePath, tempFile, overwrite: true);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                FileInfo fileInfo = new FileInfo(tempFile);

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorkbook wb = package.Workbook;
                    ExcelWorksheet sheet = wb.Worksheets.Last();
                    string sheetName = wb.Worksheets.Last().Name;
                    ExcelTable? table = sheet.Tables.FirstOrDefault();
                    if (wb.Worksheets.All(ws => ws.Name != date))
                    {

                        string oldSheetName = sheet.Name; // Имя оригинального листа
                        sheet = wb.Worksheets.Add(date, sheet);

                        // Проверяем таблицу
                        table = sheet.Tables.FirstOrDefault();
                        if (table != null) CleanExcel(sheet, table);
                    }

                    table = sheet.Tables.FirstOrDefault();

                    if (table != null)
                    {
                        foreach (var sr in shortReports)
                        {
                            int existingRow = Exixst(sheet, table, sr);
                            int startRow = table.Address.Start.Row + 1;
                            int endRow = table.Address.End.Row - 1;
                            switch (selectedTime)
                            {
                                case 0:
                                    if (existingRow != 0)
                                    {
                                        sheet.Cells[existingRow, 2].Value = sr.FieldTeam;
                                        sheet.Cells[existingRow, 7].Value = sr.Field;
                                        sheet.Cells[existingRow, 8].Value = sr.Pad;
                                        sheet.Cells[existingRow, 9].Value = sr.Well;
                                        sheet.Cells[existingRow, 10].Value = sr.Type;
                                        sheet.Cells[existingRow, 11].Value = sr.Section;
                                        sheet.Cells[existingRow, 13].Value = sr.Comment;
                                        sheet.Cells[existingRow, 14].Value = CC.ConvertStringToDouble(sr.Depth);
                                        sheet.Cells[existingRow, 15].Value = CC.ConvertStringToDouble(sr.Distance);
                                    }
                                    else
                                    {
                                        sheet.Cells[endRow, 2].Value = sr.FieldTeam;
                                        sheet.Cells[endRow, 7].Value = sr.Field;
                                        sheet.Cells[endRow, 8].Value = sr.Pad;
                                        sheet.Cells[endRow, 9].Value = sr.Well;
                                        sheet.Cells[endRow, 10].Value = sr.Type;
                                        sheet.Cells[endRow, 11].Value = sr.Section;
                                        sheet.Cells[endRow, 13].Value = sr.Comment;
                                        sheet.Cells[endRow, 14].Value = CC.ConvertStringToDouble(sr.Depth);
                                        sheet.Cells[endRow, 15].Value = CC.ConvertStringToDouble(sr.Distance);
                                        sheet.Cells[endRow, 16].Formula = $"IFERROR(O{endRow}/L{endRow},\"\")";
                                        table.AddRow();
                                    }

                                    break;

                                case 1:
                                    if (existingRow != 0)
                                    {
                                        sheet.Cells[existingRow, 2].Value = sr.FieldTeam;
                                        sheet.Cells[existingRow, 7].Value = sr.Field;
                                        sheet.Cells[existingRow, 8].Value = sr.Pad;
                                        sheet.Cells[existingRow, 9].Value = sr.Well;
                                        sheet.Cells[existingRow, 10].Value = sr.Type;
                                        sheet.Cells[existingRow, 11].Value = sr.Section;
                                        sheet.Cells[existingRow, 17].Value = sr.Comment;
                                        sheet.Cells[existingRow, 18].Value = CC.ConvertStringToDouble(sr.Depth);
                                        sheet.Cells[existingRow, 19].Value = CC.ConvertStringToDouble(sr.Distance);

                                    }
                                    else
                                    {
                                        sheet.Cells[endRow, 2].Value = sr.FieldTeam;
                                        sheet.Cells[endRow, 7].Value = sr.Field;
                                        sheet.Cells[endRow, 8].Value = sr.Pad;
                                        sheet.Cells[endRow, 9].Value = sr.Well;
                                        sheet.Cells[endRow, 10].Value = sr.Type;
                                        sheet.Cells[endRow, 11].Value = sr.Section;
                                        sheet.Cells[endRow, 17].Value = sr.Comment;
                                        sheet.Cells[endRow, 18].Value = CC.ConvertStringToDouble(sr.Depth);
                                        sheet.Cells[endRow, 19].Value = CC.ConvertStringToDouble(sr.Distance);
                                        sheet.Cells[endRow, 20].Formula = $"IFERROR(S{endRow}/L{endRow},\"\")";
                                        table.AddRow();
                                    }
                                    break;

                                case 2:
                                    if (existingRow != 0)
                                    {
                                        sheet.Cells[existingRow, 2].Value = sr.FieldTeam;
                                        sheet.Cells[existingRow, 7].Value = sr.Field;
                                        sheet.Cells[existingRow, 8].Value = sr.Pad;
                                        sheet.Cells[existingRow, 9].Value = sr.Well;
                                        sheet.Cells[existingRow, 10].Value = sr.Type;
                                        sheet.Cells[existingRow, 11].Value = sr.Section;
                                        sheet.Cells[existingRow, 21].Value = sr.Comment;
                                        sheet.Cells[existingRow, 22].Value = CC.ConvertStringToDouble(sr.Depth);
                                        sheet.Cells[existingRow, 23].Value = CC.ConvertStringToDouble(sr.Distance);
                                    }
                                    else
                                    {
                                        sheet.Cells[endRow, 2].Value = sr.FieldTeam;
                                        sheet.Cells[endRow, 7].Value = sr.Field;
                                        sheet.Cells[endRow, 8].Value = sr.Pad;
                                        sheet.Cells[endRow, 9].Value = sr.Well;
                                        sheet.Cells[endRow, 10].Value = sr.Type;
                                        sheet.Cells[endRow, 11].Value = sr.Section;
                                        sheet.Cells[endRow, 21].Value = sr.Comment;
                                        sheet.Cells[endRow, 22].Value = CC.ConvertStringToDouble(sr.Depth);
                                        sheet.Cells[endRow, 23].Value = CC.ConvertStringToDouble(sr.Distance);
                                        sheet.Cells[endRow, 24].Formula = $"IFERROR(X{endRow}/L{endRow},\"\")";
                                        table.AddRow();
                                    }
                                    break;
                            }


                        }
                        sheet.DeleteRow(table.Address.End.Row - 1);
                    }
                    package.Save();
                }
            }
            await Task.CompletedTask;
        }

        private void CleanExcel(ExcelWorksheet worksheet, ExcelTable table)
        {
            int startRow = table.Address.Start.Row + 1;
            int endRow = table.Address.End.Row;

            for (int i = endRow - startRow; i > 1; i--)
            {
                table.DeleteRow(i - 2);
            }
            worksheet.Cells["A2:AA2"].Value = "";
        }

        private int Exixst(ExcelWorksheet worksheet, ExcelTable table, Report report)
        {
            int startRow = table.Address.Start.Row + 1;
            int endRow = table.Address.End.Row;

            for (int i = startRow; i < endRow; i++)
            {
                string ft = worksheet.Cells[$"B{i}"].Text;
                string well = worksheet.Cells[$"I{i}"].Text;
                if (ft == report.FieldTeam && well == report.Well)
                {
                    return i;
                }
            }

            return 0;
        }

    }
}
