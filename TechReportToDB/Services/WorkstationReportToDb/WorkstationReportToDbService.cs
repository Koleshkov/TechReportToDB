using Microsoft.IdentityModel.Tokens;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.IO;
using TechReportToDB.Data.Models.Workstations;
using TechReportToDB.Data.Models.Workstations.Items;

namespace TechReportToDB.Services.WorkstationReportToDb
{
    public class WorkstationReportToDbService : IWorkstationReportToDbService
    {
        public List<Workstation> Workstations { get; set; } = new();

        public async Task ExportToExcel(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo fileInfo = new FileInfo("Templates\\WorkstationsTemplate.xlsx");

            Workstations = Workstations.OrderBy(f => f.FieldTeam).ThenBy(p => p.Type).ToList();

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Внешний_вид"];
                if (worksheet != null)
                {
                    var table = worksheet.Tables[$"УТ_Внешний_вид"];

                    if (table != null)
                    {
                        CleanExcel(worksheet, table);
                        int row = table.Address.End.Row;

                        foreach (var workstation in Workstations)
                        {
                            foreach (var availability in workstation.AvailabilityList)
                            {
                                worksheet.Cells[row, 1].Value = workstation.Field;
                                worksheet.Cells[row, 2].Value = workstation.Pad;
                                worksheet.Cells[row, 3].Value = workstation.Well;
                                worksheet.Cells[row, 4].Value = workstation.RegNumber;
                                worksheet.Cells[row, 5].Value = workstation.SerialNubmer;
                                worksheet.Cells[row, 6].Value = workstation.Type;
                                worksheet.Cells[row, 7].Value = workstation.Date;
                                worksheet.Cells[row, 8].Value = workstation.ActiveId;
                                worksheet.Cells[row, 9].Value = workstation.GenEngineer;
                                worksheet.Cells[row, 10].Value = availability.Name;
                                worksheet.Cells[row, 11].Value = availability.QuantityPlan;
                                worksheet.Cells[row, 12].Value = availability.QuantityFact;
                                worksheet.Cells[row, 13].Value = availability.Status;
                                worksheet.Cells[row, 14].Value = availability.Comment;
                                worksheet.Cells[row, 15].Value = workstation.FilePath;
                                table.AddRow();
                                row++;
                            }
                        }
                    }

                    worksheet = package.Workbook.Worksheets["Система_оповещения"];
                    table = worksheet.Tables[$"УТ_Система_оповещения"];

                    if (table != null)
                    {
                        CleanExcel(worksheet, table);
                        int row = table.Address.End.Row;

                        foreach (var workstation in Workstations)
                        {
                            foreach (var fireExtinguisher in workstation.FireExtinguisherList)
                            {
                                worksheet.Cells[row, 1].Value = workstation.Field;
                                worksheet.Cells[row, 2].Value = workstation.Pad;
                                worksheet.Cells[row, 3].Value = workstation.Well;
                                worksheet.Cells[row, 4].Value = workstation.RegNumber;
                                worksheet.Cells[row, 5].Value = workstation.SerialNubmer;
                                worksheet.Cells[row, 6].Value = workstation.Type;
                                worksheet.Cells[row, 7].Value = workstation.Date;
                                worksheet.Cells[row, 8].Value = workstation.ActiveId;
                                worksheet.Cells[row, 9].Value = workstation.GenEngineer;
                                worksheet.Cells[row, 10].Value = fireExtinguisher.Name;
                                worksheet.Cells[row, 11].Value = fireExtinguisher.SerialNumber;
                                worksheet.Cells[row, 12].Value = fireExtinguisher.CreationDate;
                                worksheet.Cells[row, 13].Value = fireExtinguisher.Status;
                                worksheet.Cells[row, 14].Value = fireExtinguisher.Comment;
                                worksheet.Cells[row, 15].Value = workstation.FilePath;
                                table.AddRow();
                                row++;
                            }
                        }
                    }

                    worksheet = package.Workbook.Worksheets["Электросистема"];
                    table = worksheet.Tables[$"УТ_Электросистема"];

                    if (table != null)
                    {
                        CleanExcel(worksheet, table);
                        int row = table.Address.End.Row;

                        foreach (var workstation in Workstations)
                        {
                            foreach (var eSystem in workstation.ESystemList)
                            {
                                worksheet.Cells[row, 1].Value = workstation.Field;
                                worksheet.Cells[row, 2].Value = workstation.Pad;
                                worksheet.Cells[row, 3].Value = workstation.Well;
                                worksheet.Cells[row, 4].Value = workstation.RegNumber;
                                worksheet.Cells[row, 5].Value = workstation.SerialNubmer;
                                worksheet.Cells[row, 6].Value = workstation.Type;
                                worksheet.Cells[row, 7].Value = workstation.Date;
                                worksheet.Cells[row, 8].Value = workstation.ActiveId;
                                worksheet.Cells[row, 9].Value = workstation.GenEngineer;
                                worksheet.Cells[row, 10].Value = eSystem.Name;
                                worksheet.Cells[row, 11].Value = eSystem.QuantityPlan;
                                worksheet.Cells[row, 12].Value = eSystem.QuantityFact;
                                worksheet.Cells[row, 13].Value = eSystem.Status;
                                worksheet.Cells[row, 14].Value = eSystem.Comment;
                                worksheet.Cells[row, 15].Value = workstation.FilePath;
                                table.AddRow();
                                row++;
                            }
                        }
                    }


                    worksheet = package.Workbook.Worksheets["Недостатки_общие"];
                    table = worksheet.Tables[$"УТ_Недостатки_общие"];

                    if (table != null)
                    {
                        CleanExcel(worksheet, table);
                        int row = table.Address.End.Row;

                        foreach (var workstation in Workstations)
                        {
                            var es = workstation.ESystemList.Where(l => l.Status == "Неисправно" || l.Status == "Некомплект").ToList();

                            foreach (var eSystem in es)
                            {
                                worksheet.Cells[row, 1].Value = workstation.Field;
                                worksheet.Cells[row, 2].Value = workstation.Pad;
                                worksheet.Cells[row, 3].Value = workstation.Well;
                                worksheet.Cells[row, 4].Value = workstation.RegNumber;
                                worksheet.Cells[row, 5].Value = workstation.SerialNubmer;
                                worksheet.Cells[row, 6].Value = workstation.Type;
                                worksheet.Cells[row, 7].Value = workstation.Date;
                                worksheet.Cells[row, 8].Value = workstation.ActiveId;
                                worksheet.Cells[row, 9].Value = workstation.GenEngineer;
                                worksheet.Cells[row, 10].Value = eSystem.Name;
                                worksheet.Cells[row, 11].Value = eSystem.QuantityPlan;
                                worksheet.Cells[row, 12].Value = eSystem.QuantityFact;
                                worksheet.Cells[row, 13].Value = eSystem.Status;
                                worksheet.Cells[row, 14].Value = eSystem.Comment;
                                worksheet.Cells[row, 15].Value = workstation.FilePath;
                                table.AddRow();
                                row++;
                            }

                            var al = workstation.AvailabilityList.Where(l => l.Status == "Неисправно" || l.Status == "Некомплект").ToList();

                            foreach (var eSystem in al)
                            {
                                worksheet.Cells[row, 1].Value = workstation.Field;
                                worksheet.Cells[row, 2].Value = workstation.Pad;
                                worksheet.Cells[row, 3].Value = workstation.Well;
                                worksheet.Cells[row, 4].Value = workstation.RegNumber;
                                worksheet.Cells[row, 5].Value = workstation.SerialNubmer;
                                worksheet.Cells[row, 6].Value = workstation.Type;
                                worksheet.Cells[row, 7].Value = workstation.Date;
                                worksheet.Cells[row, 8].Value = workstation.ActiveId;
                                worksheet.Cells[row, 9].Value = workstation.GenEngineer;
                                worksheet.Cells[row, 10].Value = eSystem.Name;
                                worksheet.Cells[row, 11].Value = eSystem.QuantityPlan;
                                worksheet.Cells[row, 12].Value = eSystem.QuantityFact;
                                worksheet.Cells[row, 13].Value = eSystem.Status;
                                worksheet.Cells[row, 14].Value = eSystem.Comment;
                                worksheet.Cells[row, 15].Value = workstation.FilePath;
                                table.AddRow();
                                row++;
                            }

                            var fr = workstation.FireExtinguisherList.Where(l => l.Status == "Неисправно" || l.Status == "Некомплект").ToList();

                            foreach (var eSystem in fr)
                            {
                                worksheet.Cells[row, 1].Value = workstation.Field;
                                worksheet.Cells[row, 2].Value = workstation.Pad;
                                worksheet.Cells[row, 3].Value = workstation.Well;
                                worksheet.Cells[row, 4].Value = workstation.RegNumber;
                                worksheet.Cells[row, 5].Value = workstation.SerialNubmer;
                                worksheet.Cells[row, 6].Value = workstation.Type;
                                worksheet.Cells[row, 7].Value = workstation.Date;
                                worksheet.Cells[row, 8].Value = workstation.ActiveId;
                                worksheet.Cells[row, 9].Value = workstation.GenEngineer;
                                worksheet.Cells[row, 10].Value = eSystem.Name;
                                worksheet.Cells[row, 13].Value = eSystem.Status;
                                worksheet.Cells[row, 14].Value = eSystem.Comment;
                                worksheet.Cells[row, 15].Value = workstation.FilePath;
                                table.AddRow();
                                row++;
                            }
                        }
                    }


                    worksheet = package.Workbook.Worksheets["Для_презентации"];

                    if (worksheet != null)
                    {
                        int row = 5;

                        foreach (var workstation in Workstations.Where(t => t.Type != null ? t.Type.Contains("Офис") : false).OrderBy(f => f.FieldTeam))
                        {
                            if(row!=5) CopyRangeFormat(worksheet.Cells[$"A5:P6"], worksheet.Cells[$"A{row}:P{row + 1}"]);

                            worksheet.Cells[row, 1].Value = "РН-ЮНГ";
                            worksheet.Cells[row, 2].Value = workstation.FieldTeam;
                            worksheet.Cells[row, 3].Value = workstation.Field;
                            worksheet.Cells[row, 4].Value = workstation.Pad;
                            worksheet.Cells[row, 5].Value = workstation.RegNumber;

                            var temp = workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("электрический обогреватель") : false)?.QuantityFact;

                            worksheet.Cells[row, 6].Value = temp;

                            temp = workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("тепловая завеса") : false)?.QuantityFact;

                            worksheet.Cells[row, 7].Value = temp;

                            temp = workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("оповещатель") : false)?.QuantityFact;

                            worksheet.Cells[row, 8].Value = temp;

                            temp = workstation?.FireExtinguisherList[0].Name;

                            if (!temp.IsNullOrEmpty())
                            {
                                worksheet.Cells[row, 9].Value = 1;

                                worksheet.Cells[row, 10].Value = temp;

                                worksheet.Cells[row, 11].Value = workstation?.FireExtinguisherList[0].SerialNumber;

                                worksheet.Cells[row, 12].Value = workstation?.FireExtinguisherList[0].CreationDate;
                            }

                            temp = workstation?.FireExtinguisherList[1].Name;

                            if (!temp.IsNullOrEmpty())
                            {
                                worksheet.Cells[row + 1, 9].Value = 1;

                                worksheet.Cells[row + 1, 10].Value = temp;

                                worksheet.Cells[row + 1, 11].Value = workstation?.FireExtinguisherList[0].SerialNumber;

                                worksheet.Cells[row + 1, 12].Value = workstation?.FireExtinguisherList[0].CreationDate;
                            }


                            temp = workstation?.CheckList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("состояние уплотнителей входной") : false)?.Status;

                            worksheet.Cells[row, 13].Value = temp;

                            temp = workstation?.AvailabilityList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("штыковая лопата") : false)?.QuantityFact;

                            worksheet.Cells[row, 14].Value = temp;

                            temp = workstation?.AvailabilityList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("снеговая лопата") : false)?.QuantityFact;

                            worksheet.Cells[row, 15].Value = temp;

                            temp = workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("пожарный") : false)?.Comment;

                            temp = temp + ", " + workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("конвектор") : false)?.Comment;
                            worksheet.Cells[row, 16].Value = temp;

                            row = row + 2;
                        }


                        row = 5;
                        foreach (var workstation in Workstations.Where(t => t.Type != null ? t.Type.Contains("Жилой") : false).OrderBy(f => f.FieldTeam))
                        {

                            if (row!=5) CopyRangeFormat(worksheet.Cells[$"Q5:AB6"], worksheet.Cells[$"Q{row}:AB{row + 1}"]);

                            worksheet.Cells[row, 17].Value = workstation.RegNumber;

                            var temp = workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("электрический обогреватель") : false)?.QuantityFact;

                            worksheet.Cells[row, 18].Value = temp;

                            temp = workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("тепловая завеса") : false)?.QuantityFact;

                            worksheet.Cells[row, 19].Value = temp;

                            temp = workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("оповещатель") : false)?.QuantityFact;

                            worksheet.Cells[row, 20].Value = temp;

                            temp = workstation?.FireExtinguisherList[0].Name;

                            if (!temp.IsNullOrEmpty())
                            {
                                worksheet.Cells[row, 21].Value = 1;

                                worksheet.Cells[row, 22].Value = temp;

                                worksheet.Cells[row, 23].Value = workstation?.FireExtinguisherList[0].SerialNumber;

                                worksheet.Cells[row, 24].Value = workstation?.FireExtinguisherList[0].CreationDate;
                            }

                            temp = workstation?.FireExtinguisherList[1].Name;

                            if (!temp.IsNullOrEmpty())
                            {
                                worksheet.Cells[row + 1, 21].Value = 1;

                                worksheet.Cells[row + 1, 22].Value = temp;

                                worksheet.Cells[row + 1, 23].Value = workstation?.FireExtinguisherList[0].SerialNumber;

                                worksheet.Cells[row + 1, 24].Value = workstation?.FireExtinguisherList[0].CreationDate;
                            }


                            temp = workstation?.CheckList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("состояние уплотнителей входной") : false)?.Status;

                            worksheet.Cells[row, 25].Value = temp;

                            temp = workstation?.AvailabilityList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("штыковая лопата") : false)?.QuantityFact;

                            worksheet.Cells[row, 26].Value = temp;

                            temp = workstation?.AvailabilityList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("снеговая лопата") : false)?.QuantityFact;

                            worksheet.Cells[row, 27].Value = temp;

                            temp = workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("оповещатель") : false)?.Comment;

                            temp = temp + ", " + workstation?.ESystemList.FirstOrDefault(n => n.Name != null ? n.Name.ToLower().Contains("конвектор") : false)?.Comment;
                            worksheet.Cells[row, 28].Value = temp;

                            row = row + 2;
                        }
                    }
                }

                package.SaveAs(filePath);
            }
            await Task.CompletedTask;
        }

        public async Task<string> SaveToDbAsync(IProgress<int> progress, string folderName)
        {
            try
            {
                string[] files = Directory.GetFiles(folderName, "*.xlsm", SearchOption.TopDirectoryOnly);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                int fileNumber = 0;

                foreach (var file in files)
                {
                    if (file.ToLower().Contains("ежемесячный отчет по вагон"))
                    {



                        fileNumber++;
                        progress.Report((fileNumber * 100 / files.Length));

                        using (var package = new ExcelPackage(new FileInfo(file)))
                        {

                            ExcelWorksheet worksheet = package.Workbook.Worksheets["Общие данные"];
                            Workstation workstation = new();
                            workstation.FilePath = file;
                            workstation.Field = worksheet.Cells["B2"].Text;
                            workstation.Pad = worksheet.Cells["B3"].Text;
                            workstation.Well = worksheet.Cells["B4"].Text;
                            workstation.FieldTeam = worksheet.Cells["B5"].Text;
                            workstation.RegNumber = worksheet.Cells["B6"].Text;
                            workstation.SerialNubmer = worksheet.Cells["B7"].Text;
                            workstation.Type = worksheet.Cells["B8"].Text;
                            workstation.Date = worksheet.Cells["B9"].Text;
                            workstation.ActiveId = worksheet.Cells["B10"].Text;
                            workstation.GenEngineer = worksheet.Cells["B12"].Text;

                            worksheet = package.Workbook.Worksheets["Чек лист"];

                            for (int i = 10; i < 28; i++)
                            {
                                if (i != 26)
                                {
                                    workstation.CheckList.Add(new CheckListItem
                                    {
                                        Name = worksheet.Cells[$"B{i}"].Text,
                                        Status = worksheet.Cells[$"F{i}"].Text,
                                        Comment = worksheet.Cells[$"G{i}"].Text
                                    });
                                }
                            }

                            worksheet = package.Workbook.Worksheets["Документация"];

                            for (int i = 61; i < 64; i++)
                            {
                                workstation.DocumentationList.Add(new DocumentationItem
                                {
                                    Name = worksheet.Cells[$"A{i}"].Text,
                                    StartDate = worksheet.Cells[$"B{i}"].Text,
                                    EndDate = worksheet.Cells[$"C{i}"].Text,
                                    Days = worksheet.Cells[$"D{i}"].Text,
                                    DaysToEnd = worksheet.Cells[$"E{i}"].Text,
                                });
                            }

                            worksheet = package.Workbook.Worksheets["Внешний вид"];

                            for (int i = 80; i < 91; i++)
                            {

                                workstation.AvailabilityList.Add(new AvailabilityItem
                                {
                                    Name = worksheet.Cells[$"A{i}"].Text,
                                    QuantityPlan = worksheet.Cells[$"B{i}"].Text,
                                    QuantityFact = worksheet.Cells[$"C{i}"].Text,
                                    Status = worksheet.Cells[$"D{i}"].Text,
                                    Comment = worksheet.Cells[$"E{i}"].Text,
                                });
                            }


                            worksheet = package.Workbook.Worksheets["Комплектация"];
                            if (workstation.Type.Contains("Жилой"))
                            {
                                for (int i = 80; i < 91; i++)
                                {
                                    workstation.AvailabilityList.Add(new AvailabilityItem
                                    {
                                        Name = worksheet.Cells[$"A{i}"].Text,
                                        QuantityPlan = worksheet.Cells[$"B{i}"].Text,
                                        QuantityFact = worksheet.Cells[$"C{i}"].Text,
                                        Status = worksheet.Cells[$"D{i}"].Text,
                                        Comment = worksheet.Cells[$"E{i}"].Text,
                                    });
                                }
                            }

                            if (workstation.Type.Contains("Офис"))
                            {
                                for (int i = 99; i < 120; i++)
                                {
                                    workstation.AvailabilityList.Add(new AvailabilityItem
                                    {
                                        Name = worksheet.Cells[$"A{i}"].Text,
                                        QuantityPlan = worksheet.Cells[$"B{i}"].Text,
                                        QuantityFact = worksheet.Cells[$"C{i}"].Text,
                                        Status = worksheet.Cells[$"D{i}"].Text,
                                        Comment = worksheet.Cells[$"E{i}"].Text,
                                    });
                                }
                            }





                            worksheet = package.Workbook.Worksheets["Электросистема"];

                            for (int i = 269; i < 280; i++)
                            {
                                workstation.ESystemList.Add(new ESystemItem
                                {
                                    Name = worksheet.Cells[$"A{i}"].Text,
                                    QuantityPlan = worksheet.Cells[$"B{i}"].Text,
                                    QuantityFact = worksheet.Cells[$"C{i}"].Text,
                                    Status = worksheet.Cells[$"D{i}"].Text,
                                    Comment = worksheet.Cells[$"E{i}"].Text,
                                });
                            }

                            worksheet = package.Workbook.Worksheets["СОП"];
                            if (worksheet == null)
                                worksheet = package.Workbook.Worksheets["СОП "];

                            for (int i = 99; i < 102; i++)
                            {
                                workstation.FireExtinguisherList.Add(new FireExtinguisherItem
                                {
                                    Name = worksheet.Cells[$"A{i}"].Text,
                                    SerialNumber = worksheet.Cells[$"B{i}"].Text,
                                    CreationDate = worksheet.Cells[$"C{i}"].Text,
                                    Status = worksheet.Cells[$"D{i}"].Text,
                                    Comment = worksheet.Cells[$"E{i}"].Text,
                                });
                            }

                            Workstations.Add(workstation);

                        }
                    }

                }
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            return await Task.FromResult("Данные загружены.");
        }

        private void CleanExcel(ExcelWorksheet worksheet, ExcelTable table)
        {
            int startRow = table.Address.Start.Row;
            int endRow = table.Address.End.Row;

            for (int i = endRow - startRow; i > 1; i--)
            {
                table.DeleteRow(i - 1);
            }
        }

        private void CopyRangeFormat(ExcelRange sourceRange, ExcelRange targetRange)
        {
            // Проверяем размеры диапазонов
            if (sourceRange.Rows != targetRange.Rows || sourceRange.Columns != targetRange.Columns)
            {
                throw new ArgumentException("Source and target ranges must have the same dimensions.");
            }

            // Копируем формат ячеек
            for (int row = sourceRange.Start.Row; row <= sourceRange.End.Row; row++)
            {
                for (int col = sourceRange.Start.Column; col <= sourceRange.End.Column; col++)
                {
                    var sourceCell = sourceRange.Worksheet.Cells[row, col];
                    var targetCell = targetRange.Worksheet.Cells[
                        targetRange.Start.Row + (row - sourceRange.Start.Row),
                        targetRange.Start.Column + (col - sourceRange.Start.Column)];

                    // Копируем стиль
                    targetCell.StyleID = sourceCell.StyleID;

                    // Копируем формулу
                    if (!string.IsNullOrEmpty(sourceCell.Formula))
                    {
                        targetCell.Formula = sourceCell.Formula;
                    }
                }
            }

            // Копируем объединения (если они есть)
            var mergedCellsCopy = sourceRange.Worksheet.MergedCells.ToList(); // Создаем копию коллекции
            foreach (var mergedCellAddress in mergedCellsCopy)
            {
                var mergedRange = sourceRange.Worksheet.Cells[mergedCellAddress];

                // Проверяем, входит ли объединение в диапазон sourceRange
                if (sourceRange.Start.Row <= mergedRange.Start.Row && mergedRange.End.Row <= sourceRange.End.Row &&
                    sourceRange.Start.Column <= mergedRange.Start.Column && mergedRange.End.Column <= sourceRange.End.Column)
                {
                    // Вычисляем смещение для целевого диапазона
                    int rowOffset = mergedRange.Start.Row - sourceRange.Start.Row;
                    int colOffset = mergedRange.Start.Column - sourceRange.Start.Column;

                    // Применяем объединение к целевому диапазону
                    targetRange.Worksheet.Cells[
                        targetRange.Start.Row + rowOffset, targetRange.Start.Column + colOffset,
                        targetRange.Start.Row + rowOffset + (mergedRange.End.Row - mergedRange.Start.Row),
                        targetRange.Start.Column + colOffset + (mergedRange.End.Column - mergedRange.Start.Column)
                    ].Merge = true;
                }
            }
        }


    }
}
