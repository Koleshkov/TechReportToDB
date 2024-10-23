using ExcelDataReader;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Data;
using System.IO;
using System.Windows;
using TechReportToDB.Converters;
using TechReportToDB.Data;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services.Navigation;

namespace TechReportToDB.Services.ExcelToDb
{
    internal class ExcelToDbService : IExcelToDbService
    {
        private readonly INavigationService navigationService;
        private readonly AppDbContext context;

        private readonly List<Job> jobs = new();

        public ExcelToDbService(INavigationService navigationService, AppDbContext context)
        {
            this.navigationService = navigationService;
            this.context = context;
        }

        public async Task ExportToExcel(string filePath)
        {
            var tools = await context.Tools.Include(t => t.Job).ToListAsync();
            var kits = await context.Kits.Include(k => k.Job).ToListAsync();
            var kitTools = await context.KitTools.Include(k => k.Kit).ThenInclude(k => k.Job).ToListAsync();
            await Task.Run(() => CreateToolsPivotTable(filePath, tools, kits, kitTools));
        }

        public async Task<string> SaveToolsToDbAsync(IProgress<int> progress, string folderName)
        {
            try
            {
                if (await context.Jobs.AnyAsync())
                {
                    var tables = context.Model.GetEntityTypes()
                                .Select(t => t.GetTableName())
                                .Distinct();

                    foreach (var table in tables)
                    {
                        context.Database.ExecuteSqlRaw($"DELETE FROM [{table}]");
                    }

                    context.SaveChanges();
                }

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                string[] files = Directory.GetFiles(folderName, "*.xlsb", SearchOption.TopDirectoryOnly);

                for (int i = 0; i < files.Length; i++)
                {
                    if (!files[i].Contains("~") && files[i].Contains("СР"))
                    {
                        var job = new Job();

                        string fileName = Path.GetFileName(files[i]);
                        string[] splitedFileName = fileName.Split("_");

                        job.FilePath = files[i];

                        var tempDate = files[i].Substring(files[i].Length - 15).Split(".")[0].Split("-");

                        string drListName = $"{tempDate[2]}.{tempDate[1]}.{tempDate[0].Substring(2)}";

                        using (var stream = File.Open(files[i], FileMode.Open, FileAccess.Read))
                        {
                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                int rowIndex = 0;
                                

                                do
                                {

                                    if (reader.Name == drListName)
                                    {
                                        rowIndex = 0;

                                        while (reader.Read() && rowIndex < 14)
                                        {
                                            if (rowIndex == 4) job.Field = reader.GetValue(5)?.ToString();

                                            if (rowIndex == 5)
                                            {
                                                job.FieldTeam = reader.GetValue(5)?.ToString();
                                                job.DDs.Add(new DD { Name = reader.GetValue(28)?.ToString() });
                                            }

                                            if (rowIndex == 6)
                                            {
                                                job.Pad = reader.GetValue(5)?.ToString();
                                                var name = reader.GetValue(28)?.ToString();
                                                if (name != null) job.DDs.Add(new DD { Name = reader.GetValue(28)?.ToString() });

                                            }

                                            if (rowIndex == 7)
                                            {
                                                job.Well = reader.GetValue(5)?.ToString();
                                                var name = reader.GetValue(28)?.ToString();
                                                if (name != null) job.DDs.Add(new DD { Name = reader.GetValue(28)?.ToString() });
                                            }

                                            if (rowIndex == 7)
                                            {
                                                var name = reader.GetValue(28)?.ToString();
                                                if (name != null) job.DDs.Add(new DD { Name = reader.GetValue(28)?.ToString() });
                                            }

                                            if (rowIndex == 11)
                                            {
                                                job.Phone = reader.GetValue(5)?.ToString();
                                                var name = reader.GetValue(28)?.ToString();
                                                if (name != null) job.MWDs.Add(new MWD { Name = reader.GetValue(28)?.ToString() });
                                            }
                                            if (rowIndex == 12)
                                            {
                                                job.Type = reader.GetValue(5)?.ToString();
                                                var name = reader.GetValue(28)?.ToString();
                                                if (name != null) job.MWDs.Add(new MWD { Name = reader.GetValue(28)?.ToString() });
                                            }
                                            if (rowIndex == 13)
                                            {
                                                job.DrillingContractor = reader.GetValue(5)?.ToString();
                                                var name = reader.GetValue(28)?.ToString();
                                                if (name != null) job.MWDs.Add(new MWD { Name = reader.GetValue(28)?.ToString() });
                                            }
                                            if (rowIndex == 14)
                                            {
                                                var name = reader.GetValue(28)?.ToString();
                                                if (name != null) job.MWDs.Add(new MWD { Name = reader.GetValue(28)?.ToString() });
                                            }
                                            rowIndex++;
                                        }
                                    }

                                    if (reader.Name == "Оборудование")
                                    {
                                        rowIndex = 0;

                                        while (reader.Read() && rowIndex < 150)
                                        {
                                            if (rowIndex > 10)
                                            {
                                                Tool tool = new();

                                                tool.Size = reader.GetValue(0)?.ToString();

                                                tool.ToolClass = reader.GetValue(1)?.ToString();

                                                tool.Owner = reader.GetValue(2)?.ToString();

                                                tool.Name = reader.GetValue(3)?.ToString();

                                                tool.SerialNumber = reader.GetValue(4)?.ToString();

                                                tool.Art = reader.GetValue(5)?.ToString();

                                                tool.Pasport = reader.GetValue(6)?.ToString();

                                                tool.InspectionDate = CC.ConvertStringToDateTimeString(reader.GetValue(7)?.ToString());

                                                tool.Days = CC.ConvertStringToInt(reader.GetValue(8)?.ToString()); ;

                                                tool.CircTimeAfterInspection = CC.ConvertStringToDouble(reader.GetValue(10)?.ToString());

                                                tool.CircTime = CC.ConvertStringToDouble(reader.GetValue(12)?.ToString());

                                                tool.TopThread = reader.GetValue(14)?.ToString();

                                                tool.BottomTread = reader.GetValue(15)?.ToString();

                                                tool.ArrivalDate = CC.ConvertStringToDateTimeString(reader.GetValue(20)?.ToString());

                                                tool.From = reader.GetValue(21)?.ToString();

                                                tool.DepartureDate = CC.ConvertStringToDateTimeString(reader.GetValue(22)?.ToString());

                                                tool.To = reader.GetValue(23)?.ToString();

                                                tool.Status = reader.GetValue(24)?.ToString();

                                                tool.Comment = reader.GetValue(29)?.ToString();

                                                tool.ActiveId = reader.GetValue(79)?.ToString();

                                                if (!String.IsNullOrEmpty(tool.Name)) job.Tools.Add(tool);
                                            }
                                            rowIndex++;
                                        }
                                    }

                                    if (reader.Name == "Телеметрия")
                                    {
                                        rowIndex = 0;

                                        while (reader.Read() && rowIndex < 150)
                                        {
                                            if (rowIndex > 10)
                                            {
                                                Tool tool = new Tool();

                                                tool.Size = reader.GetValue(0)?.ToString();

                                                tool.ToolClass = reader.GetValue(1)?.ToString();

                                                tool.Owner = reader.GetValue(2)?.ToString();

                                                tool.Name = reader.GetValue(3)?.ToString();

                                                tool.SerialNumber = reader.GetValue(4)?.ToString();

                                                tool.Art = reader.GetValue(5)?.ToString();

                                                tool.Pasport = reader.GetValue(6)?.ToString();

                                                tool.InspectionDate = CC.ConvertStringToDateTimeString(reader.GetValue(7)?.ToString());

                                                tool.Days = CC.ConvertStringToInt(reader.GetValue(8)?.ToString());

                                                tool.CircTimeAfterInspection = CC.ConvertStringToDouble(reader.GetValue(10)?.ToString());

                                                tool.CircTime = CC.ConvertStringToDouble(reader.GetValue(12)?.ToString());

                                                tool.TopThread = reader.GetValue(14)?.ToString();

                                                tool.BottomTread = reader.GetValue(15)?.ToString();

                                                tool.ArrivalDate = CC.ConvertStringToDateTimeString(reader.GetValue(20)?.ToString());

                                                tool.From = reader.GetValue(21)?.ToString();

                                                tool.DepartureDate = CC.ConvertStringToDateTimeString(reader.GetValue(22)?.ToString());

                                                tool.To = reader.GetValue(23)?.ToString();

                                                tool.Status = reader.GetValue(24)?.ToString();

                                                tool.Comment = reader.GetValue(29)?.ToString();

                                                tool.Battery = CC.ConvertStringToDouble(reader.GetValue(33)?.ToString());

                                                tool.ActiveId = reader.GetValue(79)?.ToString();

                                                if (!String.IsNullOrEmpty(tool.Name)) job.Tools.Add(tool);

                                            }
                                            rowIndex++;
                                        }
                                    }

                                    if (reader.Name == "БХ_НКС"
                                        || reader.Name == "БТС_НКС"
                                        || reader.Name == "СИБ_НКС"
                                        || reader.Name == "АПС_НКС"
                                        || reader.Name == "Энергия_НКС"
                                        || reader.Name == "АКСЛ_НКС"
                                        || reader.Name == "Инструментальный ящик"
                                        || reader.Name == "ИТ_ВАГОНЫ"
                                        || reader.Name == "Ящик_ИИИ")
                                    {
                                        rowIndex = 0;

                                        Kit kit = new();

                                        while (reader.Read() && rowIndex < 400)
                                        {
                                            if (rowIndex == 6)
                                            {
                                                kit.Art = reader.GetValue(1)?.ToString();

                                                kit.Name = reader.Name;

                                                kit.QuantityNorm = reader.GetValue(3)?.ToString();

                                                kit.QuantityFact = reader.GetValue(4)?.ToString();

                                                kit.SerialNumber = reader.GetValue(5)?.ToString();

                                                kit.Status = reader.GetValue(6)?.ToString();

                                                kit.Comment = reader.GetValue(7)?.ToString();

                                                kit.ArrivalDate = CC.ConvertStringToDateTimeString(reader.GetValue(8)?.ToString());

                                                kit.From = reader.GetValue(9)?.ToString();

                                                kit.DepartureDate = CC.ConvertStringToDateTimeString(reader.GetValue(10)?.ToString());

                                                kit.To = reader.GetValue(11)?.ToString();

                                                kit.InspectionDate = CC.ConvertStringToDateTimeString(reader.GetValue(13)?.ToString());

                                                kit.ActiveId = reader.GetValue(18)?.ToString();
                                            }
                                            else if (rowIndex > 7)
                                            {
                                                KitTool kitTool = new();

                                                kitTool.Art = reader.GetValue(1)?.ToString();

                                                kitTool.Name = reader.GetValue(2)?.ToString();

                                                kitTool.QuantityNorm = reader.GetValue(3)?.ToString();

                                                kitTool.QuantityFact = reader.GetValue(4)?.ToString();

                                                kitTool.SerialNumber = reader.GetValue(5)?.ToString();

                                                kitTool.Status = reader.GetValue(6)?.ToString();

                                                kitTool.Comment = reader.GetValue(7)?.ToString();

                                                kitTool.ArrivalDate = CC.ConvertStringToDateTimeString(reader.GetValue(8)?.ToString());

                                                kitTool.From = reader.GetValue(9)?.ToString();

                                                kitTool.DepartureDate = CC.ConvertStringToDateTimeString(reader.GetValue(10)?.ToString());

                                                kitTool.To = reader.GetValue(11)?.ToString();

                                                kitTool.InspectionDate = CC.ConvertStringToDateTimeString(reader.GetValue(13)?.ToString());

                                                if (!String.IsNullOrEmpty(kitTool.Name))
                                                    kit.KitTools.Add(kitTool);
                                            }
                                            rowIndex++;
                                        }
                                        if (!String.IsNullOrEmpty(kit.QuantityFact))
                                        {
                                            job.Kits.Add(kit);
                                        }
                                    }
                                }
                                while (reader.NextResult());
                            }
                        }

                        jobs.Add(job);
                    }

                    progress.Report(((i + 1) * 100 / files.Length));
                }
                await context.Jobs.AddRangeAsync(jobs);
                await context.SaveChangesAsync();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return await Task.FromResult(ex.Message);
            }


            return await Task.FromResult("Данные загружены.");
        }

        //Private methods
        private void CreateToolsPivotTable(string filePath, List<Tool> tools, List<Kit> kits, List<KitTool> kitTools)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo fileInfo = new FileInfo("Templates\\PivotTable.xlsm");
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                AddToolsToTable(package, tools, "Оборудование");

                List<Tool> toolTempList = tools
                    .Where(t => t.ToolClass != null ? t.ToolClass.ToLower().Contains("взд") : false)
                    .ToList();
                AddToolsToTable(package, toolTempList, "ВЗД");

                toolTempList = tools
                    .Where(t => t.ToolClass != null ? t.ToolClass.ToLower().Contains("яс") : false)
                    .ToList();
                AddToolsToTable(package, toolTempList, "Яс");

                toolTempList = tools
                   .Where(t => t.ToolClass != null ? t.ToolClass.ToLower().Contains("рус") : false)
                   .ToList();
                AddToolsToTable(package, toolTempList, "ATK");

                toolTempList = tools
                   .Where(t => t.Name != null ? t.Name.Contains("BCPM") : false)
                   .ToList();
                AddToolsToTable(package, toolTempList, "BCPM");

                toolTempList = tools
                   .Where(t => t.Name != null ? t.Name.ToLower().Contains("ontrak") : false)
                   .ToList();
                AddToolsToTable(package, toolTempList, "OTK");

                toolTempList = tools
                   .Where(t => (t.Name != null ? t.Name.ToLower().Contains("snd") : false)
                   || (t.Name != null ? t.Name.ToLower().Contains("ord") : false)
                   || (t.Name != null ? t.Name.ToLower().Contains("ccn") : false)
                   || (t.Name != null ? t.Name.ToLower().Contains("сcn") : false)
                   || (t.Name != null ? t.Name.ToLower().Contains("cсn") : false)
                   || (t.Name != null ? t.Name.ToLower().Contains("ссn") : false))
                   .ToList();
                AddToolsToTable(package, toolTempList, "LTK");

                toolTempList = tools
                   .Where(t => t.Name != null ? t.Name.ToLower().Contains("navigamma") : false)
                   .ToList();
                AddToolsToTable(package, toolTempList, "NAVIG");

                toolTempList = tools
                  .Where(t => t.Name != null ? t.Name.ToLower().Contains("navigamma") : false)
                  .ToList();
                AddToolsToTable(package, toolTempList, "NAVIG");

                toolTempList = tools
                  .Where(t => t.Name != null ? t.Name.ToLower().Contains("abpa") : false)
                  .ToList();
                AddToolsToTable(package, toolTempList, "ABPA");

                toolTempList = tools
                  .Where(t => ((t.ToolClass != null ? t.ToolClass.ToLower().Contains("тс_aps") : false)
                  || (t.ToolClass != null ? t.ToolClass.ToLower().Contains("переводник_aps") : false))
                  && ((t.Name != null ? t.Name.ToLower().Contains("battery") : false)
                  || (t.Name != null ? t.Name.ToLower().Contains("pulser") : false)
                  || (t.Name != null ? t.Name.ToLower().Contains("батарей") : false)
                  || (t.Name != null ? t.Name.ToLower().Contains("гамма-модуль") : false)
                  || (t.Name != null ? t.Name.ToLower().Contains("пульсатор") : false)
                  || (t.Name != null ? t.Name.ToLower().Contains("циркуляционный") : false)))
                  .ToList();
                AddToolsToTable(package, toolTempList, "APS");

                toolTempList = tools
                  .Where(t => (t.Name != null ? t.Name.ToLower().Contains("резистив") : false)
                  && (t.ToolClass != null ? t.ToolClass.ToLower().Contains("тс_aps") : false))
                  .ToList();
                AddToolsToTable(package, toolTempList, "WPR");

                toolTempList = tools
                  .Where(t => (t.Name != null ? t.Name.ToLower().Contains("каротаж lwd") : false)
                  || (t.Name != null ? t.Name.ToLower().Contains("нейтрон") : false)
                  && (t.ToolClass != null ? t.ToolClass.ToLower().Contains("энергия_тс") : false))
                  .ToList();
                AddToolsToTable(package, toolTempList, "ERG");

                toolTempList = tools
                 .Where(t => (t.ToolClass != null ? t.ToolClass.ToLower().Contains("тс_ткш") : false))
                 .ToList();
                AddToolsToTable(package, toolTempList, "SIB");

                toolTempList = tools
                 .Where(t => (t.ToolClass != null ? t.ToolClass.ToLower().Contains("тс_бтс") : false)
                 || (t.Name != null ? t.Name.ToLower().Contains("нубт 8.25") : false))
                 .ToList();
                AddToolsToTable(package, toolTempList, "BTS");

                toolTempList = tools
                 .Where(t => (t.ToolClass != null ? t.ToolClass.ToLower() == "цлс" : false))
                 .ToList();
                AddToolsToTable(package, toolTempList, "КЛС");

                toolTempList = tools
                 .Where(t => (t.Name != null ? t.Name.ToLower().Contains("подъемный патрубок") : false))
                 .ToList();
                AddToolsToTable(package, toolTempList, "LiftSub");

                toolTempList = tools
                 .Where(t => (t.ToolClass != null ? t.ToolClass.ToLower().Contains("ибн") : false)
                 || (t.Name != null ? t.Name.ToLower().Contains("иги") : false)
                 || (t.Name != null ? t.Name.ToLower().Contains("источникодержат") : false))
                 .ToList();
                AddToolsToTable(package, toolTempList, "ИИИ");

                toolTempList = tools
                 .Where(t => (t.Name != null ? t.Name.ToLower().Contains("дозимет") : false))
                 .ToList();
                AddToolsToTable(package, toolTempList, "Дозиметры");

                List<Kit> kitTempList = kits
                    .Where(k => (k.Name != null ? k.Name.ToLower().Contains("нкс") || k.Name.ToLower().Contains("иии") : false)
                    && (k.SerialNumber != null))
                    .ToList();
                AddKitsToTable(package, kitTempList, "KITS");

                List<KitTool> kitToolTempList = kitTools
                    .Where(k => (k.Kit.Name != null ? k.Kit.Name.ToLower().Contains("нкс") || k.Kit.Name.ToLower().Contains("иии") : false))
                    .ToList();
                AddKitToolsToTable(package, kitToolTempList, "Содержимое_KITS");

                kitToolTempList = kitTools
                    .Where(k => !k.QuantityFact.IsNullOrEmpty() || k.QuantityFact != "0" || !k.SerialNumber.IsNullOrEmpty() || !k.Status.IsNullOrEmpty())
                    .Where(k => (k.Kit.Name != null ? k.Kit.Name.ToLower().Contains("вагоны") || k.Kit.Name.ToLower().Contains("инструмент") : false))
                    .ToList();
                AddKitToolsToTable(package, kitToolTempList, "ИТ_Вагоны");

                toolTempList = tools
                    .Where(t => !t.ActiveId.IsNullOrEmpty())
                    .ToList()
                    .GroupBy(t => t.ActiveId)
                    .Where(g => g.Count() > 1)
                    .SelectMany(g => g)
                    .ToList();
                AddToolsToTable(package, toolTempList, "Дубликаты_Tools");

                kitTempList = kits
                    .Where(t => !t.ActiveId.IsNullOrEmpty())
                    .ToList()
                    .GroupBy(t => t.ActiveId)
                    .Where(g => g.Count() > 1)
                    .SelectMany(g => g)
                    .ToList();
                AddKitsToTable(package, kitTempList, "Дубликаты_KITS");

                package.SaveAs(filePath);
            }
        }
        private void AddToolsToTable(ExcelPackage package, List<Tool> tools, string sheetName)
        {
            var worksheet = package.Workbook.Worksheets[sheetName];

            if (worksheet != null)
            {
                var table = worksheet.Tables[$"УТ_{sheetName}"];

                if (table != null)
                {
                    CleanExcel(worksheet, table);

                    int row = table.Address.End.Row;

                    foreach (var tool in tools)
                    {
                        worksheet.Cells[row, 1].Value = tool.Job.Field;
                        worksheet.Cells[row, 2].Value = tool.Job.Pad;
                        worksheet.Cells[row, 3].Value = tool.Job.Well;
                        worksheet.Cells[row, 4].Value = tool.Size;
                        worksheet.Cells[row, 5].Value = tool.ToolClass;
                        worksheet.Cells[row, 6].Value = tool.Owner;
                        worksheet.Cells[row, 7].Value = tool.Name;
                        worksheet.Cells[row, 8].Value = tool.SerialNumber;
                        worksheet.Cells[row, 9].Value = tool.Art;
                        worksheet.Cells[row, 10].Value = tool.Pasport;
                        worksheet.Cells[row, 11].Value = tool.CircTimeAfterInspection;
                        worksheet.Cells[row, 12].Value = tool.CircTime;
                        worksheet.Cells[row, 13].Value = tool.InspectionDate;
                        worksheet.Cells[row, 14].Value = tool.Days;
                        worksheet.Cells[row, 15].Value = tool.Status;
                        worksheet.Cells[row, 16].Value = tool.ArrivalDate;
                        worksheet.Cells[row, 17].Value = tool.From;
                        worksheet.Cells[row, 18].Value = tool.DepartureDate;
                        worksheet.Cells[row, 19].Value = tool.To;
                        worksheet.Cells[row, 20].Value = tool.CollorField;
                        worksheet.Cells[row, 21].Value = tool.Comment;
                        worksheet.Cells[row, 22].Value = tool.TopThread;
                        worksheet.Cells[row, 23].Value = tool.BottomTread;
                        worksheet.Cells[row, 24].Value = tool.Battery;
                        worksheet.Cells[row, 25].Value = tool.ActiveId;
                        worksheet.Cells[row, 26].Value = tool.Job.FilePath;
                        table.AddRow();
                        row++;
                    }
                }
            }
        }
        private void AddKitsToTable(ExcelPackage package, List<Kit> kits, string sheetName)
        {
            var worksheet = package.Workbook.Worksheets[sheetName];

            if (worksheet != null)
            {
                var table = worksheet.Tables[$"УТ_{sheetName}"];

                if (table != null)
                {
                    CleanExcel(worksheet, table);

                    int row = table.Address.End.Row;

                    foreach (var kit in kits)
                    {
                        worksheet.Cells[row, 1].Value = kit.Job.Field;
                        worksheet.Cells[row, 2].Value = kit.Job.Pad;
                        worksheet.Cells[row, 3].Value = kit.Job.Well;
                        worksheet.Cells[row, 4].Value = kit.Name;
                        worksheet.Cells[row, 5].Value = kit.SerialNumber;
                        worksheet.Cells[row, 6].Value = kit.QuantityFact;
                        worksheet.Cells[row, 7].Value = kit.Status;
                        worksheet.Cells[row, 8].Value = kit.Comment;
                        worksheet.Cells[row, 9].Value = kit.ArrivalDate;
                        worksheet.Cells[row, 10].Value = kit.From;
                        worksheet.Cells[row, 11].Value = kit.DepartureDate;
                        worksheet.Cells[row, 12].Value = kit.To;
                        worksheet.Cells[row, 13].Value = kit.ActiveId;
                        worksheet.Cells[row, 14].Value = kit.Job.FilePath;
                        table.AddRow();
                        row++;
                    }
                }
            }
        }
        private void AddKitToolsToTable(ExcelPackage package, List<KitTool> kitTools, string sheetName)
        {
            var worksheet = package.Workbook.Worksheets[sheetName];

            if (worksheet != null)
            {
                var table = worksheet.Tables[$"УТ_{sheetName}"];

                if (table != null)
                {
                    CleanExcel(worksheet, table);

                    int row = table.Address.End.Row;

                    foreach (var kitTool in kitTools)
                    {
                        worksheet.Cells[row, 1].Value = kitTool.Kit.Job.Field;
                        worksheet.Cells[row, 2].Value = kitTool.Kit.Job.Pad;
                        worksheet.Cells[row, 3].Value = kitTool.Kit.Job.Well;
                        worksheet.Cells[row, 4].Value = kitTool.Kit.Name;
                        worksheet.Cells[row, 5].Value = kitTool.Name;
                        worksheet.Cells[row, 6].Value = kitTool.SerialNumber;
                        worksheet.Cells[row, 7].Value = kitTool.QuantityFact;
                        worksheet.Cells[row, 8].Value = kitTool.Status;
                        worksheet.Cells[row, 9].Value = kitTool.Comment;
                        worksheet.Cells[row, 10].Value = kitTool.ArrivalDate;
                        worksheet.Cells[row, 11].Value = kitTool.From;
                        worksheet.Cells[row, 12].Value = kitTool.DepartureDate;
                        worksheet.Cells[row, 13].Value = kitTool.To;
                        worksheet.Cells[row, 14].Value = kitTool.Kit.Job.FilePath;
                        table.AddRow();
                        row++;
                    }
                }
            }
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
    }
}
