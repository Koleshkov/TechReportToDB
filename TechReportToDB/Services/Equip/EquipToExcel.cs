using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services.Repos;

namespace TechReportToDB.Services.Equip
{
    internal class EquipToExcel : IEquipToExcel
    {
        private readonly IRepo<Job> jobRepo;

        public EquipToExcel(IRepo<Job> jobRepo)
        {
            this.jobRepo = jobRepo;
        }

        public async Task ExportToExcel(string filePath)
        {
            try
            {
            var jobs = jobRepo.List.Include(j => j.Tools).ToList();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo fileInfo = new FileInfo("Templates\\Equipment.xlsx");


            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Оборудование"];

                if (worksheet != null)
                {
                    var table = worksheet.Tables[$"УТ_Оборудование"];

                    CleanExcel(worksheet, table);



                    foreach (var job in jobs)
                    {
                        int row = table.Address.End.Row;
                        worksheet.Cells[row, 1].Value = job.Field + "_" + job.Pad;

                        worksheet.Cells[row, 2].Value = job.Tools.Where(t => 
                        (t.ToolClass != null ? t.ToolClass.ToLower().Contains("взд") : false) 
                        && (t.Size != null ? t.Size[0] == '8' : false) 
                        && !(t.Status!=null? t.Status.ToLower().Contains("вывезено"):true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                        worksheet.Cells[row, 3].Value = job.Tools.Where(t =>
                        (t.ToolClass != null ? t.ToolClass.ToLower().Contains("взд") : false)
                        && (t.Size != null ? t.Size[0] == '6' : false)
                        && !(t.Status != null ? t.Status.ToLower().Contains("вывезено") : true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                        worksheet.Cells[row, 4].Value = job.Tools.Where(t =>
                        (t.ToolClass != null ? t.ToolClass.ToLower().Contains("взд") : false)
                        && (t.Size != null ? t.Size[0] == '4' : false)
                        && !(t.Status != null ? t.Status.ToLower().Contains("вывезено") : true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                        worksheet.Cells[row, 5].Value = job.Tools.Where(t => 
                        ((t.Name != null ? t.Name.ToLower().Contains("мэс") : false) 
                        || (t.Name != null ? t.Name.ToLower().Contains("инклин") : false) 
                        || (t.Name != null ? t.Name.ToLower().Contains("ontrak") : false) 
                        || (t.Name != null ? t.Name.ToLower().Contains("navigam") : false)) 
                        && (t.Size != null ? t.Size[0] == '8' : false)
                        && !(t.Status != null ? t.Status.ToLower().Contains("вывезено") : true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                        worksheet.Cells[row, 6].Value = job.Tools.Where(t => 
                        ((t.Name != null ? t.Name.ToLower().Contains("мэс") : false) 
                        || (t.Name != null ? t.Name.ToLower().Contains("инклин") : false) 
                        || (t.Name != null ? t.Name.ToLower().Contains("ontrak") : false) 
                        || (t.Name != null ? t.Name.ToLower().Contains("navigam") : false)) 
                        && (t.Size != null ? t.Size[0] == '6' : false)
                        && !(t.Status != null ? t.Status.ToLower().Contains("вывезено") : true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                        worksheet.Cells[row, 7].Value = job.Tools.Where(t => 
                        ((t.Name != null ? t.Name.ToLower().Contains("мэс") : false) 
                        || (t.Name != null ? t.Name.ToLower().Contains("инклин") : false) 
                        || (t.Name != null ? t.Name.ToLower().Contains("ontrak") : false) 
                        || (t.Name != null ? t.Name.ToLower().Contains("navigam") : false)) 
                        && (t.Size != null ? t.Size[0] == '4' : false)
                        && !(t.Status != null ? t.Status.ToLower().Contains("вывезено") : true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                        worksheet.Cells[row, 8].Value = job.Tools.Where(t => 
                        (t.ToolClass != null ? t.ToolClass.ToLower().Contains("яс") : false) 
                        && (t.Size != null ? t.Size[0] == '6' : false)
                        && !(t.Status != null ? t.Status.ToLower().Contains("вывезено") : true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                        worksheet.Cells[row, 9].Value = job.Tools.Where(t => 
                        (t.ToolClass != null ? t.ToolClass.ToLower().Contains("яс") : false) 
                        && (t.Size != null ? t.Size[0] == '4' : false)
                        && !(t.Status != null ? t.Status.ToLower().Contains("вывезено") : true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                        worksheet.Cells[row, 10].Value = job.Tools.Where(t =>
                        (t.ToolClass != null ? t.ToolClass.ToLower().Contains("рус") : false)
                        && (t.Size != null ? t.Size[0] == '6' : false)
                        && !(t.Status != null ? t.Status.ToLower().Contains("вывезено") : true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                        worksheet.Cells[row, 11].Value = job.Tools.Where(t =>
                        (t.ToolClass != null ? t.ToolClass.ToLower().Contains("рус") : false)
                        && (t.Size != null ? t.Size[0] == '4' : false)
                        && !(t.Status != null ? t.Status.ToLower().Contains("вывезено") : true)
                        && !(t.Status != null ? t.Status.ToLower().Contains("оставлено") : true)).Count();

                            table.InsertRow(row);
                    }
                    package.SaveAs(filePath);
                }
            }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            await Task.CompletedTask;
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
