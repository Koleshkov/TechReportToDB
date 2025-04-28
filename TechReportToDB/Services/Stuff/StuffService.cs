using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechReportToDB.Data.Entities;
using TechReportToDB.Data.Models.Workstations;

namespace TechReportToDB.Services.Stuff
{
    internal class StuffService : IStuffFService
    {

        public async Task ExportToExcel(string filePath, IEnumerable<Person> persons)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                FileInfo fileInfo = new FileInfo("Templates\\StuffTemplate.xlsx");

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Персонал"];
                    if (worksheet != null)
                    {
                        var table = worksheet.Tables[$"УТ_Персонал"];

                        if (table != null)
                        {
                            CleanExcel(worksheet, table);
                            foreach (var person in persons)
                            {
                                int row = table.Address.End.Row;
                                worksheet.Cells[row, 1].Value = person?.Job?.FieldTeam;
                                worksheet.Cells[row, 2].Value = person?.Job?.Field;
                                worksheet.Cells[row, 3].Value = person?.Job?.Pad;
                                worksheet.Cells[row, 4].Value = person?.Job?.Well;
                                worksheet.Cells[row, 5].Value = person?.Job?.Type;
                                worksheet.Cells[row, 6].Value = person?.Job?.Phone;
                                worksheet.Cells[row, 7].Value = person?.Name;
                                worksheet.Cells[row, 8].Value = person?.Position;
                                worksheet.Cells[row, 9].Value = person?.DateOfJob;
                                worksheet.Cells[row, 10].Value = person?.Phone;
                                table.AddRow();
                            }
                            worksheet.DeleteRow(table.Address.End.Row);
                        }
                        package.SaveAs(filePath);
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                await Task.CompletedTask;
            }
            
            await Task.CompletedTask;
        }

        private void CleanExcel(ExcelWorksheet worksheet, ExcelTable table)
        {
            int startRow = table.Address.Start.Row + 1;
            int endRow = table.Address.End.Row;

            for (int i = endRow - startRow; i > 1; i--)
            {
                table.DeleteRow(i - 1);
            }
            worksheet.Cells["A2:AA2"].Value = "";
        }
    }
}
