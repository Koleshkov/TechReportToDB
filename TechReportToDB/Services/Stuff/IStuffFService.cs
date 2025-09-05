using TechReportToDB.Data.Entities;

namespace TechReportToDB.Services.Stuff
{
    internal interface IStuffFService
    {
        Task ExportToExcel(string filePath, IEnumerable<Person> persons);
    }
}
