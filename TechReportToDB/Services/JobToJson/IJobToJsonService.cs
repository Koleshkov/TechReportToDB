

namespace TechReportToDB.Services.JobToJson
{
    internal interface IJobToJsonService
    {
        Task ExportToJson(string filePath);
    }
}
