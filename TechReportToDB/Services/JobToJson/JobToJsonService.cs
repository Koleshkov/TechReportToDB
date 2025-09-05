using System.Data.Entity;
using System.IO;
using System.Text.Json;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services.Repos;

namespace TechReportToDB.Services.JobToJson
{
    internal class JobToJsonService : IJobToJsonService
    {
        private readonly IRepo<Job> jobRepo;

        public JobToJsonService(IRepo<Job> jobRepo)
        {
            this.jobRepo = jobRepo;
        }

        public async Task ExportToJson(string filePath)
        {
            try
            {
                var jobList = jobRepo.List.Include(c=>c.Constructions).ToList();

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,  // Форматирование с отступами
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping // Для UTF-8
                };

                // Сериализация в строку
                string json = JsonSerializer.Serialize(jobList, options);

                // Сохранение в файл
                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            await Task.CompletedTask;
        }
    }
}
