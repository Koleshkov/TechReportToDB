

namespace TechReportToDB.Data.Entities
{
    internal class DD:Base
    {
        public Guid JobId { get; set; }

        public Job Job { get; set; } = new();

        public string? Name { get; set; }
    }
}
