

using System.Text.Json.Serialization;

namespace TechReportToDB.Data.Entities
{
    class BaseEntity
    {
        [JsonIgnore]
        public Guid? Id { get; set; }
        [JsonIgnore]
        public DateTime TimeStamp { get; set; } = DateTime.Now;
    }
}
