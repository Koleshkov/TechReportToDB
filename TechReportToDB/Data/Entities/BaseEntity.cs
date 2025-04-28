

namespace TechReportToDB.Data.Entities
{
    class BaseEntity
    {
        public Guid? Id { get; set; }

        public DateTime TimeStamp { get; set; } = DateTime.Now;
    }
}
