

namespace TechReportToDB.Data.Entities
{
    class Base
    {
        public Guid Id { get; set; }

        public DateTime TimeStamp { get; set; } = DateTime.Now;
    }
}
