using System.Linq.Expressions;
using TechReportToDB.Data.Entities;

namespace TechReportToDB.Services.Repos
{
    internal interface IRepo<T> where T :  BaseEntity
    {

        IQueryable<T> List { get; }
        Task<T?> GetByIdAsync(int id);
        Task AddAsync(T entity);
        Task AddWithoutSavingAsync(T entity);
        Task SaveChangesAsync();
        Task AddRangeAsync(IEnumerable<T> entity);
        Task UpdateAsync(T entity);
        Task DeleteAsync(int id);
        Task<T?> FirstOrDefaultAsync();
        Task<bool> AnyAsync();
        Task ClearAllTablesAsync();
        IEnumerable<T> Find(Expression<Func<T, bool>> predicate);
    }
}
