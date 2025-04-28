using Microsoft.EntityFrameworkCore;
using System.Linq.Expressions;
using TechReportToDB.Data;
using TechReportToDB.Data.Entities;

namespace TechReportToDB.Services.Repos
{
    internal class Repo<T>: IRepo<T> where T : BaseEntity
    {
        private readonly AppDbContext context;
        private readonly DbSet<T> table;
        public Repo(AppDbContext context)
        {
            this.context = context;
            table = context.Set<T>();
        }

        public IQueryable<T> List =>
            context.Set<T>().AsQueryable();

        public async Task<T?> GetByIdAsync(int id) =>
            await context.Set<T>().FindAsync(id);

        public async Task AddAsync(T entity)
        {
            await context.Set<T>().AddAsync(entity);
            await context.SaveChangesAsync();
        }

        public async Task AddWithoutSavingAsync(T entity)
        {
            await context.Set<T>().AddAsync(entity);
        }

        public async Task AddRangeAsync(IEnumerable<T> entity)
        {
            await context.Set<T>().AddRangeAsync(entity);
            await context.SaveChangesAsync();
        }

        public async Task UpdateAsync(T entity)
        {
            context.Set<T>().Update(entity);
            await context.SaveChangesAsync();
        }

        public async Task DeleteAsync(int id)
        {
            var entity = await GetByIdAsync(id);
            if (entity != null)
            {
                context.Set<T>().Remove(entity);
                await context.SaveChangesAsync();
            }
        }

        public async Task<T?> FirstOrDefaultAsync() => 
            await context.Set<T>().FirstOrDefaultAsync();

        public async Task<bool> AnyAsync()=>
             await context.Set<T>().AnyAsync();

        public async Task ClearAllTablesAsync()
        {
            if(await context.Set<T>().AnyAsync())
            {
                await context.Database.EnsureDeletedAsync();

                await context.Database.EnsureCreatedAsync();
            }
        }

        public IEnumerable<T> Find(Expression<Func<T, bool>> predicate)
        {
            return context.Set<T>().Where(predicate).ToList();
        }

        

        public async Task SaveChangesAsync()
        {
            await context.SaveChangesAsync();
        }
    }
}
