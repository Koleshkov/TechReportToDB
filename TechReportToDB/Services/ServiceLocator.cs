using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using TechReportToDB.Data;
using TechReportToDB.Services.DailyReportToDb;
using TechReportToDB.Services.Equip;
using TechReportToDB.Services.JobToJson;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Outlook;
using TechReportToDB.Services.Repos;
using TechReportToDB.Services.ShortReport;
using TechReportToDB.Services.Stuff;
using TechReportToDB.Services.WorkstationReportToDb;
using TechReportToDB.ViewModels;
using TechReportToDB.ViewModels.PageModels;
using TechReportToDB.ViewModels.WindowModels;
using TechReportToDB.Views;
using TechReportToDB.Views.Pages;

namespace TechReportToDB.Services
{
    internal static class ServiceLocator
    {
        private static IServiceProvider? serviceProvider;
        public static IServiceProvider ServiceProvider =>
            serviceProvider ?? throw new Exception("Service provider has not been initialized");

        public static IConfiguration? Configuration { get; private set; }

        public static void AddServices()
        {
            #region Add Configurations
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            Configuration = builder.Build();
            #endregion


            #region Add Services
            IServiceCollection services = new ServiceCollection();

            services.AddDbContext<AppDbContext>(opt =>
                opt.UseSqlite(Configuration.GetConnectionString("SqlLite")));

            services.AddTransient(typeof(IRepo<>), typeof(Repo<>));

            services.AddSingleton<INavigationService, NavigationService>();

            services.AddTransient<IDailyReportToDbService, DailyReportToDbService>();

            services.AddTransient<IOutlookService, OutlookService>();

            services.AddSingleton<IWorkstationReportToDbService, WorkstationReportToDbService>();

            services.AddSingleton<IShortReportService, ShortReportService>();

            services.AddSingleton<IStuffFService, StuffService>();

            services.AddSingleton<IEquipToExcel, EquipToExcel>();
            services.AddSingleton<IJobToJsonService, JobToJsonService>();

            #endregion


            #region Add Views with ViewModels

            //Add Windows
            services.AddSingleton<ErrorWindowModel>();
            services.AddSingleton<PersonWindowModel>();
            services.AddSingleton<ToolInfoWindowModel>();

            //Add MainView
            services.AddSingleton<MainViewModel>();
            services.AddSingleton(s => new MainView
            {
                DataContext = s.GetRequiredService<MainViewModel>()
            });

            //Add Pages
            services.AddSingleton<DailyReportToDbPageModel>();
            services.AddSingleton(s => new DailyReportToDbPage
            {
                DataContext = s.GetRequiredService<DailyReportToDbPageModel>()
            });

            services.AddTransient<DbManagerPageModel>();
            services.AddTransient(s => new DBManagerPage
            {
                DataContext = s.GetRequiredService<DbManagerPageModel>()
            });

            services.AddTransient<SearchByToolsPageModel>();
            services.AddTransient(s => new SearchByToolsPage
            {
                DataContext = s.GetRequiredService<SearchByToolsPageModel>()
            });

            services.AddTransient<StaffPageModel>();
            services.AddTransient(s => new StaffPage
            {
                DataContext = s.GetRequiredService<StaffPageModel>()
            });

            services.AddSingleton<PadMapPageModel>();
            services.AddSingleton(s => new PadMapPage
            {
                DataContext = s.GetRequiredService<PadMapPageModel>()
            });

            services.AddSingleton<DateTimePageModel>();
            services.AddSingleton(s => new DateTimePage
            {
                DataContext = s.GetRequiredService<DateTimePageModel>()
            });

            services.AddSingleton<DateTimeShortPageModel>();
            services.AddSingleton(s => new DateTimeShortPage
            {
                DataContext = s.GetRequiredService<DateTimeShortPageModel>()
            });

            #endregion

            //Build ServiceProvider
            serviceProvider = services.BuildServiceProvider();
        }
    }
}
