using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using TechReportToDB.Data;
using TechReportToDB.Services.ExcelToDb;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Outlook;
using TechReportToDB.ViewModels;
using TechReportToDB.ViewModels.PageModels;
using TechReportToDB.ViewModels.WindowModels;
using TechReportToDB.Views;
using TechReportToDB.Views.Pages;
using TechReportToDB.Views.Windows;

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

            var serverVersion = new MySqlServerVersion(new Version(8, 0, 29));

            services.AddDbContext<AppDbContext>(opt =>
                opt.UseSqlite(Configuration.GetConnectionString("SqlLite")));


            services.AddSingleton<INavigationService, NavigationService>();

            services.AddTransient<IExcelToDbService, ExcelToDbService>();

            services.AddTransient<IOutlookService, OutlookService>();


            #endregion


            #region Add Views with ViewModels

            //Add Windows
            services.AddSingleton<ErrorWindowModel>();

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

            services.AddSingleton<DbManagerPageModel>();
            services.AddSingleton(s => new DBManagerPage
            {
                DataContext = s.GetRequiredService<DbManagerPageModel>()
            });

            services.AddSingleton<SearchByToolsPageModel>();
            services.AddSingleton(s => new SearchByToolsPage
            {
                DataContext = s.GetRequiredService<SearchByToolsPageModel>()
            });

            #endregion

            //Build ServiceProvider
            serviceProvider = services.BuildServiceProvider();
        }
    }
}
