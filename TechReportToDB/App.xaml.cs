using Microsoft.Extensions.DependencyInjection;
using System.Globalization;
using System.Windows;
using TechReportToDB.Services;
using TechReportToDB.Services.Navigation;
using TechReportToDB.ViewModels.PageModels;
using TechReportToDB.Views;
using Application = System.Windows.Application;

namespace TechReportToDB
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            ServiceLocator.AddServices();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            

            var navSrev = ServiceLocator.ServiceProvider.GetRequiredService<INavigationService>();
            navSrev.NavigateToAsync<DailyReportToDbPageModel>();
            MainWindow = ServiceLocator.ServiceProvider.GetRequiredService<MainView>();

   

            MainWindow.Show();

            base.OnStartup(e);
        }


    }
}
