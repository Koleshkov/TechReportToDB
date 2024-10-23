using Microsoft.Extensions.DependencyInjection;
using System.Windows;
using TechReportToDB.Services;
using TechReportToDB.Services.Navigation;
using TechReportToDB.ViewModels.PageModels;
using TechReportToDB.Views;

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

            MainWindow = ServiceLocator.ServiceProvider.GetRequiredService<MainView>();

            navSrev.NavigateToAsync<DailyReportToDbPageModel>();

            MainWindow.Show();

            base.OnStartup(e);
        }
    }
}
