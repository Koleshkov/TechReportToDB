using Microsoft.Extensions.DependencyInjection;
using System.Windows;
using TechReportToDB.ViewModels;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.Services.Navigation
{
    internal class NavigationService : INavigationService
    {
        private readonly Dictionary<WindowModelBase, Window> windowDictionary;

        public NavigationService()
        {
            windowDictionary = new();
        }

        public  Task NavigateToAsync<TPageModel>() where TPageModel : PageModelBase
        {
            MainViewModel mainViewModel = ServiceLocator.ServiceProvider.GetRequiredService<MainViewModel>();

            mainViewModel.CurrentPageModel = ServiceLocator.ServiceProvider.GetRequiredService<TPageModel>();

            return mainViewModel.CurrentPageModel.InitializeAsync();
        }

        public async Task OpenWindowAsync<TWindow, TWindowModel>() where TWindow : Window, new() where TWindowModel : WindowModelBase
        {
            var windowModel = ServiceLocator.ServiceProvider.GetRequiredService<TWindowModel>();
            Window window;
            if (windowDictionary.Keys.Contains(windowModel))
            {
                window = windowDictionary[windowModel];

                window.Close();

                windowDictionary.Remove(windowModel);
            }

            window = new TWindow()
            {
                DataContext = windowModel
            };

            windowDictionary.Add(windowModel, window);

            window.Show();

            await Task.CompletedTask;
        }

        public async Task CloseWindowAsync<TWindowModel>() where TWindowModel : WindowModelBase
        {
            var windowModel = ServiceLocator.ServiceProvider.GetRequiredService<TWindowModel>();

            var window = windowDictionary[windowModel];

            if (window != null)
            {
                window.Close();
            }

            windowDictionary.Remove(windowModel);

            await Task.CompletedTask;
        }
    }
}
