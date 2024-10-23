using System.Threading.Tasks;
using System.Windows;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.Services.Navigation
{
    internal interface INavigationService
    {
        Task NavigateToAsync<TPageModel>() where TPageModel : ViewModelBase;

        Task OpenWindowAsync<TWindow, TWindowModel>() where TWindow : Window, new() where TWindowModel : ViewModelBase;

        Task CloseWindowAsync<TWindowModel>() where TWindowModel : ViewModelBase;
    }
}
