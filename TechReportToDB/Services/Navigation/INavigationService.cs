using System.Threading.Tasks;
using System.Windows;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.Services.Navigation
{
    internal interface INavigationService
    {
        Task NavigateToAsync<TPageModel>() where TPageModel : PageModelBase;

        Task OpenWindowAsync<TWindow, TWindowModel>() where TWindow : Window, new() where TWindowModel : WindowModelBase;

        Task CloseWindowAsync<TWindowModel>() where TWindowModel : WindowModelBase;
    }
}
