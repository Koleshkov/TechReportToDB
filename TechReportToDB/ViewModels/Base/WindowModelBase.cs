using CommunityToolkit.Mvvm.ComponentModel;
using TechReportToDB.Services.Navigation;

namespace TechReportToDB.ViewModels.Base
{
    internal class WindowModelBase : ObservableObject
    {
        protected readonly INavigationService navigationService;

        public WindowModelBase(INavigationService navigationService)
        {
            this.navigationService = navigationService;
        }

        public virtual Task CloseWindowAsync() => Task.CompletedTask;
    }
}
