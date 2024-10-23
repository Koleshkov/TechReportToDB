using TechReportToDB.Services.Navigation;

namespace TechReportToDB.ViewModels.Base
{
    internal class WindowModelBase : ViewModelBase
    {
        protected readonly INavigationService navigationService;

        public WindowModelBase(INavigationService navigationService)
        {
            this.navigationService = navigationService;
        }

        public virtual Task CloseWindowAsync() => Task.CompletedTask;
    }
}
