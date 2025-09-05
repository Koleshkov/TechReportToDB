using CommunityToolkit.Mvvm.ComponentModel;

namespace TechReportToDB.ViewModels.Base
{
    internal partial class PageModelBase : ObservableObject
    {

        [ObservableProperty]
        private string title = string.Empty;

        [ObservableProperty]
        private string info = "Empty";

        public virtual Task InitializeAsync()
        {
            return Task.CompletedTask;
        }
    }
}
