using CommunityToolkit.Mvvm.ComponentModel;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.ViewModels
{
    internal partial class MainViewModel : ObservableObject
    {
        //Props
        [ObservableProperty]
        private PageModelBase currentPageModel = new();

        [ObservableProperty]
        private string windowState = "Normal";

    }
}
