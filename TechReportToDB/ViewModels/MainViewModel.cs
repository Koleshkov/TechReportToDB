using CommunityToolkit.Mvvm.ComponentModel;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.ViewModels
{
    internal partial class MainViewModel : ViewModelBase
    {
        //Props
        [ObservableProperty]
        private ViewModelBase currentPageModel = new();

        [ObservableProperty]
        private string windowState = "Normal";

    }
}
