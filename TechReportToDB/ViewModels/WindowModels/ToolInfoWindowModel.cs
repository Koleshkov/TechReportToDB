using CommunityToolkit.Mvvm.ComponentModel;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services.Navigation;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.ViewModels.WindowModels
{
    internal partial class ToolInfoWindowModel : WindowModelBase
    {
        [ObservableProperty]
        private Tool? tool;
        public ToolInfoWindowModel(INavigationService navigationService) : base(navigationService)
        {

        }
    }
}
