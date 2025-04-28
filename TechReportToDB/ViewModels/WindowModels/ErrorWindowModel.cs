using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TechReportToDB.Services.Navigation;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.ViewModels.WindowModels
{
    internal partial class ErrorWindowModel:WindowModelBase
    {
        [ObservableProperty]
        private string? message = "";

        public ErrorWindowModel(INavigationService navigationService) : base(navigationService)
        {
        }

        [RelayCommand]
        public override async Task CloseWindowAsync() =>
            await navigationService.CloseWindowAsync<ErrorWindowModel>();
    }
}
