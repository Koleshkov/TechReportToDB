using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechReportToDB.Services.Navigation;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class DateTimePageModel : PageModelBase
    {
        private readonly INavigationService navigationService;

        [ObservableProperty]
        private DateTime dateOfReport = DateTime.Now;

        [ObservableProperty]
        private int selectedStartTime = 0;

        [ObservableProperty]
        private int selectedEndTime = 9;

        public DateTimePageModel(INavigationService navigationService)
        {
            this.navigationService = navigationService;
        }

        public Func<DateTime, int, int, Task>? SelectAction { get; set; }

        [RelayCommand]
        public async Task Select() 
        {
            SelectAction?.Invoke(DateOfReport, SelectedStartTime, SelectedEndTime);
            await GoToDailyReportToDbPage();
        } 

        [RelayCommand]
        public async Task GoToDailyReportToDbPage() => await navigationService.NavigateToAsync<DailyReportToDbPageModel>();
    }
}
