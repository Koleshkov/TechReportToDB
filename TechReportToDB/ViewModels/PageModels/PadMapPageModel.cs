using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.DependencyInjection;
using System.ComponentModel;
using System.Windows.Data;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Repos;
using TechReportToDB.ViewModels.Base;
using System.Diagnostics;
using Microsoft.EntityFrameworkCore;
using GMap.NET.WindowsPresentation;
using GMap.NET.MapProviders;
using TechReportToDB.Views.CustomControls;

namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class PadMapPageModel : PageModelBase
    {
        private readonly IRepo<Job> jobRepo;
        private readonly INavigationService navigationService;

        public ICollectionView FiltredJobList { get; }

        [ObservableProperty] List<Job> jobList = new();
        [ObservableProperty] string filterJob = "";
        [ObservableProperty] Job selectedJob;

        public Action<IEnumerable<Job>> UpdateMapCallback { get; set; }
        public Action<Job> SelectPadMapCallback { get; set; }

        public PadMapPageModel(IRepo<Job> jobRepo, INavigationService navigationService)
        {
            this.jobRepo = jobRepo;
            this.navigationService = navigationService;

            JobList = jobRepo.List.Include(t => t.Tools)
                                  .OrderBy(f => f.FieldTeam)
                                  .Include(d => d.DDs)
                                  .Include(m => m.MWDs)
                                  .ToList();

            SelectedJob = JobList.FirstOrDefault() ?? new Job();
            FiltredJobList = CollectionViewSource.GetDefaultView(JobList);
            FiltredJobList.Filter = FilterBySearchText;

            UpdateMapCallback?.Invoke(JobList);
            SelectPadMapCallback = (s) => { };
        }

        public async override Task InitializeAsync()
        {
            var mvm = ServiceLocator.ServiceProvider.GetRequiredService<MainViewModel>();
            mvm.WindowState = "Maximized";
            await base.InitializeAsync();
        }

        partial void OnFilterJobChanged(string value) => FiltredJobList.Refresh();


        partial void OnSelectedJobChanging(Job value)
        {
            SelectPadMapCallback?.Invoke(value);
            UpdateMapCallback?.Invoke(JobList);
        }

        private bool FilterBySearchText(object item)
        {
            if (item is Job job)
            {
                if (string.IsNullOrWhiteSpace(FilterJob)) return true;

                string filter = FilterJob.ToLower();
                return job.FilterName.ToLower().Contains(filter) ||
                       job.FieldTeam?.ToLower().Contains(filter) == true ||
                       job.Tools.Any(tool => tool.Name.ToLower().Contains(filter));
            }
            return false;
        }

        [RelayCommand]
        public async Task GoToDailyReportToDbPage()
        {
            try
            {
                await navigationService.NavigateToAsync<DailyReportToDbPageModel>();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка навигации: {ex.Message}");
            }
        }
    }
}
