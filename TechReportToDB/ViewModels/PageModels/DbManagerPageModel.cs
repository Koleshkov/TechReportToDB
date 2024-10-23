using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using TechReportToDB.Data;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services;
using TechReportToDB.Services.Navigation;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class DbManagerPageModel : ViewModelBase
    {
        private readonly INavigationService navigationService;
        private readonly AppDbContext context;

        [ObservableProperty]
        private string filterJob = "";
        
        [ObservableProperty]
        private ObservableCollection<Job> jobList = new();

        public ICollectionView FiltredJobList { get; }

        [ObservableProperty]
        private ObservableCollection<Tool> bHATools = new();

        [ObservableProperty]
        private ObservableCollection<Tool> backupTools = new();

        [ObservableProperty]
        private ObservableCollection<Tool> forInspectionTools = new();

        [ObservableProperty]
        private ObservableCollection<Tool> lostTools = new();

        [ObservableProperty]
        private ObservableCollection<Tool> depTools = new();

        [ObservableProperty]
        private ObservableCollection<DD> dDList = new();

        [ObservableProperty]
        private ObservableCollection<MWD> mWDList = new();

        [ObservableProperty]
        private Job selectedJob = new();
        
        public DbManagerPageModel(INavigationService navigationService, AppDbContext context)
        {
            this.navigationService = navigationService;
            this.context = context;
            JobList = new ObservableCollection<Job>(context.Jobs.Include(t => t.Tools).OrderBy(f=>f.FieldTeam).Include(d=>d.DDs).Include(m=>m.MWDs).ToList());
            FiltredJobList = CollectionViewSource.GetDefaultView(JobList);
            FiltredJobList.Filter = FilterBySearchText;
        }

        public async override Task InitializeAsync()
        {
            var mvm = ServiceLocator.ServiceProvider.GetRequiredService<MainViewModel>();
            mvm.WindowState = "Maximized";
            await base.InitializeAsync();
        }

        partial void OnFilterJobChanged(string value)
        {
            FiltredJobList.Refresh();
        }

        private bool FilterBySearchText(object item)
        {

            if (item is Job itemName)
            {
                if (string.IsNullOrWhiteSpace(FilterJob))
                    return true;

                return itemName.Name.ToLower().Contains(FilterJob.ToLower());
            }
            return false;
        }

        partial void OnSelectedJobChanged(Job value)
        {
            if (SelectedJob != null)
            {
                var tempBHATools = SelectedJob.Tools.Where(t => t.Status == "В КНБК").ToList();
                BHATools = new ObservableCollection<Tool>(tempBHATools);

                var tempBackupTools = SelectedJob.Tools.Where(t => t.Status == "В резерве").ToList();
                BackupTools = new ObservableCollection<Tool>(tempBackupTools);

                var tempForInspectionTools = SelectedJob.Tools.Where(t => t.Status == "На инспекцию").ToList();
                ForInspectionTools = new ObservableCollection<Tool>(tempForInspectionTools);

                var tempLostTools = SelectedJob.Tools.Where(t => t.Status == "Оставлено в скв.").ToList();
                LostTools = new ObservableCollection<Tool>(tempLostTools);

                var tempDepTools = SelectedJob.Tools.Where(t => t.Status == "Вывезено").ToList();
                DepTools = new ObservableCollection<Tool>(tempDepTools);

                DDList = new ObservableCollection<DD>(SelectedJob.DDs);

                MWDList = new ObservableCollection<MWD>(SelectedJob.MWDs);
            }

        }

        [RelayCommand]
        public async Task GoToDailyReportToDbPage() =>
            await navigationService.NavigateToAsync<DailyReportToDbPageModel>();

        [RelayCommand]
        public async Task GoToSearchByToolsPage() =>
            await navigationService.NavigateToAsync<SearchByToolsPageModel>();
    }
}
