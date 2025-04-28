using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Repos;
using TechReportToDB.ViewModels.Base;
using TechReportToDB.ViewModels.CustomControlModels;
using TechReportToDB.ViewModels.WindowModels;
using TechReportToDB.Views.Windows;

namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class DbManagerPageModel : PageModelBase
    {
        private readonly IRepo<Job> jobRepo;
        private readonly INavigationService navigationService;

        public ToolListModel BHAToolsModel { get; } = new ToolListModel { Title = "В КНБК" };
        public ToolListModel BackupToolsModel { get; } = new ToolListModel { Title = "В резерве" };
        public ToolListModel ForInspectionToolsModel { get; } = new ToolListModel { Title = "На инспекцию" };
        public ToolListModel LostToolsModel { get; } = new ToolListModel { Title = "Оставлено в скважине" };
        public ToolListModel DepToolsModel { get; } = new ToolListModel { Title = "На вывоз" };

        public ICollectionView FiltredJobList { get; }

        [ObservableProperty]
        private string filterJob = "";

        [ObservableProperty]
        private ObservableCollection<Job> jobList = new();

        [ObservableProperty]
        private ObservableCollection<DD> dDList = new();

        [ObservableProperty]
        private ObservableCollection<MWD> mWDList = new();

        [ObservableProperty]
        private Job selectedJob = new();

        public IRelayCommand<Person> PersonSelectedCommand { get; }

        public DbManagerPageModel(IRepo<Job> jobRepo, INavigationService navigationService)
        {
            this.jobRepo = jobRepo;
            this.navigationService = navigationService;

            JobList = new ObservableCollection<Job>(jobRepo.List.Include(t => t.Tools)
                .OrderBy(f => f.FieldTeam)
                .Include(d => d.DDs)
                .Include(m => m.MWDs)
                .Include(c=>c.Constructions.OrderBy(t=>t.TimeStamp)).ToList());
            
            FiltredJobList = CollectionViewSource.GetDefaultView(JobList);
            FiltredJobList.Filter = FilterBySearchText;

            PersonSelectedCommand = new RelayCommand<Person>(OnPersonSelected);
        }

        public async override Task InitializeAsync()
        {
            var mvm = ServiceLocator.ServiceProvider.GetRequiredService<MainViewModel>();
            //UpdateJobList();
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

                return itemName.FilterName.ToLower().Contains(FilterJob.ToLower());
            }
            return false;
        }

        partial void OnSelectedJobChanged(Job value)
        {
            if (SelectedJob != null)
            {
                var tempBHATools = SelectedJob.Tools.Where(t => t.Status == "В КНБК").ToList();
                BHAToolsModel.Tools = new ObservableCollection<Tool>(tempBHATools);

                var tempBackupTools = SelectedJob.Tools.Where(t => t.Status == "В резерве").ToList();
                BackupToolsModel.Tools = new ObservableCollection<Tool>(tempBackupTools);

                var tempForInspectionTools = SelectedJob.Tools.Where(t => t.Status == "На инспекцию").ToList();
                ForInspectionToolsModel.Tools = new ObservableCollection<Tool>(tempForInspectionTools);

                var tempLostTools = SelectedJob.Tools.Where(t => t.Status == "Оставлено в скв.").ToList();
                LostToolsModel.Tools = new ObservableCollection<Tool>(tempLostTools);

                var tempDepTools = SelectedJob.Tools.Where(t => t.Status == "Вывезено").ToList();
                DepToolsModel.Tools = new ObservableCollection<Tool>(tempDepTools);

                DDList = new ObservableCollection<DD>(SelectedJob.DDs);

                MWDList = new ObservableCollection<MWD>(SelectedJob.MWDs);
            }

        }

        [RelayCommand]
        public async Task GoToDailyReportToDbPage() =>
            await navigationService.NavigateToAsync<DailyReportToDbPageModel>();


        //[RelayCommand]
        //public async Task OpenPersonWindow() =>
        //    await navigationService.OpenWindowAsync<PersonWindow, PersonWindowModel>();

        //[RelayCommand]
        //public async Task OpenToolInfoWindow() =>
        //   await navigationService.OpenWindowAsync<ToolInfoWindow, ToolInfoWindowModel>();

        private void OnPersonSelected(Person? person)
        {
            var personWindowModel = ServiceLocator.ServiceProvider.GetRequiredService<PersonWindowModel>();

            if (personWindowModel!=null)
            {
                personWindowModel.Person = person;

                navigationService.OpenWindowAsync<PersonWindow, PersonWindowModel>();
            }
        }
    }
}
