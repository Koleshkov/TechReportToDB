using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Net.Mail;
using System.Windows.Data;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services;
using TechReportToDB.Services.DailyReportToDb;
using TechReportToDB.Services.JobToJson;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Repos;
using TechReportToDB.ViewModels.Base;
using TechReportToDB.ViewModels.CustomControlModels;
using TechReportToDB.ViewModels.WindowModels;
using TechReportToDB.Views.Windows;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class DbManagerPageModel : PageModelBase
    {
        private readonly IRepo<Job> jobRepo;
        private readonly INavigationService navigationService;
        private readonly IJobToJsonService jobToJsonService;

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

        public DbManagerPageModel(IRepo<Job> jobRepo, INavigationService navigationService, IJobToJsonService jobToJsonService)
        {
            this.jobRepo = jobRepo;
            this.navigationService = navigationService;

            JobList = new ObservableCollection<Job>(jobRepo.List.Include(t => t.Tools)
                .OrderBy(f => f.FieldTeam)
                .Include(d => d.DDs)
                .Include(m => m.MWDs)
                .Include(c => c.Constructions.OrderBy(t => t.TimeStamp)).ToList());

            FiltredJobList = CollectionViewSource.GetDefaultView(JobList);
            FiltredJobList.Filter = FilterBySearchText;

            PersonSelectedCommand = new RelayCommand<Person>(OnPersonSelected);
            this.jobToJsonService = jobToJsonService;
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

                return itemName.FilterName!.ToLower().Contains(FilterJob.ToLower());
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


        [RelayCommand]
        public async Task ExportToJson()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Json (*.json)|*.json";

            saveFileDialog.Title = "Сохранить";

            saveFileDialog.FileName = $"JobList {DateTime.Now.ToShortDateString()}";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                await jobToJsonService.ExportToJson(filePath);
            }
            await Task.CompletedTask;
        }

        [RelayCommand]
        private void OpenPhone(string phone)
        {
            if (string.IsNullOrWhiteSpace(phone))
            {
                Console.WriteLine("Номер телефона не указан");
                return;
            }

            // Нормализация номера (удаляем всё, кроме цифр и +)
            var cleanNumber = new string(phone.Where(c => char.IsDigit(c) || c == '+').ToArray());

            // Если номер начинается с 8 (российский номер без кода страны)
            if (cleanNumber.StartsWith("8") && cleanNumber.Length > 1)
            {
                cleanNumber = "+7" + cleanNumber[1..];
            }
            // Если номер начинается с 7, но нет + (российский номер)
            else if (cleanNumber.StartsWith("7") && !cleanNumber.StartsWith("+7") && cleanNumber.Length > 1)
            {
                cleanNumber = "+" + cleanNumber;
            }

            try
            {
                // Правильный формат для tel: URI
                Process.Start(new ProcessStartInfo($"tel:{cleanNumber}") { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при открытии телефона: {ex.Message}");
            }
        }

        [RelayCommand]
        private void OpenEmail(string ft)
        {
            if (string.IsNullOrWhiteSpace(ft))
            {
                Console.WriteLine("Номер паритии не указан");
                return;
            }

            ft = $"ft{ft}@isnnb.rosneft.ru";

            try
            {
                // Правильный формат для tel: URI
                Process.Start(new ProcessStartInfo($"mailto:{ft}") { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при открытии gjxns: {ex.Message}");
            }
        }

        private void OnPersonSelected(Person? person)
        {
            var personWindowModel = ServiceLocator.ServiceProvider.GetRequiredService<PersonWindowModel>();

            if (personWindowModel != null)
            {
                personWindowModel.Person = person;

                navigationService.OpenWindowAsync<PersonWindow, PersonWindowModel>();
            }
        }
    }
}
