using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.DependencyInjection;
using System.Diagnostics;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services;
using TechReportToDB.Services.DailyReportToDb;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Outlook;
using TechReportToDB.Services.Repos;
using TechReportToDB.Services.ShortReport;
using TechReportToDB.Services.WorkstationReportToDb;
using TechReportToDB.ViewModels.Base;
using TechReportToDB.ViewModels.WindowModels;
using TechReportToDB.Views.Windows;


namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class DailyReportToDbPageModel : PageModelBase
    {
        private readonly INavigationService navigationService;
        private readonly IDailyReportToDbService dailyReportToDbService;
        private readonly IWorkstationReportToDbService workstationReportToDbService;
        private readonly IShortReportService shortReportService;
        private readonly IOutlookService outlookService;
        private readonly IRepo<Job> jobRepo;

        [ObservableProperty]
        private DateTime dateOfReport = DateTime.Now;

        [ObservableProperty]
        private string header = "Выберите дату и действие";

        [ObservableProperty]
        private string message = "";

        [ObservableProperty]
        private int progressBarValue = 0;

        [ObservableProperty]
        private bool isEnabledBtns = true;

        [ObservableProperty]
        private bool isEnabledExlBtn = false;

        public DailyReportToDbPageModel(IDailyReportToDbService dailyReportToDbService,
            IOutlookService outlookService,
            IRepo<Job> jobRepo,
            INavigationService navigationService,
            IWorkstationReportToDbService workstationReportToDbService,
            IShortReportService shortReportService)
        {
            this.dailyReportToDbService = dailyReportToDbService;
            this.outlookService = outlookService;
            this.jobRepo = jobRepo;
            this.navigationService = navigationService;
            this.workstationReportToDbService = workstationReportToDbService;
            this.shortReportService = shortReportService;
        }

        public async override Task InitializeAsync()
        {
            var mvm = ServiceLocator.ServiceProvider.GetRequiredService<MainViewModel>();
            mvm.WindowState = "Normal";

            if (await jobRepo.AnyAsync())
            {
                var dateTime = await jobRepo.FirstOrDefaultAsync();

                Message = $"Поледние обновление базы данных {dateTime?.TimeStamp.ToLongDateString()} в {dateTime?.TimeStamp.ToShortTimeString()}";
                IsEnabledExlBtn = true;
            }
            else
            {
                Message = "База данных пуста. Загрузите суточные рапорта.";
            }
        }

        [RelayCommand]
        public async Task DownloadExportCreate()
        {
            var vm = ServiceLocator.ServiceProvider.GetRequiredService<DateTimePageModel>();

            if (vm != null)
            {
                vm.SelectAction = DownloadExportCreateAction;
                await navigationService.NavigateToAsync<DateTimePageModel>();
            }
        }

        private async Task DownloadExportCreateAction(DateTime dateTime, int selectedStartTime, int selectedEndTime)
        {
            IsEnabledBtns = false;

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                int count = 0;

                var progress = new Progress<int>(s =>
                {
                    ProgressBarValue = s / 2;
                    count++;
                    Message = $"Скачано файлов: {count}";
                });

                string savedPath = await Task.Run(() => outlookService.DownloadDailyReports(progress, folderBrowserDialog.SelectedPath, DateOfReport, selectedStartTime, selectedEndTime));

                count = 1;

                progress = new Progress<int>(p =>
                {
                    ProgressBarValue = 50 + p / 2;
                    count++;
                    Message = $"Сгружено файлов в БД: {count}";
                });

                Message = await Task.Run(() => dailyReportToDbService.SaveToolsToDbAsync(progress, savedPath));

                await CreateExcel();
            }

            IsEnabledBtns = true;
            IsEnabledExlBtn = true;
            ProgressBarValue = 0;
        }

        [RelayCommand]
        public async Task ExportToDb()
        {
            IsEnabledBtns = false;
            IsEnabledExlBtn = false;

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                int count = 0;
                Progress<int> progress = new Progress<int>(p =>
                {
                    ProgressBarValue = p;
                    count++;
                    Message = $"Сгружено файлов в БД: {count}";
                });


                Message = await Task.Run(() => dailyReportToDbService.SaveToolsToDbAsync(progress, folderBrowserDialog.SelectedPath));
            }

            IsEnabledBtns = true;
            IsEnabledExlBtn = true;
            ProgressBarValue = 0;
        }

        [RelayCommand]
        public async Task DownloadFromOutlook()
        {
            var vm = ServiceLocator.ServiceProvider.GetRequiredService<DateTimePageModel>();

            if (vm != null)
            {
                vm.SelectAction = DownloadFromOutlookAction;
                await navigationService.NavigateToAsync<DateTimePageModel>();
            }
        }

        private async Task DownloadFromOutlookAction(DateTime dateTime, int selectedStartTime, int selectedEndTime)
        {
            IsEnabledBtns = false;

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {

                int count = 0;

                var progress = new Progress<int>(s =>
                {
                    ProgressBarValue = s;
                    count++;
                    Message = $"Скачано файлов: {count}";
                });

                Message = $"Скачано в: {await Task.Run(() => outlookService.DownloadDailyReports(progress, folderBrowserDialog.SelectedPath, DateOfReport, selectedStartTime, selectedEndTime))}";
            }

            IsEnabledBtns = true;
            ProgressBarValue = 0;
        }

        [RelayCommand]
        public async Task CreateExcel()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Text files (*.xlsm)|*.xlsm";

            saveFileDialog.Title = "Сохранить";

            saveFileDialog.FileName = $"Сводная координатора проекта {DateTime.Now.ToShortDateString()}";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Message = "Формируем сводную таблицу";

                string filePath = saveFileDialog.FileName;

                await dailyReportToDbService.ExportToExcel(filePath);

                Process process = new Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }

            Message = "Сводная таблица готова";

            await Task.CompletedTask;
        }

        [RelayCommand]
        public async Task CreateShortReportTable()
        {
            IsEnabledBtns = false;

            var vm = ServiceLocator.ServiceProvider.GetRequiredService<DateTimeShortPageModel>();

            if (vm != null)
            {
                vm.SelectAction = CreateShortReportTableAction;
                await navigationService.NavigateToAsync<DateTimeShortPageModel>();
            }

            IsEnabledBtns = true;
        }

        private async Task CreateShortReportTableAction(DateTime dateTime, int selectedTime)
        {

            OpenFileDialog openFileDialog = new()
            {
                Filter = "Excel file (*.xlsx)|*.xlsx",
                Title = "Открыть",
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                int count = 0;

                var progress = new Progress<int>(p =>
                {
                    ProgressBarValue = 50 + p / 2;
                    count++;
                    Message = $"Сгружено файлов в БД: {count}";
                });

                await Task.Run(() => shortReportService.ExportToExcel(progress, openFileDialog.FileName, dateTime, selectedTime));

                Process process = new Process();
                process.StartInfo.FileName = openFileDialog.FileName.Replace("xlsx","_New.xlsx");
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
            ProgressBarValue = 0;
        }

        [RelayCommand]
        public async Task CreateWorkstationExcel()
        {
            IsEnabledBtns = false;

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                int count = 0;
                Progress<int> progress = new Progress<int>(p =>
                {
                    ProgressBarValue = p;
                    count++;
                    Message = $"Сгружено файлов в БД: {count}";
                });


                Message = await Task.Run(() => workstationReportToDbService.SaveToDbAsync(progress, folderBrowserDialog.SelectedPath));
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Text files (*.xlsx)|*.xlsx";

            saveFileDialog.Title = "Сохранить";

            saveFileDialog.FileName = $"Сводный по вагонам  {DateTime.Now.ToShortDateString()}";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Message = "Формируем сводную таблицу";

                string filePath = saveFileDialog.FileName;

                await workstationReportToDbService.ExportToExcel(filePath);

                Process process = new Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }

            Message = "Сводная таблица готова";

            await Task.CompletedTask;

            IsEnabledBtns = true;

            ProgressBarValue = 0;

            await Task.CompletedTask;
        }

        [RelayCommand]
        public async Task GoToDbManagerPageModel() =>
           await navigationService.NavigateToAsync<DbManagerPageModel>();

        [RelayCommand]
        public async Task GoToPadMapPageModel() =>
            await navigationService.NavigateToAsync<PadMapPageModel>();

        [RelayCommand]
        public async Task GoToSearchByToolsPageModel() =>
            await navigationService.NavigateToAsync<SearchByToolsPageModel>();

        [RelayCommand]
        public async Task GoToSaffPageModel() =>
            await navigationService.NavigateToAsync<StaffPageModel>();
    }
}
