using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.IdentityModel.Tokens;
using System.Diagnostics;
using System.IO.Packaging;
using System.Windows.Forms;
using TechReportToDB.Data;
using TechReportToDB.Services;
using TechReportToDB.Services.ExcelToDb;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Outlook;
using TechReportToDB.ViewModels.Base;
using TechReportToDB.ViewModels.WindowModels;
using TechReportToDB.Views.Pages;
using TechReportToDB.Views.Windows;


namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class DailyReportToDbPageModel : ViewModelBase
    {
        private readonly IExcelToDbService excelToDbService;
        private readonly IOutlookService outlookService;
        private readonly INavigationService navigationService;
        private readonly AppDbContext context;
        private readonly ErrorWindowModel errorWindowModel;

        [ObservableProperty]
        private DateTime dateOfReport = DateTime.Now.AddDays(-1);

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

        public DailyReportToDbPageModel(IExcelToDbService excelToDbService,
            INavigationService navigationService,
            ErrorWindowModel errorWindowModel,
            IOutlookService outlookService,
            AppDbContext context)
        {
            this.excelToDbService = excelToDbService;
            this.navigationService = navigationService;
            this.errorWindowModel = errorWindowModel;
            this.outlookService = outlookService;
            this.context = context;
        }

        public async override Task InitializeAsync()
        {
            var mvm = ServiceLocator.ServiceProvider.GetRequiredService<MainViewModel>();
            mvm.WindowState = "Normal";

            if (await context.Jobs.AnyAsync())
                {
                    var dateTime = await context.Jobs.FirstOrDefaultAsync();

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

                string savedPath = await Task.Run(() => outlookService.DownloadDailyReports(progress, folderBrowserDialog.SelectedPath, DateOfReport));

                count = 1;

                progress = new Progress<int>(p =>
                {
                    ProgressBarValue = 50 + p / 2;
                    count++;
                    Message = $"Сгружено файлов в БД: {count}";
                });

                Message = await Task.Run(() => excelToDbService.SaveToolsToDbAsync(progress, savedPath));

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


                Message = await Task.Run(() => excelToDbService.SaveToolsToDbAsync(progress, folderBrowserDialog.SelectedPath));
            }

            IsEnabledBtns = true;
            IsEnabledExlBtn = true;
            ProgressBarValue = 0;
        }

        [RelayCommand]
        public async Task DownloadFromOutlook()
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

                Message = $"Скачано в: {await Task.Run(() => outlookService.DownloadDailyReports(progress, folderBrowserDialog.SelectedPath, DateOfReport))}";
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

                await excelToDbService.ExportToExcel(filePath);

                Process process = new Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }

            Message = "Сводная таблица готова";

            await Task.CompletedTask;
        }

        [RelayCommand]
        public async Task GoToDbManagerPageModel() =>
           await navigationService.NavigateToAsync<DbManagerPageModel>();
    }
}
