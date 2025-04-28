using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.ComponentModel;
using Microsoft.EntityFrameworkCore;
using TechReportToDB.Data;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Repos;
using TechReportToDB.ViewModels.Base;
using TechReportToDB.Views.CustomControls;
using TechReportToDB.Views.Pages;
using System.Windows.Data;
using TechReportToDB.Services;
using Microsoft.Extensions.DependencyInjection;
using TechReportToDB.Services.Stuff;
using System.Diagnostics;
using TechReportToDB.Services.DailyReportToDb;

namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class StaffPageModel : PageModelBase
    {
        private readonly IRepo<MWD> mwdRepo;
        private readonly IRepo<DD> ddRepo;
        private readonly INavigationService navigationService;
        private readonly IStuffFService stuffFService;

        public ICollectionView FiltredPersonList { get; }

        [ObservableProperty]
        private ObservableCollection<Person> personList = new();

        [ObservableProperty]
        private string filterPerson = "";

        public StaffPageModel(INavigationService navigationService, IRepo<MWD> mwdRepo, 
            IRepo<DD> ddRepo, IStuffFService stuffFService)
        {
            this.navigationService = navigationService;
            this.stuffFService = stuffFService;
            this.mwdRepo = mwdRepo;
            this.ddRepo = ddRepo;

            UpdatePersonList();

            FiltredPersonList = CollectionViewSource.GetDefaultView(PersonList);
            FiltredPersonList.Filter = FilterBySearchText;
            
        }

        private async void UpdatePersonList()
        {
            ObservableCollection<Person> temp1 = new ObservableCollection<Person>(await mwdRepo.List.Include(p => p.Job).ToListAsync());
            ObservableCollection<Person> temp2 = new ObservableCollection<Person>(await ddRepo.List.Include(p => p.Job).ToListAsync());
            PersonList = new ObservableCollection<Person>(temp1.Concat(temp2));
        }

        public override async Task InitializeAsync()
        {
            var mvm = ServiceLocator.ServiceProvider.GetRequiredService<MainViewModel>();
            mvm.WindowState = "Maximized";
            await base.InitializeAsync();
        }

        partial void OnFilterPersonChanged(string value)
        {
            try
            {
                FiltredPersonList.Refresh();
            }
            catch (Exception)
            {

                throw;
            }
            
        }

        private bool FilterBySearchText(object item)
        {

            if (item is Person itemName)
            {
                if (string.IsNullOrWhiteSpace(FilterPerson))
                    return true;

                return itemName.FilterName.ToLower().Contains(FilterPerson.ToLower());
            }
            return false;
        }


        [RelayCommand]
        public async Task DailyReportToDbPage() =>
            await navigationService.NavigateToAsync<DailyReportToDbPageModel>();

        [RelayCommand]
        private void OpenWhatsApp(string phone)
        {
            if (!string.IsNullOrWhiteSpace(phone))
            {
                // Убираем все лишние символы (пробелы, дефисы и т.д.)
                phone = phone.Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace("+", "");

                // Если номер начинается с '8', заменяем на '7'
                if (phone.StartsWith("8"))
                {
                    phone = "7" + phone.Substring(1);
                }

                // Ссылка для открытия WhatsApp
                string whatsappUrl = $"https://wa.me/{phone}";

                try
                {
                    // Открываем ссылку в браузере
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = whatsappUrl,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    // Логируем ошибку, если не удалось открыть
                    Console.WriteLine($"Ошибка при открытии WhatsApp: {ex.Message}");
                }
            }
        }

        [RelayCommand]
        private void OpenTelegram(string phone)
        {
            if (!string.IsNullOrWhiteSpace(phone))
            {
                // Убираем лишние символы (пробелы, дефисы, скобки и т.д.)
                phone = phone.Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace("+", "");

                // Если номер начинается с '8', заменяем на '7'
                if (phone.StartsWith("8"))
                {
                    phone = "7" + phone.Substring(1);
                }

                // Формируем ссылку для открытия чата в Telegram
                string telegramUrl = $"https://t.me/+{phone}";

                try
                {
                    // Открываем ссылку в браузере
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = telegramUrl,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    // Логируем ошибку, если не удалось открыть
                    Console.WriteLine($"Ошибка при открытии Telegram: {ex.Message}");
                }
            }
        }

        [RelayCommand]
        public async Task Export()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

            saveFileDialog.Title = "Сохранить";

            saveFileDialog.FileName = $"Выгрузка по персоналу {DateTime.Now.ToShortDateString()}";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;

                await stuffFService.ExportToExcel(filePath, PersonList);

                Process process = new Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
            await Task.CompletedTask;
        }
    }
}
