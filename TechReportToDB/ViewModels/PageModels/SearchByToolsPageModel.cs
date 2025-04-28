using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Data;
using TechReportToDB.Data;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services;
using TechReportToDB.Services.DailyReportToDb;
using TechReportToDB.Services.Equip;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Repos;
using TechReportToDB.ViewModels.Base;
using TechReportToDB.Views.Pages;

namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class SearchByToolsPageModel : PageModelBase
    {
        private readonly IRepo<Tool> toolRepo;
        private readonly INavigationService navigationService;
        private readonly IDailyReportToDbService dailyReportToDbService;
        private readonly IEquipToExcel equipToExcel;

        public ICollectionView FiltredToolList { get; }

        [ObservableProperty]
        private ObservableCollection<Tool> toolList = new();

        [ObservableProperty]
        private string filterTool = "";



        public SearchByToolsPageModel(IRepo<Tool> toolRepo, INavigationService navigationService,
            IDailyReportToDbService dailyReportToDbService, IEquipToExcel equipToExcel)
        {
            this.toolRepo = toolRepo;
            this.navigationService = navigationService;
            this.dailyReportToDbService = dailyReportToDbService;
            this.equipToExcel = equipToExcel;

            UpdateToolList();

            FiltredToolList = CollectionViewSource.GetDefaultView(ToolList);
            FiltredToolList.Filter = FilterBySearchText;
            
        }

        private async void UpdateToolList()
        {
            ToolList = new ObservableCollection<Tool>(await toolRepo.List.Include(j => j.Job).ToListAsync());
        }

        public override async Task InitializeAsync()
        {
            var mvm = ServiceLocator.ServiceProvider.GetRequiredService<MainViewModel>();
            mvm.WindowState = "Maximized";

            await base.InitializeAsync();
        }

        partial void OnFilterToolChanged(string value)
        {
            FiltredToolList.Refresh();
        }

        private bool FilterBySearchText(object item)
        {

            if (item is Tool itemName)
            {
                if (string.IsNullOrWhiteSpace(FilterTool))
                    return true;

                return itemName.FilterName.ToLower().Contains(FilterTool.ToLower());
            }
            return false;
        }

        [RelayCommand]
        public async Task DailyReportToDbPage() =>
            await navigationService.NavigateToAsync<DailyReportToDbPageModel>();

        [RelayCommand]
        public async Task Export()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Text files (*.xlsm)|*.xlsm";

            saveFileDialog.Title = "Сохранить";

            saveFileDialog.FileName = $"Сводная координатора проекта {DateTime.Now.ToShortDateString()}";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                await dailyReportToDbService.ExportToExcel(filePath);
                Process process = new Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
            await Task.CompletedTask;

        }

        [RelayCommand]
        public async Task ExportEquip()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Text files (*.xlsm)|*.xlsx";

            saveFileDialog.Title = "Сохранить";

            saveFileDialog.FileName = $"Оборудование на КП {DateTime.Now.ToShortDateString()}";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;

                await equipToExcel.ExportToExcel(filePath);

                Process process = new Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
        }
    }
}
