using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using TechReportToDB.Data;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services.Navigation;
using TechReportToDB.ViewModels.Base;
using TechReportToDB.Views.Pages;

namespace TechReportToDB.ViewModels.PageModels
{
    internal partial class SearchByToolsPageModel : ViewModelBase
    {
        private readonly INavigationService navigationService;
        private readonly AppDbContext context;

        [ObservableProperty]
        private ObservableCollection<Tool> tools = new();

        [ObservableProperty]
        private string filterTool = "";

        public ICollectionView FiltredToolList { get; }

        public SearchByToolsPageModel(INavigationService navigationService, AppDbContext context)
        {
            this.navigationService = navigationService;
            this.context = context;
            Tools = new ObservableCollection<Tool>(context.Tools.Include(j=>j.Job).ToList());
            FiltredToolList = CollectionViewSource.GetDefaultView(Tools);
            FiltredToolList.Filter = FilterBySearchText;
        }

        public override async Task InitializeAsync()
        {
            
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
        public async Task GoToDBManagerPage() =>
            await navigationService.NavigateToAsync<DbManagerPageModel>();


    }
}
