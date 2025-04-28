using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.DependencyInjection;
using System.Collections.ObjectModel;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services;
using TechReportToDB.Services.Navigation;
using TechReportToDB.ViewModels.Base;
using TechReportToDB.ViewModels.WindowModels;
using TechReportToDB.Views.Windows;

namespace TechReportToDB.ViewModels.CustomControlModels
{
    internal partial class ToolListModel : PageModelBase
    {
        private readonly INavigationService navigationService;
        private readonly ToolInfoWindowModel toolInfoWindowModel =
            ServiceLocator.ServiceProvider.GetRequiredService<ToolInfoWindowModel>();

        [ObservableProperty]
        private ObservableCollection<Tool> tools = new();

        public IRelayCommand<Tool> ToolSelectedCommand { get; }

        public ToolListModel()
        {
            navigationService = ServiceLocator.ServiceProvider.GetRequiredService<INavigationService>();

            ToolSelectedCommand = new RelayCommand<Tool>(OnToolSelected);   
        }

        private void OnToolSelected(Tool? selectedItem)
        {
            toolInfoWindowModel.Tool = selectedItem ?? new Tool();
            navigationService.OpenWindowAsync<ToolInfoWindow, ToolInfoWindowModel>();
        }
    }
}
