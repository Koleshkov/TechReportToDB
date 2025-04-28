using Microsoft.Extensions.DependencyInjection;
using System.Windows.Controls;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services;
using TechReportToDB.ViewModels.PageModels;
using UserControl = System.Windows.Controls.UserControl;

namespace TechReportToDB.Views.Pages
{
    /// <summary>
    /// Interaction logic for PadMapPage.xaml
    /// </summary>
    public partial class PadMapPage : UserControl
    {
        public PadMapPage()
        {
            InitializeComponent();

            var vm = ServiceLocator.ServiceProvider.GetRequiredService<PadMapPageModel>();
            vm.UpdateMapCallback = UpdateMapMarker;
            vm.SelectPadMapCallback = SelectMapMarker;


        }

        private void UpdateMapMarker(IEnumerable<Job> jobs)
        {
            YandexMapControl?.AddMarkers(jobs);
        }

        private void SelectMapMarker(Job jobs)
        {
            YandexMapControl?.MoveToMarker(jobs);
        }
    }
}
