using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechReportToDB.Services.Navigation;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.ViewModels.WindowModels
{
    internal partial class MapWindowModel : WindowModelBase
    {
        public MapWindowModel(INavigationService navigationService) : base(navigationService)
        {
        }
    }
}
