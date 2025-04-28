using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TechReportToDB.Data.Entities;
using TechReportToDB.Services.Navigation;
using TechReportToDB.Services.Repos;
using TechReportToDB.ViewModels.Base;

namespace TechReportToDB.ViewModels.WindowModels
{
    internal partial class PersonWindowModel : WindowModelBase
    {
        [ObservableProperty]
        private Person? person;
        public PersonWindowModel(INavigationService navigationService) : base(navigationService)
        {
        }


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
    }
}
