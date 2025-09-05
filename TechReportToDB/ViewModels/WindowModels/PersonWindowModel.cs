using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Diagnostics;
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
    }
}
