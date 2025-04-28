using GMap.NET;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using TechReportToDB.Data.Entities;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;



namespace TechReportToDB.Views.CustomControls
{
    public partial class YandexMap : UserControl
    {
        public YandexMap()
        {
            InitializeComponent();
            InitializeMap();
        }

        private async void InitializeMap()
        {
            var options = new CoreWebView2EnvironmentOptions
            {
                AdditionalBrowserArguments = "--disable-features=PrivacySandbox,SiteIsolation,BlockThirdPartyCookies"
            };
            var environment = await CoreWebView2Environment.CreateAsync(null, null, options);
            await PART_MapWebView.EnsureCoreWebView2Async(environment);

            // Путь к HTML файлу
            string htmlFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates", "Map.html");

            if (File.Exists(htmlFilePath))
            {
                string fileUri = new Uri(htmlFilePath).AbsoluteUri;
                PART_MapWebView.Source = new Uri(fileUri);
            }
            else
            {
                MessageBox.Show("HTML файл не найден.");
            }
        }

        // Метод для добавления списка маркеров
        internal void AddMarkers(IEnumerable<Job> jobs)
        {
            var markerList = jobs.Select(j => new
            {
                latitude = j.Latitude,
                longitude = j.Longitude,
                hint = j.Well,
                balloon = j.Label
            });

            string jsonMarkers = System.Text.Json.JsonSerializer.Serialize(markerList);

            string script = $@"
                addMarkers({jsonMarkers});
            ";

            if (PART_MapWebView.CoreWebView2 == null)
            {
                MessageBox.Show("WebView2 еще не инициализирован.");
               
            }
            else
            {
                PART_MapWebView.CoreWebView2.ExecuteScriptAsync(script);
            }

            // Выполняем скрипт в WebView2
            
        }

        // Метод для перемещения к выбранному маркеру
        internal void MoveToMarker(Job job)
        {
            string script = $@"
                moveToMarker({job.Latitude}, {job.Longitude});
            ";

            // Выполняем скрипт в WebView2
            PART_MapWebView.CoreWebView2.ExecuteScriptAsync(script);
        }
    }
}
