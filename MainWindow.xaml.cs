using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using System.Net.Http;
using System.IO;

namespace OfficeDeploymentTool
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            btnConfirm.Click += BtnConfirm_Click;
        }

        private async void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var officeVersion = GetSelectedOfficeVersion();
                if (string.IsNullOrEmpty(officeVersion))
                {
                    MessageBox.Show("Vui lòng chọn phiên bản Office hợp lệ.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (officeVersion == "2016")
                {
                    MessageBox.Show("Phiên bản Office 2016 hiện tại không còn được hỗ trợ.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                AppendLog("Bắt đầu tạo file cấu hình...");

                var languageId = GetSelectedLanguage();
                var selectedApps = GetSelectedApps();

                string selectedArch = ((ComboBoxItem?)comboArch.SelectedItem)?.Tag?.ToString() ?? "64";

                AppendLog($"Phiên bản: Office {officeVersion}");
                AppendLog($"Ngôn ngữ: {languageId}, Kiến trúc: {selectedArch}-bit");
                AppendLog($"Ứng dụng đã chọn: {string.Join(", ", selectedApps)}");

                var xml = await Task.Run(() =>
                    GenerateConfigurationXml(officeVersion, languageId, selectedApps, selectedArch));

                AppendLog("Đang tải Office Deployment Tool (setup.exe)...");
                await DownloadOdtAndSaveXmlAsync(xml);

                AppendLog("Đã lưu Configuration.xml và setup.exe vào C:\\Office");

                AppendLog("Đang tiến hành cài đặt Office...");
                await Task.Run(() =>
                {
                    var process = new System.Diagnostics.Process
                    {
                        StartInfo = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = "setup.exe",
                            Arguments = "/configure Configuration.xml",
                            WorkingDirectory = @"C:\Office",
                            UseShellExecute = true,
                            CreateNoWindow = true
                        }
                    };
                    process.Start();
                    process.WaitForExit();
                });

                AppendLog("Cài đặt hoàn tất.");
                MessageBox.Show("Đã cài đặt xong!", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppendLog("Lỗi xảy ra: " + ex.Message);
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task DownloadOdtAndSaveXmlAsync(XDocument xml)
        {
            string folderPath = @"C:\Office";
            string odtUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe";
            string odtPath = Path.Combine(folderPath, "setup.exe");
            string xmlPath = Path.Combine(folderPath, "Configuration.xml");

            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            using var client = new HttpClient();
            var data = await client.GetByteArrayAsync(odtUrl);
            await File.WriteAllBytesAsync(odtPath, data);

            xml.Save(xmlPath);
        }

        private XDocument GenerateConfigurationXml(string officeVersion, string languageId, string[] selectedApps, string arch)
        {
            var (channel, productIdMain, pidKeyMain, productIdVisio, pidKeyVisio, productIdProject, pidKeyProject) = officeVersion switch
            {
                "2019" => ("PerpetualVL2019", "ProPlus2019Volume", "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP", "VisioPro2019Volume", "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB", "ProjectPro2019Volume", "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B"),
                "2021" => ("PerpetualVL2021", "ProPlus2021Volume", "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH", "VisioPro2021Volume", "KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4", "ProjectPro2021Volume", "FTNWT-C6WBT-8HMGF-K9PRX-QV9H8"),
                "365" => ("Current", "O365ProPlusRetail", null, "VisioPro2024Volume", "B7TN8-FJ8V3-7QYCP-HQPMV-YY89G", "ProjectPro2024Volume", "FQQ23-N4YCY-73HQ3-FM9WC-76HF4"),
                _ => throw new Exception("Chưa chọn phiên bản Office hợp lệ!")
            };

            string[] allApps = { "Access", "OneNote", "Powerpoint", "Teams", "Excel", "Outlook", "Publisher", "Word" };
            var excludeApps = allApps.Except(selectedApps, System.StringComparer.OrdinalIgnoreCase)
                                     .Concat(new[] { "OneDrive", "Groove", "Lync" }).Distinct();

            var configuration = new XElement("Configuration", new XAttribute("ID", System.Guid.NewGuid().ToString()));
            var add = new XElement("Add",
                new XAttribute("OfficeClientEdition", arch),
                new XAttribute("Channel", channel)
            );

            void AddProduct(string productId, string? pidKey)
            {
                var product = new XElement("Product", new XAttribute("ID", productId));
                if (!string.IsNullOrEmpty(pidKey))
                    product.Add(new XAttribute("PIDKEY", pidKey));
                product.Add(new XElement("Language", new XAttribute("ID", ConvertLangId(languageId))));
                foreach (var app in excludeApps)
                    product.Add(new XElement("ExcludeApp", new XAttribute("ID", app)));
                add.Add(product);
            }

            AddProduct(productIdMain, pidKeyMain);

            if (selectedApps.Contains("Visio"))
                AddProduct(productIdVisio, pidKeyVisio);

            if (selectedApps.Contains("Project"))
                AddProduct(productIdProject, pidKeyProject);

            configuration.Add(add);

            configuration.Add(
                new XElement("Property", new XAttribute("Name", "SharedComputerLicensing"), new XAttribute("Value", "0")),
                new XElement("Property", new XAttribute("Name", "FORCEAPPSHUTDOWN"), new XAttribute("Value", "FALSE")),
                new XElement("Property", new XAttribute("Name", "DeviceBasedLicensing"), new XAttribute("Value", "0")),
                new XElement("Property", new XAttribute("Name", "SCLCacheOverride"), new XAttribute("Value", "0")),
                new XElement("Property", new XAttribute("Name", "AUTOACTIVATE"), new XAttribute("Value", "1")),
                new XElement("Updates", new XAttribute("Enabled", "TRUE")),
                new XElement("RemoveMSI"),
                new XElement("Display", new XAttribute("Level", "Full"), new XAttribute("AcceptEULA", "TRUE"))
            );

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), configuration);
        }

        private string? GetSelectedOfficeVersion()
        {
            if (rbOffice2016?.IsChecked == true) return "2016";
            if (rbOffice2019?.IsChecked == true) return "2019";
            if (rbOffice2021?.IsChecked == true) return "2021";
            if (rbOffice365?.IsChecked == true) return "365";
            return null;
        }

        private string GetSelectedLanguage()
        {
            return ((ComboBoxItem?)comboLanguage.SelectedItem)?.Tag?.ToString() ?? "en_US";
        }

        private string ConvertLangId(string lang) => lang.Replace('_', '-').ToLower();

        private string[] GetSelectedApps()
        {
            var apps = new[] {
                (cbAccess, "Access"), (cbOneNote, "OneNote"), (cbPowerpoint, "Powerpoint"),
                (cbTeams, "Teams"), (cbExcel, "Excel"), (cbOutlook, "Outlook"),
                (cbPublisher, "Publisher"), (cbWord, "Word"),
                (cbProject, "Project"), (cbVisio, "Visio")
            };

            return apps.Where(a => a.Item1?.IsChecked == true).Select(a => a.Item2).ToArray();
        }

        private void AppendLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                LogTextBox.ScrollToEnd();
            });
        }
    }
}
