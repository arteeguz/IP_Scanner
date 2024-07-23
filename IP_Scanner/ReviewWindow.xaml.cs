using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace IPProcessingTool
{
    public partial class ReviewWindow : Window
    {
        private string _scanType;
        private string _input;
        private List<string> _selectedFields;
        private MainWindow _mainWindow;

        public ReviewWindow(string scanType, string input, List<string> selectedFields, MainWindow mainWindow)
        {
            InitializeComponent();
            _scanType = scanType;
            _input = input;
            _selectedFields = selectedFields;
            _mainWindow = mainWindow;

            ReviewLabel.Content = $"Scan Type: {_scanType}\nInput: {_input}\nFields: {string.Join(", ", _selectedFields)}";
        }

        private async void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            await Task.Run(() =>
            {
                if (_scanType == "Single IP")
                {
                    string ip = _input;
                    ProcessIP(ip);
                }
                else if (_scanType == "IP Segment")
                {
                    string segment = _input;
                    for (int i = 0; i < 256; i++)
                    {
                        string ip = $"{segment}.{i}";
                        ProcessIP(ip);
                    }
                }
                else if (_scanType == "Load from CSV")
                {
                    // Logic to read IPs from CSV and process them
                }
            });

            _mainWindow.Show();
            this.Close();
        }

        private void ProcessIP(string ip)
        {
            var scanStatus = new ScanStatus { IPAddress = ip, Status = "Processing", Details = "", MachineSKU = "", InstalledCoreSoftware = "", RAMSize = "" };
            _mainWindow.AddScanStatus(scanStatus);

            // Simulate processing
            System.Threading.Thread.Sleep(1000);

            // Update status
            scanStatus.Status = "Complete";
            scanStatus.Details = "N/A";
            scanStatus.MachineSKU = "SampleSKU";
            scanStatus.InstalledCoreSoftware = "SampleSoftware";
            scanStatus.RAMSize = "16GB";
            _mainWindow.UpdateScanStatus(scanStatus);
        }
    }
}
