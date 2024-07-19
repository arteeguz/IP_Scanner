using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using IPProcessingTool;

namespace IPProcessingTool
{
    public partial class MainWindow : Window
    {
        private string outputFilePath;
        public ObservableCollection<ScanStatus> ScanStatuses { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            outputFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "output.csv");
            EnsureCsvFile();

            ScanStatuses = new ObservableCollection<ScanStatus>();
            StatusDataGrid.ItemsSource = ScanStatuses;

            Logger.Log(LogLevel.INFO, "Application started");
        }

        private async void Button1_Click(object sender, RoutedEventArgs e)
        {
            var inputWindow = new InputWindow("Enter the IP address:", false);
            if (inputWindow.ShowDialog() == true)
            {
                string ip = inputWindow.InputText;
                if (IsValidIP(ip))
                {
                    Logger.Log(LogLevel.INFO, "User input IP address", context: "Button1_Click", additionalInfo: ip);
                    await ProcessIPAsync(ip);
                }
                else
                {
                    Logger.Log(LogLevel.WARNING, "Invalid IP address input", context: "Button1_Click", additionalInfo: ip);
                    ShowInvalidInputMessage();
                }
            }
        }

        private async void Button2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                string csvPath = openFileDialog.FileName;
                Logger.Log(LogLevel.INFO, "User selected CSV file", context: "Button2_Click", additionalInfo: csvPath);

                var ips = File.ReadAllLines(csvPath).Select(line => line.Trim()).ToList();
                foreach (var ip in ips)
                {
                    if (IsValidIP(ip))
                    {
                        await ProcessIPAsync(ip);
                    }
                    else
                    {
                        HighlightInvalidInput(ip);
                    }
                }
            }
        }

        private async void Button3_Click(object sender, RoutedEventArgs e)
        {
            var inputWindow = new InputWindow("Enter the IP segment (e.g., 192.168.1):", true);
            if (inputWindow.ShowDialog() == true)
            {
                string segment = inputWindow.InputText;
                if (IsValidIPSegment(segment))
                {
                    Logger.Log(LogLevel.INFO, "User input IP segment", context: "Button3_Click", additionalInfo: segment);

                    for (int i = 0; i < 256; i++)
                    {
                        string ip = $"{segment}.{i}";
                        await ProcessIPAsync(ip);
                    }
                }
                else
                {
                    Logger.Log(LogLevel.WARNING, "Invalid IP segment input", context: "Button3_Click", additionalInfo: segment);
                    ShowInvalidInputMessage();
                }
            }
        }

        private async void Button4_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                string csvPath = openFileDialog.FileName;
                Logger.Log(LogLevel.INFO, "User selected CSV file for segment scan", context: "Button4_Click", additionalInfo: csvPath);

                var segments = File.ReadAllLines(csvPath).Select(line => line.Trim()).ToList();
                foreach (var segment in segments)
                {
                    if (IsValidIPSegment(segment))
                    {
                        for (int i = 0; i < 256; i++)
                        {
                            string ip = $"{segment}.{i}";
                            await ProcessIPAsync(ip);
                        }
                    }
                    else
                    {
                        HighlightInvalidInput(segment);
                    }
                }
            }
        }

        private async Task ProcessIPAsync(string ip)
        {
            var scanStatus = new ScanStatus { IPAddress = ip, Status = "Processing", Details = "" };
            AddScanStatus(scanStatus);

            string date = DateTime.Now.ToString("M/dd/yyyy");
            string time = DateTime.Now.ToString("HH:mm");
            string status = "Not Reachable";
            string errorDetails = "N/A";
            string hostname = "N/A";
            string lastLoggedUser = "N/A";
            string machineType = "N/A";
            string windowsVersion = "N/A";

            Logger.Log(LogLevel.INFO, "Started processing IP", context: "ProcessIPAsync", additionalInfo: ip);

            await Task.Run(() =>
            {
                if (PingHost(ip))
                {
                    try
                    {
                        ConnectionOptions options = new ConnectionOptions
                        {
                            Impersonation = ImpersonationLevel.Impersonate,
                            EnablePrivileges = true,
                            Authentication = AuthenticationLevel.PacketPrivacy
                        };

                        ManagementScope scope = new ManagementScope($"\\\\{ip}\\root\\cimv2", options);
                        scope.Connect();

                        var machineQuery = new ObjectQuery("SELECT * FROM Win32_ComputerSystem");
                        var machineSearcher = new ManagementObjectSearcher(scope, machineQuery);
                        var machine = machineSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                        if (machine != null)
                        {
                            hostname = machine["Name"]?.ToString();
                            machineType = machine["Model"]?.ToString();

                            var userQuery = new ObjectQuery("SELECT * FROM Win32_NetworkLoginProfile");
                            var userSearcher = new ManagementObjectSearcher(scope, userQuery);
                            var user = userSearcher.Get().Cast<ManagementObject>().OrderByDescending(u => u["LastLogon"]).FirstOrDefault();
                            if (user != null)
                            {
                                lastLoggedUser = user["Name"]?.ToString();
                            }

                            windowsVersion = GetWindowsVersion(scope);
                            status = "Complete";
                        }
                        else
                        {
                            status = "WMI Error";
                            errorDetails = "Machine information not found.";
                        }
                    }
                    catch (ManagementException ex)
                    {
                        status = "WMI Error";
                        errorDetails = ex.Message;
                        Logger.Log(LogLevel.ERROR, $"WMI ManagementException for IP {ip}", context: "ProcessIPAsync", additionalInfo: ex.Message);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        status = "Access denied";
                        errorDetails = "Access denied when attempting to connect to " + ip;
                        Logger.Log(LogLevel.ERROR, $"UnauthorizedAccessException for IP {ip}", context: "ProcessIPAsync", additionalInfo: ex.Message);
                    }
                    catch (Exception ex)
                    {
                        status = "Unknown error";
                        errorDetails = ex.Message;
                        Logger.Log(LogLevel.ERROR, $"Exception for IP {ip}", context: "ProcessIPAsync", additionalInfo: ex.Message);
                    }
                }
                else
                {
                    errorDetails = "Host not reachable";
                    Logger.Log(LogLevel.WARNING, $"Host not reachable for IP {ip}", context: "ProcessIPAsync");
                }
            });

            scanStatus.Status = status;
            scanStatus.Details = errorDetails;
            SaveOutput(ip, hostname, lastLoggedUser, machineType, windowsVersion, date, time, status, errorDetails);
            UpdateScanStatus(scanStatus);

            Logger.Log(LogLevel.INFO, $"Processed IP {ip}", context: "ProcessIPAsync", additionalInfo: $"Status: {status}, Details: {errorDetails}");
        }

        private void AddScanStatus(ScanStatus scanStatus)
        {
            Dispatcher.Invoke(() =>
            {
                ScanStatuses.Add(scanStatus);
            });
        }

        private void UpdateScanStatus(ScanStatus scanStatus)
        {
            Dispatcher.Invoke(() =>
            {
                // This is to refresh the DataGrid display
                var index = ScanStatuses.IndexOf(scanStatus);
                ScanStatuses.RemoveAt(index);
                ScanStatuses.Insert(index, scanStatus);
            });
        }

        private bool PingHost(string ip)
        {
            try
            {
                Ping ping = new Ping();
                PingReply reply = ping.Send(ip, 1000);
                return reply.Status == IPStatus.Success;
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.ERROR, $"Ping exception for IP {ip}", context: "PingHost", additionalInfo: ex.Message);
                return false;
            }
        }

        private string GetWindowsVersion(ManagementScope scope)
        {
            try
            {
                var osQuery = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                var osSearcher = new ManagementObjectSearcher(scope, osQuery);
                var os = osSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                return os?["Caption"]?.ToString() ?? "Unknown";
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.ERROR, "Exception while getting Windows version", context: "GetWindowsVersion", additionalInfo: ex.Message);
                return "Unknown";
            }
        }

        private void SaveOutput(string ip, string hostname, string lastLoggedUser, string machineType, string windowsVersion, string date, string time, string status, string errorDetails)
        {
            var newLine = $"\"{ip}\",\"{hostname}\",\"{lastLoggedUser}\",\"{machineType}\",\"{windowsVersion}\",\"{date}\",\"{time}\",\"{status}\",\"{errorDetails}\"";
            using (var writer = new StreamWriter(outputFilePath, true, Encoding.UTF8))
            {
                writer.WriteLine(newLine);
            }
        }

        private void EnsureCsvFile()
        {
            if (!File.Exists(outputFilePath) || new FileInfo(outputFilePath).Length == 0)
            {
                var header = "\"IP\",\"Hostname\",\"LastLoggedUser\",\"MachineType\",\"WindowsVersion\",\"Date\",\"Time\",\"Status\",\"ErrorDetails\"";
                using (var writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                {
                    writer.WriteLine(header);
                }
            }
        }

        private bool IsValidIP(string ip)
        {
            return IPAddress.TryParse(ip, out _);
        }

        private bool IsValidIPSegment(string segment)
        {
            string[] parts = segment.Split('.');
            if (parts.Length != 3) return false;
            return parts.All(part => byte.TryParse(part, out _));
        }

        private void HighlightInvalidInput(string input)
        {
            var scanStatus = new ScanStatus { IPAddress = input, Status = "Invalid", Details = "Invalid IP/Segment" };
            AddScanStatus(scanStatus);
            Logger.Log(LogLevel.WARNING, "Invalid IP/Segment input", context: "HighlightInvalidInput", additionalInfo: input);
        }

        private void ShowInvalidInputMessage()
        {
            MessageBox.Show("Invalid IP or Segment format. Please enter a valid IP or Segment.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            Logger.Log(LogLevel.WARNING, "Displayed invalid input message", context: "ShowInvalidInputMessage");
        }
    }

    public class ScanStatus
    {
        public string IPAddress { get; set; }
        public string Status { get; set; }
        public string Details { get; set; }
    }
}
