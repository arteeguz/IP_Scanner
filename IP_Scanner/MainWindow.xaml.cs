﻿using System;
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

        private async void OpenInputWindowButton_Click(object sender, RoutedEventArgs e)
        {
            InputWindow inputWindow = new InputWindow();
            if (inputWindow.ShowDialog() == true)
            {
                var inputList = inputWindow.InputList;
                foreach (var input in inputList)
                {
                    if (IsValidIP(input))
                    {
                        Logger.Log(LogLevel.INFO, "User input IP address", context: "OpenInputWindowButton_Click", additionalInfo: input);
                        await ProcessIPAsync(input);
                    }
                    else if (IsValidIPSegment(input))
                    {
                        Logger.Log(LogLevel.INFO, "User input IP segment", context: "OpenInputWindowButton_Click", additionalInfo: input);
                        for (int i = 0; i < 256; i++)
                        {
                            string ip = $"{input}.{i}";
                            await ProcessIPAsync(ip);
                        }
                    }
                    else
                    {
                        Logger.Log(LogLevel.WARNING, "Invalid input", context: "OpenInputWindowButton_Click", additionalInfo: input);
                        HighlightInvalidInput(input);
                    }
                }
            }
        }

        private async Task ProcessCSV(string path, bool isSegment)
        {
            var lines = File.ReadAllLines(path).Select(line => line.Trim()).ToList();
            foreach (var line in lines)
            {
                if (isSegment)
                {
                    if (IsValidIPSegment(line))
                    {
                        for (int i = 0; i < 256; i++)
                        {
                            string ip = $"{line}.{i}";
                            await ProcessIPAsync(ip);
                        }
                    }
                    else
                    {
                        HighlightInvalidInput(line);
                    }
                }
                else
                {
                    if (IsValidIP(line))
                    {
                        await ProcessIPAsync(line);
                    }
                    else
                    {
                        HighlightInvalidInput(line);
                    }
                }
            }
        }

        private bool IsCSVFile(string path)
        {
            return Path.GetExtension(path).Equals(".csv", StringComparison.OrdinalIgnoreCase);
        }

        private async Task ProcessIPAsync(string ip)
        {
            var scanStatus = new ScanStatus { IPAddress = ip, Status = "Processing", Details = "", MachineSKU = "", InstalledCoreSoftware = "", RAMSize = "" };
            AddScanStatus(scanStatus);

            UpdateStatusBar("Processing IP: " + ip);

            string date = DateTime.Now.ToString("M/dd/yyyy");
            string time = DateTime.Now.ToString("HH:mm");
            string status = "Not Reachable";
            string errorDetails = "N/A";
            string hostname = "N/A";
            string lastLoggedUser = "N/A";
            string machineType = "N/A";
            string machineSKU = "N/A";
            string installedCoreSoftware = "N/A";
            string ramSize = "N/A";
            string windowsVersion = "N/A";
            string windowsRelease = "N/A";

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

                            var skuQuery = new ObjectQuery("SELECT * FROM Win32_ComputerSystemProduct");
                            var skuSearcher = new ManagementObjectSearcher(scope, skuQuery);
                            var sku = skuSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                            if (sku != null)
                            {
                                machineSKU = sku["Version"]?.ToString();
                            }

                            var userQuery = new ObjectQuery("SELECT * FROM Win32_NetworkLoginProfile");
                            var userSearcher = new ManagementObjectSearcher(scope, userQuery);
                            var user = userSearcher.Get().Cast<ManagementObject>().OrderByDescending(u => u["LastLogon"]).FirstOrDefault();
                            if (user != null)
                            {
                                lastLoggedUser = user["Name"]?.ToString();
                            }

                            // Fetch installed core software
                            var softwareQuery = new ObjectQuery("SELECT * FROM Win32_Product WHERE Name LIKE 'Core%'");
                            var softwareSearcher = new ManagementObjectSearcher(scope, softwareQuery);
                            var softwareList = softwareSearcher.Get().Cast<ManagementObject>().Select(soft => soft["Name"].ToString());
                            installedCoreSoftware = string.Join(", ", softwareList);

                            // Fetch RAM size
                            var ramQuery = new ObjectQuery("SELECT * FROM Win32_PhysicalMemory");
                            var ramSearcher = new ManagementObjectSearcher(scope, ramQuery);
                            var totalRam = ramSearcher.Get().Cast<ManagementObject>().Sum(ram => Convert.ToDouble(ram["Capacity"]));
                            ramSize = $"{totalRam / (1024 * 1024 * 1024)} GB";

                            // Fetch Windows version and release
                            var osQuery = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                            var osSearcher = new ManagementObjectSearcher(scope, osQuery);
                            var os = osSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                            if (os != null)
                            {
                                windowsVersion = os["Caption"]?.ToString();
                                windowsRelease = MapWindowsRelease(os["BuildNumber"]?.ToString());
                            }

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
            scanStatus.MachineSKU = machineSKU;
            scanStatus.InstalledCoreSoftware = installedCoreSoftware;
            scanStatus.RAMSize = ramSize;
            SaveOutput(ip, hostname, lastLoggedUser, machineType, machineSKU, installedCoreSoftware, ramSize, windowsVersion, windowsRelease, date, time, status, errorDetails);
            UpdateScanStatus(scanStatus);

            Logger.Log(LogLevel.INFO, $"Processed IP {ip}", context: "ProcessIPAsync", additionalInfo: $"Status: {status}, Details: {errorDetails}");
            UpdateStatusBar("Completed processing IP: " + ip);
        }

        private string MapWindowsRelease(string buildNumber)
        {
            if (string.IsNullOrEmpty(buildNumber)) return "Unknown";

            // Map build numbers to Windows version names
            switch (buildNumber)
            {
                case "19041":
                case "19042":
                case "19043":
                case "19044":
                    return "Windows 10 20H2";
                case "19045":
                    return "Windows 10 21H2";
                case "19046":
                    return "Windows 10 22H2";
                case "22000":
                    return "Windows 11 21H2";
                case "22621":
                case "22622":
                    return "Windows 11 22H2";
                case "22631":
                case "22632":
                    return "Windows 11 23H2";
                default:
                    return "Unknown";
            }
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

        private void SaveOutput(string ip, string hostname, string lastLoggedUser, string machineType, string machineSKU, string installedCoreSoftware, string ramSize, string windowsVersion, string windowsRelease, string date, string time, string status, string errorDetails)
        {
            EnsureCsvFile();

            var newLine = $"\"{ip}\",\"{hostname}\",\"{lastLoggedUser}\",\"{machineType}\",\"{machineSKU}\",\"{installedCoreSoftware}\",\"{ramSize}\",\"{windowsVersion} ({windowsRelease})\",\"{date}\",\"{time}\",\"{status}\",\"{errorDetails}\"";
            try
            {
                using (var writer = new StreamWriter(outputFilePath, true, Encoding.UTF8))
                {
                    writer.WriteLine(newLine);
                }
            }
            catch (IOException ioEx)
            {
                Logger.Log(LogLevel.ERROR, $"IOException while writing to file {outputFilePath}: {ioEx.Message}", context: "SaveOutput");
                MessageBox.Show($"IOException while writing to file {outputFilePath}: {ioEx.Message}", "File Access Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.ERROR, $"Exception while writing to file {outputFilePath}: {ex.Message}", context: "SaveOutput");
                MessageBox.Show($"Exception while writing to file {outputFilePath}: {ex.Message}", "File Access Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void EnsureCsvFile()
        {
            try
            {
                var expectedHeader = "\"IP\",\"Hostname\",\"LastLoggedUser\",\"MachineType\",\"MachineSKU\",\"InstalledCoreSoftware\",\"RAMSize\",\"WindowsVersion\",\"WindowsBuild\",\"Date\",\"Time\",\"Status\",\"ErrorDetails\"";

                if (!File.Exists(outputFilePath))
                {
                    using (var writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                    {
                        writer.WriteLine(expectedHeader);
                    }
                }
                else
                {
                    var lines = File.ReadAllLines(outputFilePath).ToList();
                    if (lines.Count == 0 || !lines[0].Contains("IP"))
                    {
                        lines.Insert(0, expectedHeader);
                        File.WriteAllLines(outputFilePath, lines);
                    }
                }
            }
            catch (IOException ioEx)
            {
                Logger.Log(LogLevel.ERROR, $"IOException while ensuring file {outputFilePath}: {ioEx.Message}", context: "EnsureCsvFile");
                MessageBox.Show($"IOException while ensuring file {outputFilePath}: {ioEx.Message}", "File Access Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.ERROR, $"Exception while ensuring file {outputFilePath}: {ex.Message}", context: "EnsureCsvFile");
                MessageBox.Show($"Exception while ensuring file {outputFilePath}: {ex.Message}", "File Access Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool IsFileLocked(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    return false;
                }

                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                return true;
            }
            return false;
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

        private void UpdateStatusBar(string message)
        {
            Dispatcher.Invoke(() =>
            {
                StatusBarText.Text = message;
            });
        }
    }

    public class ScanStatus
    {
        public string IPAddress { get; set; }
        public string Status { get; set; }
        public string Details { get; set; }
        public string MachineSKU { get; set; }
        public string InstalledCoreSoftware { get; set; }
        public string RAMSize { get; set; }
    }
}
