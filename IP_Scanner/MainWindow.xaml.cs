using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace IPProcessingTool
{
    public partial class MainWindow : Window
    {
        private string outputFilePath;
        public ObservableCollection<ScanStatus> ScanStatuses { get; set; }
        private CancellationTokenSource cancellationTokenSource;
        private ParallelOptions parallelOptions;
        private ObservableCollection<ColumnSetting> dataColumnSettings;
        private bool autoSave;
        private int pingTimeout = 1000; // Default value
        private int totalIPs;
        private int processedIPs;

        public MainWindow()
        {
            InitializeComponent();
            ScanStatuses = new ObservableCollection<ScanStatus>();
            StatusDataGrid.ItemsSource = ScanStatuses;

            parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };

            dataColumnSettings = new ObservableCollection<ColumnSetting>();
            autoSave = false; // Default value

            InitializeColumnSettings();
            UpdateDataGridColumns();

            Logger.Log(LogLevel.INFO, "Application started");
        }

        private void InitializeColumnSettings()
        {
            dataColumnSettings = new ObservableCollection<ColumnSetting>
            {
                new ColumnSetting { Name = "IP Address", IsSelected = true },
                new ColumnSetting { Name = "Hostname", IsSelected = true },
                new ColumnSetting { Name = "Last Logged User", IsSelected = true },
                new ColumnSetting { Name = "Machine Type", IsSelected = true },
                new ColumnSetting { Name = "Machine SKU", IsSelected = true },
                new ColumnSetting { Name = "Installed Core Software", IsSelected = true },
                new ColumnSetting { Name = "RAM Size", IsSelected = true },
                new ColumnSetting { Name = "Windows Version", IsSelected = true },
                new ColumnSetting { Name = "Windows Release", IsSelected = true },
                new ColumnSetting { Name = "Date", IsSelected = true },
                new ColumnSetting { Name = "Time", IsSelected = true },
                new ColumnSetting { Name = "Status", IsSelected = true },
                new ColumnSetting { Name = "Details", IsSelected = true }
            };
        }

        private void UpdateDataGridColumns()
        {
            StatusDataGrid.Columns.Clear();
            foreach (var column in dataColumnSettings.Where(c => c.IsSelected))
            {
                StatusDataGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = column.Name,
                    Binding = new System.Windows.Data.Binding(column.Name.Replace(" ", ""))
                });
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new Settings(dataColumnSettings, autoSave, pingTimeout, parallelOptions.MaxDegreeOfParallelism);
            if (settingsWindow.ShowDialog() == true)
            {
                dataColumnSettings = new ObservableCollection<ColumnSetting>(settingsWindow.DataColumns);
                autoSave = settingsWindow.AutoSave;
                pingTimeout = settingsWindow.PingTimeout;
                parallelOptions.MaxDegreeOfParallelism = settingsWindow.MaxConcurrentScans;

                UpdateDataGridColumns();
            }
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
                    await ProcessIPsAsync(new[] { ip });
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
                await ProcessIPsAsync(ips);
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
                    var ips = Enumerable.Range(0, 256).Select(i => $"{segment}.{i}");
                    await ProcessIPsAsync(ips);
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
                var ips = segments.SelectMany(segment => Enumerable.Range(0, 256).Select(i => $"{segment}.{i}"));
                await ProcessIPsAsync(ips);
            }
        }

        private async Task ProcessIPsAsync(IEnumerable<string> ips)
        {
            totalIPs = ips.Count();
            processedIPs = 0;
            UpdateProgressBar(0);

            DisableButtons();

            cancellationTokenSource = new CancellationTokenSource();
            var tasks = new List<Task>();

            try
            {
                foreach (var ip in ips)
                {
                    if (cancellationTokenSource.IsCancellationRequested)
                        break;

                    if (IsValidIP(ip))
                    {
                        tasks.Add(ProcessIPAsync(ip, cancellationTokenSource.Token));
                    }
                    else
                    {
                        HighlightInvalidInput(ip);
                    }

                    if (tasks.Count >= parallelOptions.MaxDegreeOfParallelism)
                    {
                        await Task.WhenAny(tasks);
                        tasks.RemoveAll(t => t.IsCompleted);
                    }
                }

                await Task.WhenAll(tasks);
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.ERROR, "Error processing IPs", context: "ProcessIPsAsync", additionalInfo: ex.Message);
                MessageBox.Show($"An error occurred while processing IPs: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                EnableButtons();
                UpdateStatusBar("Completed processing all IPs.");
                UpdateProgressBar(100);
                HandleAutoSave();
            }
        }

        private async Task ProcessIPAsync(string ip, CancellationToken cancellationToken = default)
        {
            var scanStatus = new ScanStatus
            {
                IPAddress = ip,
                Status = "Processing",
                Details = "",
                Date = DateTime.Now.ToString("M/dd/yyyy"),
                Time = DateTime.Now.ToString("HH:mm")
            };
            AddScanStatus(scanStatus);

            UpdateStatusBar($"Processing IP: {ip} ({processedIPs + 1}/{totalIPs})");

            Logger.Log(LogLevel.INFO, "Started processing IP", context: "ProcessIPAsync", additionalInfo: ip);

            try
            {
                if (await PingHostAsync(ip, cancellationToken))
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

                        // Check which columns are selected in the settings
                        var selectedColumns = dataColumnSettings.Where(c => c.IsSelected).Select(c => c.Name).ToList();

                        if (selectedColumns.Contains("Hostname") || selectedColumns.Contains("Machine Type"))
                        {
                            var machineQuery = new ObjectQuery("SELECT Name, Model FROM Win32_ComputerSystem");
                            var machineSearcher = new ManagementObjectSearcher(scope, machineQuery);
                            var machine = machineSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                            if (machine != null)
                            {
                                if (selectedColumns.Contains("Hostname"))
                                    scanStatus.Hostname = machine["Name"]?.ToString();
                                if (selectedColumns.Contains("Machine Type"))
                                    scanStatus.MachineType = machine["Model"]?.ToString();
                            }
                        }

                        if (selectedColumns.Contains("Machine SKU"))
                        {
                            var skuQuery = new ObjectQuery("SELECT Version FROM Win32_ComputerSystemProduct");
                            var skuSearcher = new ManagementObjectSearcher(scope, skuQuery);
                            var sku = skuSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                            if (sku != null)
                            {
                                scanStatus.MachineSKU = sku["Version"]?.ToString();
                            }
                        }

                        if (selectedColumns.Contains("Last Logged User"))
                        {
                            var userQuery = new ObjectQuery("SELECT UserName FROM Win32_ComputerSystem");
                            var userSearcher = new ManagementObjectSearcher(scope, userQuery);
                            var user = userSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                            if (user != null)
                            {
                                scanStatus.LastLoggedUser = user["UserName"]?.ToString();
                            }
                        }

                        if (selectedColumns.Contains("Installed Core Software"))
                        {
                            var softwareQuery = new ObjectQuery("SELECT Name, Version FROM Win32_Product");
                            var softwareSearcher = new ManagementObjectSearcher(scope, softwareQuery);
                            var softwareList = softwareSearcher.Get().Cast<ManagementObject>()
                                .Select(soft => $"{soft["Name"]} ({soft["Version"]})")
                                .Take(10)
                                .ToList();
                            scanStatus.InstalledCoreSoftware = string.Join(", ", softwareList);
                        }

                        if (selectedColumns.Contains("RAM Size"))
                        {
                            var ramQuery = new ObjectQuery("SELECT Capacity FROM Win32_PhysicalMemory");
                            var ramSearcher = new ManagementObjectSearcher(scope, ramQuery);
                            var totalRam = ramSearcher.Get().Cast<ManagementObject>().Sum(ram => Convert.ToDouble(ram["Capacity"]));
                            scanStatus.RAMSize = $"{totalRam / (1024 * 1024 * 1024):F2} GB";
                        }

                        if (selectedColumns.Contains("Windows Version") || selectedColumns.Contains("Windows Release"))
                        {
                            var osQuery = new ObjectQuery("SELECT Caption, BuildNumber FROM Win32_OperatingSystem");
                            var osSearcher = new ManagementObjectSearcher(scope, osQuery);
                            var os = osSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                            if (os != null)
                            {
                                if (selectedColumns.Contains("Windows Version"))
                                    scanStatus.WindowsVersion = os["Caption"]?.ToString();
                                if (selectedColumns.Contains("Windows Release"))
                                {
                                    string buildNumber = os["BuildNumber"]?.ToString();
                                    scanStatus.WindowsRelease = MapWindowsRelease(buildNumber);
                                }
                            }
                        }

                        scanStatus.Status = "Complete";
                        scanStatus.Details = "N/A";
                    }
                    catch (ManagementException ex)
                    {
                        scanStatus.Status = "Error";
                        scanStatus.Details = $"WMI Error: {ex.Message}";
                        Logger.Log(LogLevel.ERROR, $"WMI ManagementException for IP {ip}", context: "ProcessIPAsync", additionalInfo: ex.Message);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        scanStatus.Status = "Error";
                        scanStatus.Details = $"Access denied: {ex.Message}";
                        Logger.Log(LogLevel.ERROR, $"UnauthorizedAccessException for IP {ip}", context: "ProcessIPAsync", additionalInfo: ex.Message);
                    }
                    catch (Exception ex)
                    {
                        scanStatus.Status = "Error";
                        scanStatus.Details = $"Unknown error: {ex.Message}";
                        Logger.Log(LogLevel.ERROR, $"Exception for IP {ip}", context: "ProcessIPAsync", additionalInfo: ex.Message);
                    }
                }
                else
                {
                    scanStatus.Status = "Not Reachable";
                    scanStatus.Details = "Host not reachable";
                    Logger.Log(LogLevel.WARNING, $"Host not reachable for IP {ip}", context: "ProcessIPAsync");
                }

                if (cancellationToken.IsCancellationRequested)
                {
                    Logger.Log(LogLevel.INFO, "Cancellation requested", context: "ProcessIPAsync");
                    cancellationToken.ThrowIfCancellationRequested();
                }
            }
            catch (OperationCanceledException)
            {
                scanStatus.Status = "Cancelled";
                scanStatus.Details = "Operation canceled by user";
                Logger.Log(LogLevel.INFO, "Operation was canceled", context: "ProcessIPAsync");
            }
            finally
            {
                processedIPs++;
                UpdateProgressBar((int)((double)processedIPs / totalIPs * 100));
                UpdateScanStatus(scanStatus);
                UpdateStatusBar($"Completed processing IP: {ip} ({processedIPs}/{totalIPs})");
            }
        }

        private async Task<bool> PingHostAsync(string ip, CancellationToken cancellationToken)
        {
            try
            {
                using (var ping = new Ping())
                {
                    var pingTask = ping.SendPingAsync(ip, pingTimeout, new byte[32], new PingOptions(64, true));
                    var timeoutTask = Task.Delay(pingTimeout, cancellationToken);

                    var completedTask = await Task.WhenAny(pingTask, timeoutTask);

                    if (completedTask == pingTask)
                    {
                        var reply = await pingTask;
                        return reply.Status == IPStatus.Success;
                    }
                    else
                    {
                        // Timeout occurred
                        return false;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                Logger.Log(LogLevel.INFO, $"Ping operation cancelled for IP {ip}", context: "PingHostAsync");
                throw; // Re-throw the cancellation exception to be handled by the caller
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.ERROR, $"Ping exception for IP {ip}", context: "PingHostAsync", additionalInfo: ex.Message);
                return false;
            }
        }

        private string MapWindowsRelease(string buildNumber)
        {
            if (string.IsNullOrEmpty(buildNumber)) return "Unknown";

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
                    return $"Unknown (Build {buildNumber})";
            }
        }

        private void AddScanStatus(ScanStatus scanStatus)
        {
            Dispatcher.Invoke(() =>
            {
                lock (ScanStatuses)
                {
                    ScanStatuses.Add(scanStatus);
                }
            });
        }

        private void UpdateScanStatus(ScanStatus scanStatus)
        {
            Dispatcher.Invoke(() =>
            {
                lock (ScanStatuses)
                {
                    var index = ScanStatuses.IndexOf(scanStatus);
                    if (index != -1)
                    {
                        ScanStatuses[index] = scanStatus;
                    }
                }
            });
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

        private void UpdateProgressBar(int value)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = value;
            });
        }

        private void DisableButtons()
        {
            Dispatcher.Invoke(() =>
            {
                Button1.IsEnabled = false;
                Button2.IsEnabled = false;
                Button3.IsEnabled = false;
                Button4.IsEnabled = false;
            });
        }

        private void EnableButtons()
        {
            Dispatcher.Invoke(() =>
            {
                Button1.IsEnabled = true;
                Button2.IsEnabled = true;
                Button3.IsEnabled = true;
                Button4.IsEnabled = true;
            });
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveOutputFile();
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            if (cancellationTokenSource != null)
            {
                cancellationTokenSource.Cancel();
                UpdateStatusBar("Scanning stopped by user.");
                EnableButtons();
            }
        }

        private void HandleAutoSave()
        {
            if (autoSave)
            {
                SaveOutputFile();
            }
            else
            {
                ShowSavePrompt();
            }
        }

        private void ShowSavePrompt()
        {
            var result = MessageBox.Show("IP scanning is finished. Would you like to save the output?", "Save Results", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                SaveOutputFile();
            }
        }

        private void SaveOutputFile()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv",
                Title = "Save Output File"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                outputFilePath = saveFileDialog.FileName;
                bool fileExists = File.Exists(outputFilePath);

                if (fileExists)
                {
                    var result = MessageBox.Show("File already exists. Do you want to append to it?", "File Exists", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    else if (result == MessageBoxResult.No)
                    {
                        // Overwrite the file
                        File.WriteAllText(outputFilePath, string.Empty);
                        EnsureCsvFile();
                    }
                    // If Yes, we'll append to the existing file
                }
                else
                {
                    EnsureCsvFile();
                }

                // Now save all the scan results
                SaveAllScanResults();
            }
        }

        private void SaveAllScanResults()
        {
            try
            {
                using (var writer = new StreamWriter(outputFilePath, true, Encoding.UTF8))
                {
                    // Write header
                    writer.WriteLine(string.Join(",", dataColumnSettings.Where(c => c.IsSelected).Select(c => $"\"{c.Name}\"")));

                    foreach (var scanStatus in ScanStatuses)
                    {
                        var line = string.Join(",", dataColumnSettings.Where(c => c.IsSelected).Select(c =>
                        {
                            var value = GetPropertyValue(scanStatus, c.Name.Replace(" ", ""));
                            return $"\"{value}\"";
                        }));
                        writer.WriteLine(line);
                    }
                }
                MessageBox.Show("Output saved successfully.", "Save Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.ERROR, $"Error saving output: {ex.Message}", context: "SaveAllScanResults");
                MessageBox.Show($"Error saving output: {ex.Message}", "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GetPropertyValue(ScanStatus scanStatus, string propertyName)
        {
            var property = typeof(ScanStatus).GetProperty(propertyName);
            return property?.GetValue(scanStatus)?.ToString() ?? "N/A";
        }

        private void EnsureCsvFile()
        {
            try
            {
                if (!File.Exists(outputFilePath))
                {
                    var header = string.Join(",", dataColumnSettings.Where(c => c.IsSelected).Select(c => $"\"{c.Name}\""));
                    File.WriteAllText(outputFilePath, header + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.ERROR, $"Error ensuring CSV file: {ex.Message}", context: "EnsureCsvFile");
                MessageBox.Show($"Error ensuring CSV file: {ex.Message}", "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    public class ScanStatus
    {
        public string IPAddress { get; set; }
        public string Hostname { get; set; }
        public string LastLoggedUser { get; set; }
        public string MachineType { get; set; }
        public string MachineSKU { get; set; }
        public string InstalledCoreSoftware { get; set; }
        public string RAMSize { get; set; }
        public string WindowsVersion { get; set; }
        public string WindowsRelease { get; set; }
        public string Date { get; set; }
        public string Time { get; set; }
        public string Status { get; set; }
        public string Details { get; set; }

        public ScanStatus()
        {
            IPAddress = "";
            Hostname = "N/A";
            LastLoggedUser = "N/A";
            MachineType = "N/A";
            MachineSKU = "N/A";
            InstalledCoreSoftware = "N/A";
            RAMSize = "N/A";
            WindowsVersion = "N/A";
            WindowsRelease = "N/A";
            Date = DateTime.Now.ToString("M/dd/yyyy");
            Time = DateTime.Now.ToString("HH:mm");
            Status = "Not Started";
            Details = "N/A";
        }
    }
}