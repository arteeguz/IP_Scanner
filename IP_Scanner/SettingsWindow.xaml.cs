using System;
using System.Windows;

namespace IPProcessingTool
{
    public partial class SettingsWindow : Window
    {
        private Config config;

        public SettingsWindow(Config currentConfig)
        {
            InitializeComponent();
            config = currentConfig;
            LoadSettings();
        }

        private void LoadSettings()
        {
            OutputFilePathTextBox.Text = config.OutputFilePath;
            MaxConcurrentScansTextBox.Text = config.MaxConcurrentScans.ToString();
            ScanTimeoutTextBox.Text = config.ScanTimeout.ToString();
            EnableDetailedLoggingCheckBox.IsChecked = config.EnableDetailedLogging;

            // Load data collection options
            LastLoggedUserCheckBox.IsChecked = config.DataCollectionSettings.ShouldCollect(DataCollectionOptions.LastLoggedUser);
            MachineTypeCheckBox.IsChecked = config.DataCollectionSettings.ShouldCollect(DataCollectionOptions.MachineType);
            MachineSKUCheckBox.IsChecked = config.DataCollectionSettings.ShouldCollect(DataCollectionOptions.MachineSKU);
            InstalledCoreSoftwareCheckBox.IsChecked = config.DataCollectionSettings.ShouldCollect(DataCollectionOptions.InstalledCoreSoftware);
            RAMSizeCheckBox.IsChecked = config.DataCollectionSettings.ShouldCollect(DataCollectionOptions.RAMSize);
            WindowsVersionCheckBox.IsChecked = config.DataCollectionSettings.ShouldCollect(DataCollectionOptions.WindowsVersion);
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateSettings())
            {
                config.OutputFilePath = OutputFilePathTextBox.Text;
                config.MaxConcurrentScans = int.Parse(MaxConcurrentScansTextBox.Text);
                config.ScanTimeout = int.Parse(ScanTimeoutTextBox.Text);
                config.EnableDetailedLogging = EnableDetailedLoggingCheckBox.IsChecked ?? false;

                // Save data collection options
                config.DataCollectionSettings.Options = DataCollectionOptions.Hostname | DataCollectionOptions.Timestamp | DataCollectionOptions.Status;
                if (LastLoggedUserCheckBox.IsChecked == true) config.DataCollectionSettings.Options |= DataCollectionOptions.LastLoggedUser;
                if (MachineTypeCheckBox.IsChecked == true) config.DataCollectionSettings.Options |= DataCollectionOptions.MachineType;
                if (MachineSKUCheckBox.IsChecked == true) config.DataCollectionSettings.Options |= DataCollectionOptions.MachineSKU;
                if (InstalledCoreSoftwareCheckBox.IsChecked == true) config.DataCollectionSettings.Options |= DataCollectionOptions.InstalledCoreSoftware;
                if (RAMSizeCheckBox.IsChecked == true) config.DataCollectionSettings.Options |= DataCollectionOptions.RAMSize;
                if (WindowsVersionCheckBox.IsChecked == true) config.DataCollectionSettings.Options |= DataCollectionOptions.WindowsVersion;

                ConfigManager.SaveConfig(config);
                DialogResult = true;
                Close();
            }
        }

        private bool ValidateSettings()
        {
            if (string.IsNullOrWhiteSpace(OutputFilePathTextBox.Text))
            {
                MessageBox.Show("Output File Path cannot be empty.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (!int.TryParse(MaxConcurrentScansTextBox.Text, out int maxConcurrentScans) || maxConcurrentScans < 1)
            {
                MessageBox.Show("Max Concurrent Scans must be a positive integer.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (!int.TryParse(ScanTimeoutTextBox.Text, out int scanTimeout) || scanTimeout < 1)
            {
                MessageBox.Show("Scan Timeout must be a positive integer.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }
    }
}