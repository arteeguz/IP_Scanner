using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace IPProcessingTool
{
    public partial class OptionsWindow : Window
    {
        private MainWindow _mainWindow;

        public OptionsWindow(MainWindow mainWindow)
        {
            InitializeComponent();
            _mainWindow = mainWindow;
        }

        private void ReviewButton_Click(object sender, RoutedEventArgs e)
        {
            string scanType = ((ComboBoxItem)ScanTypeComboBox.SelectedItem).Content.ToString();
            string input = InputTextBox.Text;

            var reviewWindow = new ReviewWindow(scanType, input, GetSelectedDataFields(), _mainWindow);
            reviewWindow.Show();
            this.Close();
        }

        private List<string> GetSelectedDataFields()
        {
            List<string> selectedFields = new List<string>();
            if (HostnameCheckBox.IsChecked == true) selectedFields.Add("Hostname");
            if (LastLoggedUserCheckBox.IsChecked == true) selectedFields.Add("LastLoggedUser");
            if (MachineTypeCheckBox.IsChecked == true) selectedFields.Add("MachineType");
            if (MachineSKUCheckBox.IsChecked == true) selectedFields.Add("MachineSKU");
            if (InstalledCoreSoftwareCheckBox.IsChecked == true) selectedFields.Add("InstalledCoreSoftware");
            if (RAMSizeCheckBox.IsChecked == true) selectedFields.Add("RAMSize");
            if (WindowsVersionCheckBox.IsChecked == true) selectedFields.Add("WindowsVersion");
            if (WindowsBuildCheckBox.IsChecked == true) selectedFields.Add("WindowsBuild");
            return selectedFields;
        }
    }
}
