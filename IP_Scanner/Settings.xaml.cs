﻿using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Linq;

namespace IPProcessingTool
{
    public partial class Settings : Window
    {
        public ObservableCollection<ColumnSetting> GridColumns { get; set; }
        public ObservableCollection<ColumnSetting> OutputColumns { get; set; }
        public bool AutoSave { get; set; }
        public int PingTimeout { get; set; }
        public int MaxConcurrentScans { get; set; }

        public Settings(ObservableCollection<ColumnSetting> currentGridColumns, ObservableCollection<ColumnSetting> currentOutputColumns, bool autoSave, int pingTimeout, int maxConcurrentScans)
        {
            InitializeComponent();

            GridColumns = new ObservableCollection<ColumnSetting>(currentGridColumns.Select(c => new ColumnSetting { Name = c.Name, IsSelected = c.IsSelected }));
            OutputColumns = new ObservableCollection<ColumnSetting>(currentOutputColumns.Select(c => new ColumnSetting { Name = c.Name, IsSelected = c.IsSelected }));

            GridColumnsList.ItemsSource = GridColumns;
            OutputColumnsList.ItemsSource = OutputColumns;

            AutoSave = autoSave;
            PingTimeout = pingTimeout;
            MaxConcurrentScans = maxConcurrentScans;

            AutoSaveCheckBox.IsChecked = AutoSave;
            PingTimeoutTextBox.Text = PingTimeout.ToString();
            MaxConcurrentScansTextBox.Text = MaxConcurrentScans.ToString();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateSettings())
            {
                AutoSave = AutoSaveCheckBox.IsChecked ?? false;
                PingTimeout = int.Parse(PingTimeoutTextBox.Text);
                MaxConcurrentScans = int.Parse(MaxConcurrentScansTextBox.Text);

                DialogResult = true;
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private bool ValidateSettings()
        {
            if (!int.TryParse(PingTimeoutTextBox.Text, out int pingTimeout) || pingTimeout <= 0)
            {
                MessageBox.Show("Please enter a valid positive integer for Ping Timeout.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (!int.TryParse(MaxConcurrentScansTextBox.Text, out int maxConcurrentScans) || maxConcurrentScans <= 0)
            {
                MessageBox.Show("Please enter a valid positive integer for Max Concurrent Scans.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }
    }

    public class ColumnSetting
    {
        public string Name { get; set; }
        public bool IsSelected { get; set; }
    }
}