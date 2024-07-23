using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace IPProcessingTool
{
    public partial class InputWindow : Window
    {
        public List<string> InputList { get; private set; }
        private bool isSegment;
        private bool isCSV;

        public InputWindow()
        {
            InitializeComponent();
            InputList = new List<string>();
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            InputList.Clear();
            IPListBox.Items.Clear();

            if (RadioButtonIP.IsChecked == true || RadioButtonSegment.IsChecked == true)
            {
                InputPanel.Visibility = Visibility.Visible;
                FilePanel.Visibility = Visibility.Collapsed;
                IPListBox.Visibility = Visibility.Visible;
            }
            else
            {
                InputPanel.Visibility = Visibility.Collapsed;
                FilePanel.Visibility = Visibility.Visible;
                IPListBox.Visibility = Visibility.Collapsed;
            }

            isSegment = RadioButtonSegment.IsChecked == true || RadioButtonSegmentCSV.IsChecked == true;
            isCSV = RadioButtonIPCSV.IsChecked == true || RadioButtonSegmentCSV.IsChecked == true;
        }

        private void InputTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            InputTextBox.TextChanged -= InputTextBox_TextChanged;

            string input = InputTextBox.Text;
            string formattedInput = FormatIP(input);

            InputTextBox.Text = formattedInput;
            InputTextBox.CaretIndex = formattedInput.Length;

            InputTextBox.TextChanged += InputTextBox_TextChanged;
        }

        private void InputTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !Regex.IsMatch(e.Text, "[0-9.]");
        }

        private string FormatIP(string input)
        {
            input = input.Replace("..", "."); // Replace double dots with a single dot
            if (isSegment)
            {
                input = string.Join(".", input.Split('.').Take(3).Select(part => part.Length > 3 ? part.Substring(0, 3) : part));
            }
            else
            {
                input = string.Join(".", input.Split('.').Take(4).Select(part => part.Length > 3 ? part.Substring(0, 3) : part));
            }

            return input;
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            string input = InputTextBox.Text;
            if (isSegment ? IsValidIPSegment(input) : IsValidIP(input))
            {
                InputList.Add(input);
                IPListBox.Items.Add(input);
                InputTextBox.Clear();
            }
            else
            {
                MessageBox.Show("Invalid input. Please enter a valid IP address or segment.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                FilePathTextBlock.Text = openFileDialog.FileName;
                InputList = File.ReadAllLines(openFileDialog.FileName).Select(line => line.Trim()).ToList();
            }
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            if (InputList.Count == 0 && string.IsNullOrEmpty(FilePathTextBlock.Text))
            {
                MessageBox.Show("No input provided. Please add IPs or browse a CSV file.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                DialogResult = true;
                Close();
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
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
    }
}
