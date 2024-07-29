﻿using System;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace IPProcessingTool
{
    public partial class InputWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private Brush _inputTextBoxBorderBrush = Brushes.Gray;
        public Brush InputTextBoxBorderBrush
        {
            get => _inputTextBoxBorderBrush;
            set
            {
                _inputTextBoxBorderBrush = value;
                OnPropertyChanged(nameof(InputTextBoxBorderBrush));
            }
        }

        private string _inputTextBoxToolTip;
        public string InputTextBoxToolTip
        {
            get => _inputTextBoxToolTip;
            set
            {
                _inputTextBoxToolTip = value;
                OnPropertyChanged(nameof(InputTextBoxToolTip));
            }
        }

        private string _errorMessage;
        public string ErrorMessage
        {
            get => _errorMessage;
            set
            {
                _errorMessage = value;
                OnPropertyChanged(nameof(ErrorMessage));
            }
        }

        public string InputText { get; private set; }
        private bool isSegment;

        public InputWindow(string labelText, bool isSegment = false)
        {
            InitializeComponent();
            DataContext = this;

            InputLabel.Content = labelText;
            InputTextBox.TextChanged += InputTextBox_TextChanged;
            InputTextBox.PreviewTextInput += InputTextBox_PreviewTextInput;
            this.isSegment = isSegment;
        }

        private void InputTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            int caretIndex = InputTextBox.CaretIndex;
            string input = InputTextBox.Text;

            // Only format if we're adding characters, not deleting
            if (e.Changes.Any(change => change.AddedLength > 0))
            {
                input = FormatInputWithDots(input);
            }

            InputTextBox.Text = input;
            SetCaretIndex(caretIndex, input);

            ValidateInput();
        }

        private void InputTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            string currentText = InputTextBox.Text;
            int dotCount = currentText.Count(c => c == '.');
            bool isNumberOrDot = Regex.IsMatch(e.Text, "[0-9.]");

            // Prevent further input if it's a dot and there are already 2 dots (for segments), or if it's not a valid character
            if ((e.Text == "." && dotCount >= (isSegment ? 2 : 3)) || !isNumberOrDot)
            {
                e.Handled = true;
            }
        }

        private string FormatInputWithDots(string input)
        {
            string[] parts = input.Split('.');
            string formattedInput = "";
            int maxParts = isSegment ? 3 : 4;

            for (int i = 0; i < parts.Length && i < maxParts; i++)
            {
                if (parts[i].Length > 3)
                {
                    parts[i] = parts[i].Substring(0, 3);
                }
                formattedInput += parts[i];

                if (i < maxParts - 1 && (parts[i].Length == 3 || i < parts.Length - 1))
                {
                    formattedInput += ".";
                }
            }

            return formattedInput;
        }

        private void SetCaretIndex(int oldIndex, string newText)
        {
            // Calculate how many characters were added
            int diff = newText.Length - InputTextBox.Text.Length;

            // Set new caret position
            InputTextBox.CaretIndex = Math.Min(oldIndex + diff, newText.Length);
        }

        private void ValidateInput()
        {
            string input = InputTextBox.Text.Trim();
            (bool isValid, string errorMessage) = isSegment ? ValidateIPSegment(input) : ValidateIP(input);

            if (isValid)
            {
                InputTextBoxBorderBrush = Brushes.Gray;
                InputTextBoxToolTip = null;
                ErrorMessage = null;
            }
            else
            {
                InputTextBoxBorderBrush = Brushes.Red;
                ErrorMessage = errorMessage;
                InputTextBoxToolTip = errorMessage;
            }
        }

        private (bool isValid, string errorMessage) ValidateIP(string ip)
        {
            if (string.IsNullOrWhiteSpace(ip))
                return (false, "Input cannot be empty.");

            if (!IPAddress.TryParse(ip, out _))
                return (false, "Invalid IP address format.");

            string[] parts = ip.Split('.');
            if (parts.Length != 4)
                return (false, "IP address must have four parts separated by dots.");

            foreach (var part in parts)
            {
                if (!byte.TryParse(part, out byte b))
                    return (false, $"'{part}' is not a valid number between 0 and 255.");
            }

            return (true, null);
        }

        private (bool isValid, string errorMessage) ValidateIPSegment(string segment)
        {
            if (string.IsNullOrWhiteSpace(segment))
                return (false, "Input cannot be empty.");

            string[] parts = segment.Split('.');
            if (parts.Length != 3)
                return (false, "IP segment must have exactly three parts separated by dots.");

            foreach (var part in parts)
            {
                if (!byte.TryParse(part, out byte b))
                    return (false, $"'{part}' is not a valid number between 0 and 255.");
            }

            return (true, null);
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            string input = InputTextBox.Text.Trim();
            (bool isValid, string errorMessage) = isSegment ? ValidateIPSegment(input) : ValidateIP(input);

            if (isValid)
            {
                IPListTextBox.AppendText(input + Environment.NewLine);
                InputTextBox.Clear();
                InputTextBoxBorderBrush = Brushes.Gray;
                InputTextBoxToolTip = null;
                ErrorMessage = null;
            }
            else
            {
                InputTextBoxBorderBrush = Brushes.Red;
                ErrorMessage = errorMessage;
                InputTextBoxToolTip = errorMessage;
                MessageBox.Show(errorMessage, "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            InputText = IPListTextBox.Text.Trim();
            if (!string.IsNullOrEmpty(InputText))
            {
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show("No IPs added. Please enter at least one IP address or segment.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}