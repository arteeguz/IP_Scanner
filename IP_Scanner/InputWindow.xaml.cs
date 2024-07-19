using System;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace IPProcessingTool
{
    public partial class InputWindow : Window
    {
        public string InputText { get; private set; }
        private bool isSegment;

        public InputWindow(string labelText, bool isSegment = false)
        {
            InitializeComponent();
            InputLabel.Content = labelText;
            InputTextBox.TextChanged += InputTextBox_TextChanged;
            InputTextBox.PreviewTextInput += InputTextBox_PreviewTextInput;
            this.isSegment = isSegment;
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

        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            if (isSegment ? IsValidIPSegment(InputTextBox.Text) : IsValidIP(InputTextBox.Text))
            {
                InputText = InputTextBox.Text;
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show("Invalid input. Please enter a valid IP address or segment.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Back_Click(object sender, RoutedEventArgs e)
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
