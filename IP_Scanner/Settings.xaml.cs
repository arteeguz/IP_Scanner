using System.Collections.ObjectModel;
using System.Windows;

namespace IPProcessingTool
{
    public partial class Settings : Window
    {
        public ObservableCollection<ColumnSetting> Columns { get; set; }

        public Settings(ObservableCollection<ColumnSetting> currentColumns)
        {
            InitializeComponent();
            Columns = new ObservableCollection<ColumnSetting>(currentColumns);
            ColumnsList.ItemsSource = Columns;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    public class ColumnSetting
    {
        public string Name { get; set; }
        public bool IsSelected { get; set; }
    }
}