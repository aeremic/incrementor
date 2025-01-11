using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace Incrementor.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string inputFilePath = string.Empty;
        private string outputFilePath = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void InputLocationButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new()
            {
                DefaultExt = ".xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
            };

            var result = dlg.ShowDialog();

            if (result != true)
            {
                return;
            }

            inputFilePath = dlg.FileName;
            InputLocationTextBox.Text = inputFilePath;
        }

        private void SaveLocationButton_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new OpenFolderDialog();

            if (folderDialog.ShowDialog() != true)
            {
                return;
            }

            outputFilePath = folderDialog.FolderName;
            SaveLocationTextBox.Text = outputFilePath;
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(inputFilePath)
                || string.IsNullOrEmpty(outputFilePath))
            {
                return;
            }
            
            var incrementorParsingResult = Logic.Incrementor
                .ProcessData(inputFilePath, Path.Combine(outputFilePath,"Output.xlsx"));
            if (incrementorParsingResult.ParsingResult)
            {
                MessageBox.Show(
                    $"New file saved at {incrementorParsingResult.OutputFilePath}",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show(
                    $"New file not saved." +
                    $" Code: {(int)incrementorParsingResult.ErrorType}. " +
                    $"{incrementorParsingResult.ErrorMessage}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }
}