using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace Incrementor.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private string _inputFilePath = string.Empty;
        private string _outputFilePath = string.Empty;

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

            _inputFilePath = dlg.FileName;
            InputLocationTextBox.Text = _inputFilePath;
        }

        private void SaveLocationButton_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new OpenFolderDialog();

            if (folderDialog.ShowDialog() != true)
            {
                return;
            }

            _outputFilePath = folderDialog.FolderName;
            SaveLocationTextBox.Text = _outputFilePath;
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_inputFilePath)
                || string.IsNullOrEmpty(_outputFilePath))
            {
                MessageBox.Show(
                    $"Please select both input and save locations.",
                    "Input locations missing",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return;
            }

            var incrementorParsingResult = Logic.Incrementor
                .ProcessData(_inputFilePath, Path.Combine(_outputFilePath,
                    $"Output-{Guid.NewGuid().ToString()}.xlsx"));
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