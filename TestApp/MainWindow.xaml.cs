using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace TestApp
{
    public partial class MainWindow : Window
    {
        private readonly List<string> selectedFiles = new();
        private Button? activeNavButton;

        public MainWindow()
        {
            InitializeComponent();
            activeNavButton = NavConvert;
        }

        // ── Custom title bar ─────────────────────────────────

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
                ToggleMaximize();
            else
                DragMove();
        }

        private void MinimizeWindow_Click(object sender, RoutedEventArgs e)
            => WindowState = WindowState.Minimized;

        private void MaxRestoreWindow_Click(object sender, RoutedEventArgs e)
            => ToggleMaximize();

        private void CloseWindow_Click(object sender, RoutedEventArgs e)
            => Close();

        private void ToggleMaximize()
        {
            if (WindowState == WindowState.Maximized)
            {
                WindowState = WindowState.Normal;
                MaxRestoreBtn.Content = "\uE922";
            }
            else
            {
                WindowState = WindowState.Maximized;
                MaxRestoreBtn.Content = "\uE923";
            }
        }

        // ── Navigation ───────────────────────────────────────

        private void SetActivePage(Button clicked, UIElement page)
        {
            if (activeNavButton != null)
                activeNavButton.Style = (Style)Resources["NavButtonStyle"];

            clicked.Style = (Style)Resources["NavButtonActiveStyle"];
            activeNavButton = clicked;

            PageConvert.Visibility = Visibility.Collapsed;
            page.Visibility = Visibility.Visible;
        }

        private void NavConvert_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavConvert, PageConvert);

        // ── File list helpers ────────────────────────────────

        private void AddFiles(IEnumerable<string> paths)
        {
            foreach (var path in paths)
            {
                if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)
                    && !selectedFiles.Contains(path))
                {
                    selectedFiles.Add(path);
                }
            }
            RefreshFileList();
        }

        private void RefreshFileList()
        {
            FileListPanel.Children.Clear();

            foreach (var path in selectedFiles)
            {
                // Row grid
                var row = new Grid { Margin = new Thickness(4, 2, 4, 2) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                // File path label
                var label = new TextBlock
                {
                    Text = path,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#7DD3FC")),
                    FontSize = 11.5,
                    FontFamily = new FontFamily("Consolas, Segoe UI Mono, Segoe UI"),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = path
                };
                Grid.SetColumn(label, 0);

                // Remove button — capture path in closure
                var capturedPath = path;
                var removeBtn = new Button
                {
                    Width = 18,
                    Height = 18,
                    Content = "\uE711",
                    FontSize = 10,
                    FontFamily = new FontFamily("Segoe MDL2 Assets, Segoe UI"),
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4A5568")),
                    Cursor = Cursors.Hand,
                    ToolTip = "Remove",
                    Margin = new Thickness(6, 0, 2, 0),
                    VerticalAlignment = VerticalAlignment.Center
                };
                removeBtn.Click += (_, _) =>
                {
                    selectedFiles.Remove(capturedPath);
                    RefreshFileList();
                };
                Grid.SetColumn(removeBtn, 1);

                row.Children.Add(label);
                row.Children.Add(removeBtn);
                FileListPanel.Children.Add(row);
            }

            int count = selectedFiles.Count;
            FileCountLabel.Text = count == 0
                ? "No files selected"
                : count == 1
                    ? "1 file selected"
                    : $"{count} files selected";

            ClearAllButton.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        // ── Convert Response Times page ──────────────────────

        private void BrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv",
                Multiselect = true
            };
            if (dialog.ShowDialog() == true)
                AddFiles(dialog.FileNames);
        }

        private void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            selectedFiles.Clear();
            RefreshFileList();
        }

        private void FileDropped(object sender, DragEventArgs e)
        {
            ResetDropZone();
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                AddFiles(files);
            }
        }

        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                DropZone.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB"));
                DropZone.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#111827"));
            }
        }

        private void DropZone_DragLeave(object sender, DragEventArgs e)
            => ResetDropZone();

        private void ResetDropZone()
        {
            DropZone.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640"));
            DropZone.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0D1020"));
        }

        private void RunProcessing_Click(object sender, RoutedEventArgs e)
        {
            if (selectedFiles.Count == 0)
            {
                MessageBox.Show("Please select or drop one or more CSV files first.",
                    "No Files", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int succeeded = 0;
            var errors = new List<string>();

            foreach (var csvPath in selectedFiles)
            {
                try
                {
                    var output = Path.ChangeExtension(csvPath, ".xlsx");
                    ResponseTimeConverter.Convert(csvPath, output);
                    succeeded++;
                }
                catch (Exception ex)
                {
                    errors.Add($"{Path.GetFileName(csvPath)}: {ex.Message}");
                }
            }

            if (errors.Count == 0)
            {
                string msg = succeeded == 1
                    ? "Excel file created successfully."
                    : $"{succeeded} Excel files created successfully.";
                MessageBox.Show(msg, "Done", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                string msg = succeeded > 0
                    ? $"{succeeded} file(s) converted. {errors.Count} failed:\n\n{string.Join("\n", errors)}"
                    : $"All conversions failed:\n\n{string.Join("\n", errors)}";
                MessageBox.Show(msg, "Completed with Errors", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}