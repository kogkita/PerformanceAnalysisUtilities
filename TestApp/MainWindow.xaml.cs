using Microsoft.Win32;
using OfficeOpenXml;
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
        private readonly List<string> jtlSelectedFiles = new();
        private Button? activeNavButton;

        private bool _isManuallyMaximized = false;

        public MainWindow()
        {
            InitializeComponent();
            activeNavButton = NavConvert;
            Loaded += (_, _) =>
            {
                UpdateBLGUI();
                _isManuallyMaximized = true;
                ApplyMaximizedLayout();
            };
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private void ApplyMaximizedLayout()
        {
            // Get the real working area in physical pixels via SystemParameters
            // then convert to WPF device-independent units using DPI scale
            var source = PresentationSource.FromVisual(this);
            double dpiX = 1.0, dpiY = 1.0;
            if (source?.CompositionTarget != null)
            {
                dpiX = source.CompositionTarget.TransformFromDevice.M11;
                dpiY = source.CompositionTarget.TransformFromDevice.M22;
            }

            var area = SystemParameters.WorkArea;

            // Apply position and size manually — no WindowState.Maximized so no resize border overhang
            WindowState = WindowState.Normal;
            Left = area.Left;
            Top = area.Top;
            Width = area.Width;
            Height = area.Height;

            MaxWidth = area.Width;
            MaxHeight = area.Height;

            RootBorder.Margin = new Thickness(0);
            RootBorder.CornerRadius = new CornerRadius(0);
            RootBorder.BorderThickness = new Thickness(0);
            MaxRestoreBtn.Content = "\uE923";
            MaxRestoreBtn.ToolTip = "Restore";
        }

        private void ApplyRestoredLayout()
        {
            MaxWidth = double.PositiveInfinity;
            MaxHeight = double.PositiveInfinity;
            Width = 1280;
            Height = 720;
            WindowStartupLocation = WindowStartupLocation.Manual;
            var area = SystemParameters.WorkArea;
            Left = area.Left + (area.Width - 1280) / 2;
            Top = area.Top + (area.Height - 720) / 2;

            RootBorder.Margin = new Thickness(0);
            RootBorder.CornerRadius = new CornerRadius(12);
            RootBorder.BorderThickness = new Thickness(1);
            MaxRestoreBtn.Content = "\uE922";
            MaxRestoreBtn.ToolTip = "Maximize";
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
            if (_isManuallyMaximized)
            {
                _isManuallyMaximized = false;
                ApplyRestoredLayout();
            }
            else
            {
                _isManuallyMaximized = true;
                ApplyMaximizedLayout();
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
            PageJTL.Visibility = Visibility.Collapsed;
            PageBLG.Visibility = Visibility.Collapsed;
            PageNmon.Visibility = Visibility.Collapsed;
            PageCompare.Visibility = Visibility.Collapsed;
            page.Visibility = Visibility.Visible;
        }

        private void NavConvert_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavConvert, PageConvert);

        private void NavJTL_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavJTL, PageJTL);

        private void NavBLG_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavBLG, PageBLG);

        private void NavNmon_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavNmon, PageNmon);

        private void NavCompare_Click(object sender, RoutedEventArgs e)
        {
            SetActivePage(NavCompare, PageCompare);
            if (CmpFileRowsPanel.Children.Count == 0)
                CmpRebuildRows();
        }

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

        private async void RunProcessing_Click(object sender, RoutedEventArgs e)
        {
            if (selectedFiles.Count == 0)
            {
                MessageBox.Show("Please select or drop one or more CSV files first.",
                    "No Files", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool club = ClubOutputCheckbox.IsChecked == true;
            bool includeCharts = IncludeChartsCheckbox.IsChecked == true;

            if (club)
            {
                await RunConvertClubbed(includeCharts);
            }
            else
            {
                ShowLogPanel(ConvertLogPanel, ConvertProgress, ConvertLog);
                LogInfo(ConvertLog, $"Processing {selectedFiles.Count} file(s)…");

                var files = selectedFiles.ToList();
                int succeeded = 0;
                var errors = new List<string>();

                await System.Threading.Tasks.Task.Run(() =>
                {
                    foreach (var csvPath in files)
                    {
                        try
                        {
                            var output = Path.ChangeExtension(csvPath, ".xlsx");
                            ResponseTimeConverter.Convert(csvPath, output, includeCharts);
                            succeeded++;
                            Dispatcher.Invoke(() => LogMsg(ConvertLog, $"✓ {Path.GetFileName(csvPath)} → {Path.GetFileName(output)}"));
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"{Path.GetFileName(csvPath)}: {ex.Message}");
                        }
                    }
                });

                LogResult(ConvertLog, ConvertProgress, succeeded, errors);
            }
        }

        private async System.Threading.Tasks.Task RunConvertClubbed(bool includeCharts)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Save Combined Excel Workbook",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "ResponseTimes_Combined.xlsx"
            };
            if (dlg.ShowDialog() != true)
            {
                HideLogPanel(ConvertLogPanel, ConvertProgress);
                return;
            }

            ShowLogPanel(ConvertLogPanel, ConvertProgress, ConvertLog);
            LogInfo(ConvertLog, $"Processing {selectedFiles.Count} file(s)…");

            var files = selectedFiles.ToList();
            var outputPath = dlg.FileName;
            var errors = new List<string>();
            int succeeded = 0;

            await System.Threading.Tasks.Task.Run(() =>
            {
                ResponseTimeConverter.ClearPendingCharts();
                ExcelPackage.License.SetNonCommercialPersonal("Response Time Converter");
                using var package = new ExcelPackage();

                foreach (var csvPath in files)
                {
                    try
                    {
                        string prefix = SanitizeSheetName(Path.GetFileNameWithoutExtension(csvPath), 20);
                        ResponseTimeConverter.AppendToPackage(package, csvPath, prefix, includeCharts);
                        succeeded++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{Path.GetFileName(csvPath)}: {ex.Message}");
                    }
                }

                if (succeeded > 0)
                {
                    package.SaveAs(new FileInfo(outputPath));
                    if (includeCharts)
                        ResponseTimeConverter.InjectPendingCharts(outputPath);
                }
            });

            LogResult(ConvertLog, ConvertProgress, succeeded, errors, outputPath);
        }
        // ── JTL File Processing page ─────────────────────────

        private void JTLAddFiles(IEnumerable<string> paths)
        {
            foreach (var path in paths)
            {
                var ext = Path.GetExtension(path).ToLowerInvariant();
                if (ext != ".jtl") continue;
                if (jtlSelectedFiles.Contains(path)) continue;
                jtlSelectedFiles.Add(path);
            }
            JTLRefreshFileList();
        }

        private void JTLRefreshFileList()
        {
            JTLFileListPanel.Children.Clear();

            foreach (var path in jtlSelectedFiles)
            {
                var row = new Grid { Margin = new Thickness(4, 2, 4, 2) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

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
                    jtlSelectedFiles.Remove(capturedPath);
                    JTLRefreshFileList();
                };
                Grid.SetColumn(removeBtn, 1);

                row.Children.Add(label);
                row.Children.Add(removeBtn);
                JTLFileListPanel.Children.Add(row);
            }

            int count = jtlSelectedFiles.Count;
            JTLFileCountLabel.Text = count == 0
                ? "No files selected"
                : count == 1
                    ? "1 file selected"
                    : $"{count} files selected";

            JTLClearAllButton.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void JTLBrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "JTL Files (*.jtl)|*.jtl",
                Multiselect = true
            };
            if (dialog.ShowDialog() == true)
                JTLAddFiles(dialog.FileNames);
        }

        private void JTLClearAll_Click(object sender, RoutedEventArgs e)
        {
            jtlSelectedFiles.Clear();
            JTLRefreshFileList();
        }

        private void JTLFileDropped(object sender, DragEventArgs e)
        {
            JTLResetDropZone();
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                JTLAddFiles(files);
            }
        }

        private void JTLDropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                JTLDropZone.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB"));
                JTLDropZone.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#111827"));
            }
        }

        private void JTLDropZone_DragLeave(object sender, DragEventArgs e)
            => JTLResetDropZone();

        private void JTLResetDropZone()
        {
            JTLDropZone.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640"));
            JTLDropZone.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0D1020"));
        }

        private async void JTLRunProcessing_Click(object sender, RoutedEventArgs e)
        {
            if (jtlSelectedFiles.Count == 0)
            {
                MessageBox.Show("Please select or drop one or more JTL files first.",
                    "No Files", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool club = JTLClubOutputCheckbox.IsChecked == true;
            bool includeCharts = JTLIncludeChartsCheckbox.IsChecked == true;

            if (club)
            {
                await RunJTLClubbed(includeCharts);
            }
            else
            {
                ShowLogPanel(JTLLogPanel, JTLProgress, JTLLog);
                LogInfo(JTLLog, $"Processing {jtlSelectedFiles.Count} file(s)…");

                var files = jtlSelectedFiles.ToList();
                int succeeded = 0;
                var errors = new List<string>();

                await System.Threading.Tasks.Task.Run(() =>
                {
                    foreach (var jtlPath in files)
                    {
                        try
                        {
                            var output = Path.ChangeExtension(jtlPath, ".xlsx");
                            JTLFileProcessing.Convert(jtlPath, output, includeCharts);
                            succeeded++;
                            Dispatcher.Invoke(() => LogMsg(JTLLog, $"✓ {Path.GetFileName(jtlPath)} → {Path.GetFileName(output)}"));
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"{Path.GetFileName(jtlPath)}: {ex.Message}");
                        }
                    }
                });

                LogResult(JTLLog, JTLProgress, succeeded, errors);
            }
        }

        private async System.Threading.Tasks.Task RunJTLClubbed(bool includeCharts)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Save Combined JTL Excel Workbook",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "JTLResults_Combined.xlsx"
            };
            if (dlg.ShowDialog() != true)
            {
                HideLogPanel(JTLLogPanel, JTLProgress);
                return;
            }

            ShowLogPanel(JTLLogPanel, JTLProgress, JTLLog);
            LogInfo(JTLLog, $"Processing {jtlSelectedFiles.Count} file(s)…");

            var files = jtlSelectedFiles.ToList();
            var outputPath = dlg.FileName;
            var errors = new List<string>();
            int succeeded = 0;

            await System.Threading.Tasks.Task.Run(() =>
            {
                JTLFileProcessing.ClearPendingCharts();
                ExcelPackage.License.SetNonCommercialPersonal("JTL File Processing");
                using var package = new ExcelPackage();

                foreach (var jtlPath in files)
                {
                    try
                    {
                        string prefix = SanitizeSheetName(Path.GetFileNameWithoutExtension(jtlPath), 20);
                        JTLFileProcessing.AppendToPackage(package, jtlPath, prefix, includeCharts);
                        succeeded++;
                        Dispatcher.Invoke(() => LogMsg(JTLLog, $"✓ {Path.GetFileName(jtlPath)} added to workbook"));
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{Path.GetFileName(jtlPath)}: {ex.Message}");
                    }
                }

                if (succeeded > 0)
                {
                    package.SaveAs(new FileInfo(outputPath));
                    if (includeCharts)
                        JTLFileProcessing.InjectPendingCharts(outputPath);
                }
            });

            LogResult(JTLLog, JTLProgress, succeeded, errors, outputPath);
        }

        // ── Shared helpers ───────────────────────────────────

        private static string SanitizeSheetName(string name, int maxLen)
        {
            var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (var c in invalid) name = name.Replace(c, '_');
            return name.Length > maxLen ? name[..maxLen] : name;
        }

        // ── BLG File Conversion page ──────────────────────────

        private readonly List<string> blgSelectedFiles = new();
        private string? blgCustomCounterFile = null;

        private void BLGAddFiles(IEnumerable<string> paths)
        {
            foreach (var path in paths)
            {
                if (!path.EndsWith(".blg", StringComparison.OrdinalIgnoreCase)) continue;
                if (blgSelectedFiles.Contains(path)) continue;

                blgSelectedFiles.Add(path);

                var row = new Grid { Margin = new Thickness(0, 2, 0, 2) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var lbl = new TextBlock
                {
                    Text = System.IO.Path.GetFileName(path),
                    Foreground = new SolidColorBrush(Color.FromRgb(0xCB, 0xD5, 0xE1)),
                    FontSize = 12,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(6, 0, 0, 0),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = path
                };
                Grid.SetColumn(lbl, 0);

                var removeBtn = new Button
                {
                    Content = "\uE711",
                    Style = (Style)Resources["RemoveButtonStyle"],
                    Tag = path,
                    Margin = new Thickness(4, 0, 4, 0)
                };
                removeBtn.Click += (s, _) =>
                {
                    var p = (string)((Button)s).Tag;
                    blgSelectedFiles.Remove(p);
                    BLGFileListPanel.Children.Remove(row);
                    UpdateBLGUI();
                };
                Grid.SetColumn(removeBtn, 1);

                row.Children.Add(lbl);
                row.Children.Add(removeBtn);
                BLGFileListPanel.Children.Add(row);
            }
            UpdateBLGUI();
        }

        private void UpdateBLGUI()
        {
            int count = blgSelectedFiles.Count;
            BLGFileCountLabel.Text = count == 0
                ? "No files selected"
                : count == 1 ? "1 file selected" : $"{count} files selected";
            BLGClearAllButton.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;

            RefreshBLGCommandPreview();
            RefreshBLGCounterPreview();

            // Keep server label inputs in sync with file list
            if (BLGProduceGraphsCheckbox?.IsChecked == true)
                RebuildBLGServerLabelInputs();
        }

        private BlgServerType SelectedBlgServerType =>
            BLGRadioDb?.IsChecked == true ? BlgServerType.DbServer : BlgServerType.AppServer;

        private void RefreshBLGCommandPreview()
        {
            if (BLGCommandPreview == null) return;
            var opts = BuildBlgOptions(blgSelectedFiles.FirstOrDefault() ?? string.Empty);
            BLGCommandPreview.Text = BLGConverter.BuildCommandPreview(opts);
        }

        private void RefreshBLGCounterPreview()
        {
            if (BLGCounterPreviewList == null) return;
            var opts = BuildBlgOptions(string.Empty);
            BLGCounterPreviewList.ItemsSource = BLGConverter.PreviewCounters(opts);
        }

        private BlgConvertOptions BuildBlgOptions(string blgPath) => new()
        {
            BlgPath = blgPath,
            ServerType = SelectedBlgServerType,
            CustomCounterFilePath = blgCustomCounterFile,
        };

        private void BLGServerType_Changed(object sender, RoutedEventArgs e)
            => UpdateBLGUI();

        private void BLGProduceGraphs_Changed(object sender, RoutedEventArgs e)
        {
            bool show = BLGProduceGraphsCheckbox.IsChecked == true;
            BLGServerLabelPanel.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            if (show) RebuildBLGServerLabelInputs();
        }

        private void RebuildBLGServerLabelInputs()
        {
            BLGServerLabelInputs.Children.Clear();
            foreach (var path in blgSelectedFiles)
            {
                var row = new Grid { Margin = new Thickness(0, 4, 0, 0) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var lbl = new TextBlock
                {
                    Text = System.IO.Path.GetFileNameWithoutExtension(path),
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8B93A5")),
                    FontSize = 12,
                    FontFamily = new FontFamily("Consolas, Segoe UI Mono, Segoe UI"),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = path
                };
                Grid.SetColumn(lbl, 0);

                var box = new TextBox
                {
                    Tag = path,
                    ToolTip = $"Label for {System.IO.Path.GetFileNameWithoutExtension(path)}",
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0F1117")),
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E8F0")),
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2A2F3E")),
                    CaretBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E8F0")),
                    FontSize = 12,
                    FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                    Height = 30,
                    Padding = new Thickness(8, 0, 8, 0),
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(10, 0, 0, 0)
                };
                Grid.SetColumn(box, 1);

                row.Children.Add(lbl);
                row.Children.Add(box);
                BLGServerLabelInputs.Children.Add(row);
            }
        }

        private List<string> GetBLGServerLabels()
        {
            var labels = new List<string>();
            foreach (var child in BLGServerLabelInputs.Children)
            {
                if (child is Grid row)
                {
                    var box = row.Children.OfType<TextBox>().FirstOrDefault();
                    labels.Add(box?.Text?.Trim() ?? "");
                }
            }
            return labels;
        }

        private void BLGCounterFileBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Select counter filter file",
                Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
            };
            if (dlg.ShowDialog() != true) return;

            blgCustomCounterFile = dlg.FileName;
            BLGCounterFileLabel.Text = System.IO.Path.GetFileName(dlg.FileName);
            BLGCounterFileLabel.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#7DD3FC"));
            BLGCounterFileClearBtn.Visibility = Visibility.Visible;
            UpdateBLGUI();
        }

        private void BLGCounterFileClear_Click(object sender, RoutedEventArgs e)
        {
            blgCustomCounterFile = null;
            BLGCounterFileLabel.Text = "Using default template";
            BLGCounterFileLabel.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#4A5568"));
            BLGCounterFileClearBtn.Visibility = Visibility.Collapsed;
            UpdateBLGUI();
        }

        private void BLGBrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Select BLG file(s)",
                Filter = "Performance Monitor Log (*.blg)|*.blg|All Files (*.*)|*.*",
                Multiselect = true
            };
            if (dlg.ShowDialog() == true)
                BLGAddFiles(dlg.FileNames);
        }

        private void BLGClearAll_Click(object sender, RoutedEventArgs e)
        {
            blgSelectedFiles.Clear();
            BLGFileListPanel.Children.Clear();
            UpdateBLGUI();
        }

        private void BLGFileDropped(object sender, DragEventArgs e)
        {
            BLGDropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(0x1E, 0x26, 0x40));
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                BLGAddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
        }

        private void BLGDropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                BLGDropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(0x25, 0x63, 0xEB));
        }

        private void BLGDropZone_DragLeave(object sender, DragEventArgs e)
            => BLGDropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(0x1E, 0x26, 0x40));

        private void BLGRunProcessing_Click(object sender, RoutedEventArgs e)
        {
            if (blgSelectedFiles.Count == 0)
            {
                MessageBox.Show("Please select at least one .blg file.",
                    "No Files Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool produceGraphs = BLGProduceGraphsCheckbox.IsChecked == true;

            ShowLogPanel(BLGLogPanel, BLGProgress, BLGLog);
            LogInfo(BLGLog, $"Converting {blgSelectedFiles.Count} file(s)…");
            BLGStatusLabel.Text = $"Converting {blgSelectedFiles.Count} file(s)…";
            BLGStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x60, 0xA5, 0xFA));

            var filesToProcess = blgSelectedFiles.ToList();
            var serverType = SelectedBlgServerType;
            var customCf = blgCustomCounterFile;
            var serverLabels = produceGraphs ? GetBLGServerLabels() : new List<string>();

            System.Threading.Tasks.Task.Run(() =>
            {
                var succeeded = new List<string>();  // CSV paths
                var errors = new List<string>();

                foreach (var blgPath in filesToProcess)
                {
                    try
                    {
                        var opts = new BlgConvertOptions
                        {
                            BlgPath = blgPath,
                            ServerType = serverType,
                            CustomCounterFilePath = customCf,
                        };
                        string csv = BLGConverter.ConvertToCsv(opts);
                        succeeded.Add(csv);
                        Dispatcher.Invoke(() => LogMsg(BLGLog, $"✓ {System.IO.Path.GetFileName(blgPath)} → {System.IO.Path.GetFileName(csv)}"));
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{System.IO.Path.GetFileName(blgPath)}: {ex.Message}");
                    }
                }

                // ── Produce graphs if requested and at least one CSV was made ──
                if (produceGraphs && succeeded.Count > 0)
                {
                    Dispatcher.Invoke(() => LogMsg(BLGLog, "Generating Excel charts…", "#60A5FA"));

                    try
                    {
                        // Save graphs workbook alongside the first CSV
                        string outDir = System.IO.Path.GetDirectoryName(succeeded[0])!;
                        string graphOut = System.IO.Path.Combine(outDir, "BLG_Graphs.xlsx");

                        BLGGraphProducer.ProduceGraphs(succeeded, graphOut, serverLabels);
                        Dispatcher.Invoke(() => LogMsg(BLGLog, $"✓ Charts saved → {System.IO.Path.GetFileName(graphOut)}"));
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() => LogError(BLGLog, $"Graph generation failed: {ex.Message}"));
                    }
                }

                Dispatcher.Invoke(() =>
                {
                    HideProgress(BLGProgress);
                    if (errors.Count == 0)
                    {
                        string suffix = produceGraphs && succeeded.Count > 0 ? " + charts" : "";
                        BLGStatusLabel.Text = $"Done — {succeeded.Count} CSV file(s) created{suffix}.";
                        BLGStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
                        LogSuccess(BLGLog, $"Done — {succeeded.Count} CSV file(s) created successfully{suffix}.");
                    }
                    else
                    {
                        BLGStatusLabel.Text = $"Completed with {errors.Count} error(s).";
                        BLGStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71));
                        if (succeeded.Count > 0)
                            LogMsg(BLGLog, $"{succeeded.Count} succeeded, {errors.Count} failed:", "#FBBF24");
                        else
                            LogError(BLGLog, "All conversions failed:");
                        foreach (var err in errors)
                            LogError(BLGLog, $"  • {err}");
                    }
                });
            });
        }

        // ── nmon Analyzer page ────────────────────────────────

        private readonly List<string> nmonSelectedFiles = new();


        private void NmonAddFiles(IEnumerable<string> paths)
        {
            foreach (var path in paths)
            {
                string ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
                if (ext != ".nmon" && ext != ".csv") continue;
                if (nmonSelectedFiles.Contains(path)) continue;

                nmonSelectedFiles.Add(path);

                var row = new Grid { Margin = new Thickness(0, 2, 0, 2) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var lbl = new TextBlock
                {
                    Text = System.IO.Path.GetFileName(path),
                    Foreground = new SolidColorBrush(Color.FromRgb(0xCB, 0xD5, 0xE1)),
                    FontSize = 12,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(6, 0, 0, 0),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = path
                };
                Grid.SetColumn(lbl, 0);

                var removeBtn = new Button
                {
                    Content = "\uE711",
                    Style = (Style)Resources["RemoveButtonStyle"],
                    Tag = path,
                    Margin = new Thickness(4, 0, 4, 0)
                };
                removeBtn.Click += (s, _) =>
                {
                    var p = (string)((Button)s).Tag;
                    nmonSelectedFiles.Remove(p);
                    NmonFileListPanel.Children.Remove(row);
                    UpdateNmonUI();
                };
                Grid.SetColumn(removeBtn, 1);

                row.Children.Add(lbl);
                row.Children.Add(removeBtn);
                NmonFileListPanel.Children.Add(row);
            }
            UpdateNmonUI();
        }

        private void UpdateNmonUI()
        {
            int count = nmonSelectedFiles.Count;
            NmonFileCountLabel.Text = count == 0 ? "No files selected"
                : count == 1 ? "1 file selected" : $"{count} files selected";
            NmonClearAllButton.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void NmonBrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Select NMON file(s)",
                Filter = "NMON files (*.nmon;*.csv)|*.nmon;*.csv|All Files (*.*)|*.*",
                Multiselect = true
            };
            if (dlg.ShowDialog() == true) NmonAddFiles(dlg.FileNames);
        }

        private void NmonClearAll_Click(object sender, RoutedEventArgs e)
        {
            nmonSelectedFiles.Clear();
            NmonFileListPanel.Children.Clear();
            UpdateNmonUI();
        }

        private void NmonFileDropped(object sender, DragEventArgs e)
        {
            NmonDropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(0x1E, 0x26, 0x40));
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                NmonAddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
        }

        private void NmonDropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                NmonDropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(0x25, 0x63, 0xEB));
        }

        private void NmonDropZone_DragLeave(object sender, DragEventArgs e)
            => NmonDropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(0x1E, 0x26, 0x40));

        private void NmonOutDirBrowse_Click(object sender, RoutedEventArgs e)
        {
            // WPF has no native FolderBrowserDialog — use SaveFileDialog pointed at a folder
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Select output directory (type any filename, only the folder is used)",
                Filter = "Folder|*.folder",
                FileName = "Select Folder",
                CheckFileExists = false,
                CheckPathExists = true,
                ValidateNames = false
            };
            if (dlg.ShowDialog() == true)
                NmonOutDirBox.Text = System.IO.Path.GetDirectoryName(dlg.FileName) ?? string.Empty;
        }

        private void NmonXlsmBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Locate nmon_analyser_v69_2.xlsm",
                Filter = "nmon Analyser (*.xlsm)|*.xlsm|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
                NmonXlsmPathBox.Text = dlg.FileName;
        }

        private void NmonRunAnalysis_Click(object sender, RoutedEventArgs e)
        {
            if (nmonSelectedFiles.Count == 0)
            {
                MessageBox.Show("Please select at least one .nmon file.",
                    "No Files Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Resolve XLSM path — check box first, then app directory
            string xlsmPath = NmonXlsmPathBox.Text.Trim();
            if (string.IsNullOrEmpty(xlsmPath))
            {
                var appDir = System.IO.Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "";
                var candidate = System.IO.Path.Combine(appDir, "nmon_analyser_v69_2.xlsm");
                if (System.IO.File.Exists(candidate))
                    xlsmPath = candidate;
            }

            if (!System.IO.File.Exists(xlsmPath))
            {
                MessageBox.Show(
                    "Cannot find nmon_analyser_v69_2.xlsm.\n\n" +
                    "Please browse to locate it using the 'Browse…' button, " +
                    "or place it in the same folder as this application.",
                    "XLSM Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Build options from UI
                var opts = BuildNmonOptions(xlsmPath);

                ShowLogPanel(NmonLogPanel, NmonProgress, NmonLog);
                LogInfo(NmonLog, "Running analysis… Excel will open in the background.");
                NmonStatusLabel.Text = "Running analysis… Excel will open in the background.";
                NmonStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x60, 0xA5, 0xFA));

                // Run on background thread so UI stays responsive
                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        NmonAnalyzer.Run(opts);
                        Dispatcher.Invoke(() =>
                        {
                            HideProgress(NmonProgress);
                            NmonStatusLabel.Text = "Analysis complete. Check the output directory for Excel files.";
                            NmonStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
                            string outDir = string.IsNullOrEmpty(opts.OutDir) ? "Same directory as each input file" : opts.OutDir;
                            LogSuccess(NmonLog, $"Analysis complete — Excel files saved to: {outDir}");
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            HideProgress(NmonProgress);
                            NmonStatusLabel.Text = $"Error: {ex.Message}";
                            NmonStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71));
                            LogError(NmonLog, $"Analysis failed: {ex.Message}");
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                LogError(NmonLog, $"Failed to start analysis: {ex.Message}");
            }
        }

        private NmonAnalyzerOptions BuildNmonOptions(string xlsmPath)
        {
            // Parse GRAPHS combo: "ALL|CHARTS" etc.
            var graphsTag = ((NmonGraphsCombo.SelectedItem as ComboBoxItem)?.Tag as string ?? "ALL|CHARTS")
                .Split('|');

            return new NmonAnalyzerOptions
            {
                XlsmPath = xlsmPath,
                NmonFiles = nmonSelectedFiles.ToList(),
                GraphsScope = graphsTag.Length > 0 ? graphsTag[0] : "ALL",
                GraphsOutput = graphsTag.Length > 1 ? graphsTag[1] : "CHARTS",
                Merge = (NmonMergeCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "NO",
                IntervalFirst = NmonIntervalFirst.Text.Trim(),
                IntervalLast = NmonIntervalLast.Text.Trim(),
                Ess = NmonEssChk.IsChecked == true,
                Scatter = NmonScatterChk.IsChecked == true,
                BigData = NmonBigdataChk.IsChecked == true,
                ShowLinuxCpuUtil = NmonLinuxCpuChk.IsChecked == true,
                Reorder = NmonReorderChk.IsChecked == true,
                SortDefault = NmonSortDefaultChk.IsChecked == true,
                List = NmonListBox.Text.Trim(),
                OutDir = NmonOutDirBox.Text.Trim(),
            };
        }

        // ── Log panel helpers ─────────────────────────────────

        private void ShowLogPanel(Border panel, ProgressBar progress, TextBlock log)
        {
            log.Text = "";
            panel.Visibility = Visibility.Visible;
            progress.Visibility = Visibility.Visible;
        }

        private void HideLogPanel(Border panel, ProgressBar progress)
        {
            progress.Visibility = Visibility.Collapsed;
            panel.Visibility = Visibility.Collapsed;
        }

        private void HideProgress(ProgressBar progress)
        {
            progress.Visibility = Visibility.Collapsed;
        }

        private void LogMsg(TextBlock log, string message, string colorHex = "#8B93A5")
        {
            string ts = DateTime.Now.ToString("HH:mm:ss");
            if (log.Text.Length > 0) log.Text += "\n";
            log.Text += $"[{ts}]  {message}";
            log.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));
        }

        private void LogSuccess(TextBlock log, string message) => LogMsg(log, message, "#4ADE80");
        private void LogError(TextBlock log, string message) => LogMsg(log, message, "#F87171");
        private void LogInfo(TextBlock log, string message) => LogMsg(log, message, "#60A5FA");

        private void LogResult(TextBlock log, ProgressBar progress, int succeeded, List<string> errors, string? savedPath = null)
        {
            HideProgress(progress);
            if (errors.Count == 0)
            {
                string msg = savedPath != null
                    ? $"Done — Combined workbook saved to: {savedPath}"
                    : succeeded == 1
                        ? "Done — Excel file created successfully."
                        : $"Done — {succeeded} Excel files created successfully.";
                LogSuccess(log, msg);
            }
            else
            {
                if (succeeded > 0)
                    LogMsg(log, $"{succeeded} file(s) processed. {errors.Count} failed:", "#FBBF24");
                else
                    LogError(log, $"All processing failed:");
                foreach (var err in errors)
                    LogError(log, $"  • {err}");
            }
        }

        // ── Run Comparison page ──────────────────────────────

        private readonly List<string> cmpRunFiles = new();

        // Called on page load and whenever files change — rebuilds the row list
        private void CmpRebuildRows()
        {
            CmpFileRowsPanel.Children.Clear();

            // Always show at least 2 rows (Baseline + Run 2)
            int displayCount = Math.Max(2, cmpRunFiles.Count);

            for (int i = 0; i < displayCount; i++)
            {
                string path = i < cmpRunFiles.Count ? cmpRunFiles[i] : string.Empty;
                bool isBaseline = i == 0;
                bool hasFile = !string.IsNullOrEmpty(path);
                int capturedIdx = i;

                // Badge label
                string badgeText = isBaseline ? "Baseline" : $"Run {i + 1}";
                string badgeBg = isBaseline ? "#1E3A8A" : "#1E293B";
                string badgeFg = isBaseline ? "#93C5FD" : "#7DD3FC";

                // Outer row grid: badge | drop-zone border (path label) | browse btn | clear btn
                var row = new Grid { Margin = new Thickness(0, 0, 0, 10) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                // Badge
                var badge = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(badgeBg)),
                    CornerRadius = new CornerRadius(5),
                    Padding = new Thickness(0),
                    Margin = new Thickness(0, 0, 10, 0),
                    Height = 40,
                    VerticalAlignment = VerticalAlignment.Center
                };
                badge.Child = new TextBlock
                {
                    Text = badgeText,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(badgeFg)),
                    FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(badge, 0);

                // Path label border (also acts as drop target)
                var pathBorder = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#161B2A")),
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252D42")),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(6),
                    Height = 40,
                    AllowDrop = true
                };
                var pathLabel = new TextBlock
                {
                    Text = hasFile ? path : "No file selected — browse or drag & drop here",
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                                            hasFile ? "#7DD3FC" : "#8B93A5")),
                    FontSize = 12,
                    FontFamily = new FontFamily("Consolas, Segoe UI Mono, Segoe UI"),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    Padding = new Thickness(12, 0, 12, 0),
                    ToolTip = hasFile ? path : null
                };
                pathBorder.Child = pathLabel;

                // Drag-drop on the path border
                pathBorder.DragEnter += (s, e) =>
                {
                    if (e.Data.GetDataPresent(DataFormats.FileDrop))
                        ((Border)s).BorderBrush = new SolidColorBrush(
                            (Color)ColorConverter.ConvertFromString("#2563EB"));
                };
                pathBorder.DragLeave += (s, e) =>
                {
                    ((Border)s).BorderBrush = new SolidColorBrush(
                        (Color)ColorConverter.ConvertFromString("#252D42"));
                };
                pathBorder.Drop += (s, e) =>
                {
                    ((Border)s).BorderBrush = new SolidColorBrush(
                        (Color)ColorConverter.ConvertFromString("#252D42"));
                    if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
                        CmpSetFile(capturedIdx, files[0]);
                };
                Grid.SetColumn(pathBorder, 1);

                // Browse button
                var browseBtn = new Button
                {
                    Content = "Browse\u2026",
                    Width = 90,
                    Height = 40,
                    Margin = new Thickness(10, 0, 0, 0),
                    FontSize = 13,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Brushes.White,
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB")),
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand
                };
                browseBtn.Style = (Style)Resources["ActionButtonStyle"];
                browseBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB"));
                browseBtn.Foreground = Brushes.White;
                browseBtn.Click += (_, _) => CmpBrowseRow(capturedIdx);
                Grid.SetColumn(browseBtn, 2);

                // Clear button (only visible when file is set, and only for rows beyond the first two)
                var clearBtn = new Button
                {
                    Content = "Clear",
                    Width = 60,
                    Height = 40,
                    Margin = new Thickness(8, 0, 0, 0),
                    FontSize = 12,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CA3AF")),
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#374151")),
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    Visibility = hasFile ? Visibility.Visible : Visibility.Collapsed
                };
                clearBtn.Style = (Style)Resources["ActionButtonStyle"];
                clearBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#374151"));
                clearBtn.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CA3AF"));
                clearBtn.Click += (_, _) =>
                {
                    // If this is a "bonus" row (index >= 2) and empty after clearing — remove it
                    if (capturedIdx >= 2)
                    {
                        cmpRunFiles.RemoveAt(capturedIdx);
                    }
                    else
                    {
                        // For baseline / run 2, just clear the path
                        if (capturedIdx < cmpRunFiles.Count)
                            cmpRunFiles[capturedIdx] = string.Empty;
                    }
                    CmpRebuildRows();
                };
                Grid.SetColumn(clearBtn, 3);

                row.Children.Add(badge);
                row.Children.Add(pathBorder);
                row.Children.Add(browseBtn);
                row.Children.Add(clearBtn);
                CmpFileRowsPanel.Children.Add(row);
            }

            CmpUpdateFileTypeLabel();
        }

        /// <summary>
        /// Updates the auto-detected "File type" label based on the extensions
        /// of the files currently selected in the comparison rows.
        /// </summary>
        private void CmpUpdateFileTypeLabel()
        {
            var paths = cmpRunFiles.Where(p => !string.IsNullOrEmpty(p)).ToList();
            if (paths.Count == 0)
            {
                CmpFileTypeLabel.Text = "No files selected";
                return;
            }

            bool hasCsv = paths.Any(p => p.EndsWith(".csv", StringComparison.OrdinalIgnoreCase));
            bool hasJtl = paths.Any(p => p.EndsWith(".jtl", StringComparison.OrdinalIgnoreCase));

            if (hasCsv && hasJtl)
                CmpFileTypeLabel.Text = "Mix of CSV & JTL";
            else if (hasJtl)
                CmpFileTypeLabel.Text = "JTL Only";
            else
                CmpFileTypeLabel.Text = "CSV Only";
        }

        private void CmpSetFile(int index, string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext != ".csv" && ext != ".jtl")
            {
                MessageBox.Show("Please select a CSV or JTL file.",
                    "Wrong File Type", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Grow the list if needed
            while (cmpRunFiles.Count <= index)
                cmpRunFiles.Add(string.Empty);

            cmpRunFiles[index] = path;
            CmpRebuildRows();
        }

        private void CmpBrowseRow(int index)
        {
            string label = index == 0 ? "Select Baseline File" : $"Select Run {index + 1} File";
            var dlg = new OpenFileDialog
            {
                Title = label,
                Filter = "Supported Files (*.csv;*.jtl)|*.csv;*.jtl|CSV Files (*.csv)|*.csv|JTL Files (*.jtl)|*.jtl|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() != true) return;
            CmpSetFile(index, dlg.FileName);
        }

        private void CmpAddRun_Click(object sender, RoutedEventArgs e)
        {
            // Add a new empty slot at the end — user browses or drags manually
            cmpRunFiles.Add(string.Empty);
            CmpRebuildRows();
        }

        // ── Run ──────────────────────────────────────────────────────────────

        private async void CmpRun_Click(object sender, RoutedEventArgs e)
        {
            // Collect non-empty paths in order
            var paths = cmpRunFiles.Where(p => !string.IsNullOrEmpty(p)).ToList();

            if (paths.Count < 2)
            {
                MessageBox.Show(
                    "Please select at least two files — the first is the baseline.",
                    "Not Enough Files", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var missing = paths.Where(f => !File.Exists(f)).ToList();
            if (missing.Count > 0)
            {
                MessageBox.Show(
                    $"These files no longer exist:\n\n{string.Join("\n", missing.Select(Path.GetFileName))}",
                    "Missing Files", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            double slaMs = 0;
            string slaText = CmpSlaTextBox.Text.Trim();
            if (!string.IsNullOrEmpty(slaText))
            {
                if (!double.TryParse(slaText,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out slaMs) || slaMs <= 0)
                {
                    MessageBox.Show("SLA threshold must be a positive number (milliseconds).",
                        "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            var baseName = Path.GetFileNameWithoutExtension(paths[0]);
            var curName = Path.GetFileNameWithoutExtension(paths[1]);
            string defName = paths.Count == 2
                ? $"Comparison_{baseName}_vs_{curName}.xlsx"
                : $"Comparison_{baseName}_vs_{paths.Count - 1}runs.xlsx";

            var saveDlg = new SaveFileDialog
            {
                Title = "Save Comparison Report",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = defName
            };
            if (saveDlg.ShowDialog() != true) return;

            var mode = CmpModeSequential.IsChecked == true
                ? ComparisonMode.Sequential
                : ComparisonMode.AllVsBaseline;

            ShowLogPanel(CmpLogPanel, CmpProgress, CmpLog);
            LogInfo(CmpLog, $"Comparing {paths.Count} files…");

            var outputPath = saveDlg.FileName;
            try
            {
                await System.Threading.Tasks.Task.Run(() =>
                    RunComparisonProcessor.Compare(paths, outputPath, slaMs, mode));
                HideProgress(CmpProgress);
                LogSuccess(CmpLog, $"Done — Comparison report saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                HideProgress(CmpProgress);
                LogError(CmpLog, $"Comparison failed: {ex.Message}");
            }
        }
    }
}
