using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
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
                LoadLibrary();
                LoadTrendsLibrary();
                LoadSettings();
                LoadDbApiHostsToComboBox();
                LoadDbTrendsLibrary();
                InitTrayOnLoad();
                DarkMessageBox.SetOwner(this);
                Task.Run(CleanOrphanTempFiles); // clean up any leftover temp files from last session
                InitScriptParamPanel();
                InitAiChat();
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

            PageConvert.Visibility      = Visibility.Collapsed;
            PageJTL.Visibility          = Visibility.Collapsed;
            PageBLG.Visibility          = Visibility.Collapsed;
            PageNmon.Visibility         = Visibility.Collapsed;
            PageCompare.Visibility      = Visibility.Collapsed;
            PageScriptRunner.Visibility   = Visibility.Collapsed;
            TrendsSubPageLocal.Visibility = Visibility.Collapsed;
            TrendsSubPageDB.Visibility    = Visibility.Collapsed;
            PageSettings.Visibility       = Visibility.Collapsed;
            PageAiChat.Visibility         = Visibility.Collapsed;
            page.Visibility = Visibility.Visible;
        }

        private void NavScriptRunner_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavScriptRunner, PageScriptRunner);


        // ── Nav group collapse handlers ──────────────────────────────────────────

        private static void ToggleNavGroup(StackPanel panel, TextBlock chevron)
        {
            bool collapse = panel.Visibility == Visibility.Visible;
            panel.Visibility = collapse ? Visibility.Collapsed : Visibility.Visible;
            chevron.Text     = collapse ? "▸" : "▾";
        }

        private void NavGroupJmeter_Click(object sender, RoutedEventArgs e)
            => ToggleNavGroup(NavGroupJmeterPanel, NavGroupJmeterChevron);

        private void NavGroupSysmon_Click(object sender, RoutedEventArgs e)
            => ToggleNavGroup(NavGroupSysmonPanel, NavGroupSysmonChevron);

        private void NavGroupScripts_Click(object sender, RoutedEventArgs e)
            => ToggleNavGroup(NavGroupScriptsPanel, NavGroupScriptsChevron);

        private void NavGroupAutomation_Click(object sender, RoutedEventArgs e)
            => ToggleNavGroup(NavGroupAutomationPanel, NavGroupAutomationChevron);

        private void NavGroupAi_Click(object sender, RoutedEventArgs e)
            => ToggleNavGroup(NavGroupAiPanel, NavGroupAiChevron);

        private void NavTrendsLocal_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavTrendsLocal, TrendsSubPageLocal);

        private void NavTrendsDB_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavTrendsDB, TrendsSubPageDB);

        private void NavSettings_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavSettings, PageSettings);

        private void NavAiChat_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavAiChat, PageAiChat);

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
                DarkMessageBox.Show("Please select or drop one or more CSV files first.",
                    "No Files");
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
                DarkMessageBox.Show("Please select or drop one or more JTL files first.",
                    "No Files");
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
                DarkMessageBox.Show("Please select at least one .blg file.",
                    "No Files Selected");
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

        private void NmonRunAnalysis_Click(object sender, RoutedEventArgs e)
        {
            if (nmonSelectedFiles.Count == 0)
            {
                DarkMessageBox.Show("Please select at least one .nmon file.",
                    "No Files Selected");
                return;
            }

            // Determine output path
            string outDir = NmonOutDirBox.Text.Trim();
            if (string.IsNullOrEmpty(outDir))
                outDir = System.IO.Path.GetDirectoryName(nmonSelectedFiles[0]) ?? "";

            string firstName = System.IO.Path.GetFileNameWithoutExtension(nmonSelectedFiles[0]);
            string outputPath = System.IO.Path.Combine(outDir,
                nmonSelectedFiles.Count == 1
                    ? $"{firstName}_analysis.xlsx"
                    : $"nmon_analysis_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            var saveDlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save nmon Analysis Workbook",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = System.IO.Path.GetFileName(outputPath),
                InitialDirectory = outDir
            };
            if (saveDlg.ShowDialog() != true) return;

            var files = nmonSelectedFiles.ToList();

            ShowLogPanel(NmonLogPanel, NmonProgress, NmonLog);
            LogInfo(NmonLog, $"Analysing {files.Count} file(s)…");
            NmonStatusLabel.Text = "Running analysis…";
            NmonStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x60, 0xA5, 0xFA));

            var progress = new Progress<string>(msg => Dispatcher.Invoke(() => LogMsg(NmonLog, msg, "#60A5FA")));

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    NmonExcelProducer.Produce(files, saveDlg.FileName, progress);
                    Dispatcher.Invoke(() =>
                    {
                        HideProgress(NmonProgress);
                        NmonStatusLabel.Text = "Analysis complete.";
                        NmonStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
                        LogSuccess(NmonLog, $"Done — saved to: {saveDlg.FileName}");
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

        // ── Script Library ────────────────────────────────────

        private static readonly string LibraryPath = AppDataManager.ScriptLibraryPath;

        private List<ScriptEntry> _library = new();

        public class ScriptEntry
        {
            public string Id          { get; set; } = Guid.NewGuid().ToString("N")[..8];
            public string Name        { get; set; } = "";
            public string Description { get; set; } = "";
            public string ScriptPath  { get; set; } = "";
            public string Runtime     { get; set; } = "";
            public string Arguments   { get; set; } = "";
            public string WorkingDir  { get; set; } = "";
            public Dictionary<string, string> EnvVars { get; set; } = new();
            public DateTime SavedAt   { get; set; } = DateTime.Now;

            // ── Schedule ─────────────────────────────────────────
            public ScriptSchedule? Schedule   { get; set; } = null;
            public DateTime? LastRanAt        { get; set; } = null;
            public string?   LastExitCode     { get; set; } = null;
            /// <summary>Windows Task Scheduler task name — null if not registered.</summary>
            public string?   WinTaskName      { get; set; } = null;
            /// <summary>Path to the log file written by scheduled runs.</summary>
            public string?   ScheduledLogPath { get; set; } = null;
        }

        public class ScriptSchedule
        {
            public string Type      { get; set; } = "Once";
            public string TimeOfDay { get; set; } = "09:00";
            public int    DayOfWeek { get; set; } = 1;
            public string? RunOnce  { get; set; }
            public bool   Enabled   { get; set; } = true;
            public string? LogFilePath { get; set; } = null;
        }

        private void LoadLibrary()
        {
            try { _library = AppDataManager.LoadScriptLibrary(); }
            catch { _library = new List<ScriptEntry>(); }

            bool anyChanged = false;

            foreach (var entry in _library)
            {
                // Sync Windows Task existence
                if (!string.IsNullOrEmpty(entry.WinTaskName) &&
                    !WindowsTaskScheduler.TaskExists(entry.WinTaskName))
                {
                    entry.WinTaskName = null;
                    WindowsTaskScheduler.DeleteLauncher(entry.ScriptPath);
                    anyChanged = true;
                }

                // Sync LastRanAt from log file's last-write time
                // Task Scheduler runs independently so the app never sees the execution —
                // the log file timestamp is the only reliable indicator a run happened.
                if (!string.IsNullOrEmpty(entry.ScheduledLogPath) &&
                    System.IO.File.Exists(entry.ScheduledLogPath))
                {
                    var logWriteTime = System.IO.File.GetLastWriteTime(entry.ScheduledLogPath);
                    if (entry.LastRanAt == null || logWriteTime > entry.LastRanAt.Value)
                    {
                        entry.LastRanAt = logWriteTime;
                        anyChanged = true;

                        // Auto-disable Once schedule — it has fired
                        if (entry.Schedule?.Type == "Once" && entry.Schedule.Enabled)
                        {
                            entry.Schedule.Enabled = false;
                            anyChanged = true;
                        }
                    }
                }
            }

            if (anyChanged) SaveLibrary();

            RefreshLibraryUI();
            StartScheduler();
        }

        // ── In-app scheduler (fallback while app is open) ─────────────────────

        private void StartScheduler() { /* Windows Task Scheduler handles execution — nothing needed here */ }

        private string NextRunText(ScriptEntry entry)
        {
            if (entry.Schedule == null || !entry.Schedule.Enabled) return "";
            var s = entry.Schedule;
            var now = DateTime.Now;
            string schedText = s.Type switch
            {
                "Once"   => DateTime.TryParse(s.RunOnce, out var at)
                                ? (at > now ? $"Once at {at:dd MMM HH:mm}" : "Once (expired)")
                                : "",
                "Daily"  => $"Daily at {s.TimeOfDay}",
                "Weekly" => $"Weekly {System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetDayName((DayOfWeek)s.DayOfWeek)} {s.TimeOfDay}",
                _        => ""
            };

            // Append task registration indicator
            if (!string.IsNullOrEmpty(entry.WinTaskName))
                schedText += " · ✓ Win Task";
            return schedText;
        }

        private void AppendScheduleLog(ScriptEntry entry, string msg)
        {
            // Show log panel and route to script log if on script runner page
            if (PageScriptRunner.Visibility == Visibility.Visible)
            {
                ScriptLogPanel.Visibility = Visibility.Visible;
                AppendScriptLog($"[Schedule:{entry.Name}] {msg}", "#7DD3FC");
            }
        }

        private void RunScriptEntry(ScriptEntry entry)
        {
            // Load entry into runner and execute
            LoadLibraryEntry(entry);
            ScriptRun_Click(this, new RoutedEventArgs());
        }

        private void SaveLibrary()
        {
            AppDataManager.SaveScriptLibrary(_library);
        }

        private void RefreshLibraryUI()
        {
            LibraryPanel.Children.Clear();
            LibraryEmptyLabel.Visibility = _library.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

            foreach (var entry in _library.OrderByDescending(e => e.SavedAt))
            {
                var card = BuildLibraryCard(entry);
                LibraryPanel.Children.Add(card);
            }
        }

        private UIElement BuildLibraryCard(ScriptEntry entry)
        {
            string ext = System.IO.Path.GetExtension(entry.ScriptPath).ToLowerInvariant();
            string typeColor = ScriptTypes.TryGetValue(ext, out var info) ? info.Color : "#A8B3C8";
            string typeLabel = ScriptTypes.TryGetValue(ext, out var info2) ? info2.Label : "Script";

            var card = new Border
            {
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0D1020")),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640")),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Margin = new Thickness(2, 2, 2, 2),
                Padding = new Thickness(10, 8, 10, 8),
                Cursor = System.Windows.Input.Cursors.Hand,
                Tag = entry
            };

            var outer = new Grid();
            outer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            outer.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var textStack = new StackPanel();

            // Name + type badge
            var nameRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 3) };
            nameRow.Children.Add(new TextBlock
            {
                Text = entry.Name,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E8F0")),
                FontSize = 12.5, FontWeight = FontWeights.SemiBold,
                FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis, MaxWidth = 130
            });
            var badge = new Border
            {
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640")),
                CornerRadius = new CornerRadius(3), Padding = new Thickness(5, 1, 5, 1),
                Margin = new Thickness(6, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center
            };
            badge.Child = new TextBlock
            {
                Text = typeLabel, FontSize = 10, FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(typeColor)),
                FontFamily = new FontFamily("Segoe UI Variable, Segoe UI")
            };
            nameRow.Children.Add(badge);
            textStack.Children.Add(nameRow);

            // Description
            if (!string.IsNullOrEmpty(entry.Description))
                textStack.Children.Add(new TextBlock
                {
                    Text = entry.Description,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B7A99")),
                    FontSize = 11, TextTrimming = TextTrimming.CharacterEllipsis,
                    FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                    Margin = new Thickness(0, 0, 0, 3)
                });

            // File path
            textStack.Children.Add(new TextBlock
            {
                Text = System.IO.Path.GetFileName(entry.ScriptPath),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4A5F88")),
                FontSize = 10.5, TextTrimming = TextTrimming.CharacterEllipsis,
                FontFamily = new FontFamily("Consolas, Segoe UI Mono, Segoe UI"),
                ToolTip = entry.ScriptPath
            });

            // Schedule badge
            string nextRun = NextRunText(entry);
            if (!string.IsNullOrEmpty(nextRun))
            {
                var schedBadge = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0D2140")),
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E3A6E")),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(5, 2, 5, 2),
                    Margin = new Thickness(0, 4, 0, 0),
                    HorizontalAlignment = HorizontalAlignment.Left
                };
                var schedRow = new StackPanel { Orientation = Orientation.Horizontal };
                schedRow.Children.Add(new TextBlock { Text = "⏰ ", FontSize = 10 });
                schedRow.Children.Add(new TextBlock
                {
                    Text = nextRun, FontSize = 10,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#60A5FA")),
                    FontFamily = new FontFamily("Segoe UI Variable, Segoe UI")
                });
                if (entry.LastRanAt != null)
                    schedRow.Children.Add(new TextBlock
                    {
                        Text = $"  · Last ran {entry.LastRanAt.Value:HH:mm dd MMM}",
                        FontSize = 10,
                        Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4A5F88")),
                        FontFamily = new FontFamily("Segoe UI Variable, Segoe UI")
                    });
                schedBadge.Child = schedRow;
                textStack.Children.Add(schedBadge);

                // View Log button
                if (!string.IsNullOrEmpty(entry.ScheduledLogPath))
                {
                    var logBtn = new Button
                    {
                        Content = new TextBlock
                        {
                            Text = "View Log", FontSize = 9, FontWeight = FontWeights.SemiBold,
                            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#7DD3FC")),
                            FontFamily = new FontFamily("Segoe UI Variable, Segoe UI")
                        },
                        Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0A1A2E")),
                        BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E3A6E")),
                        BorderThickness = new Thickness(1),
                        Margin = new Thickness(0, 3, 0, 0),
                        Padding = new Thickness(5, 1, 5, 1),
                        Cursor = System.Windows.Input.Cursors.Hand,
                        HorizontalAlignment = HorizontalAlignment.Left,
                        ToolTip = entry.ScheduledLogPath,
                        Tag = entry.ScheduledLogPath
                    };
                    logBtn.Click += (s, ev) =>
                    {
                        ev.Handled = true;
                        string logPath = (string)((Button)s).Tag;
                        if (System.IO.File.Exists(logPath))
                            System.Diagnostics.Process.Start("notepad.exe", logPath);
                        else
                            DarkMessageBox.Show(
                                $"Log file not yet created.\nIt will appear after the first scheduled run:\n\n{logPath}",
                                "Log Not Found");
                    };
                    textStack.Children.Add(logBtn);
                }
            }

            Grid.SetColumn(textStack, 0);
            outer.Children.Add(textStack);

            // Right-side buttons (schedule + delete)
            var btnStack = new StackPanel { VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(6, 0, 0, 0), MinWidth = 46 };

            bool hasSchedule = entry.Schedule?.Enabled == true;
            var schedBtn = new Button
            {
                Width = 42, Height = 20,
                Content = new TextBlock
                {
                    Text = hasSchedule ? "⏱ On" : "⏱ Off",
                    FontSize = 9, FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                        hasSchedule ? "#60A5FA" : "#6B7A99")),
                    FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                    HorizontalAlignment = HorizontalAlignment.Center
                },
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                    hasSchedule ? "#0D2140" : "#161B2A")),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                    hasSchedule ? "#1E3A6E" : "#252D42")),
                BorderThickness = new Thickness(1),
                Margin = new Thickness(0, 0, 0, 3),
                Cursor = System.Windows.Input.Cursors.Hand,
                ToolTip = hasSchedule
                    ? $"Schedule: {NextRunText(entry)} (click to edit)"
                    : "No schedule set (click to add)",
                Tag = entry
            };
            schedBtn.Click += (s, ev) =>
            {
                ev.Handled = true;
                var e = (ScriptEntry)((Button)s).Tag;
                OpenScheduleDialog(e);
            };
            btnStack.Children.Add(schedBtn);

            // Delete button
            var delBtn = new Button
            {
                Width = 42, Height = 20,
                Content = new TextBlock
                {
                    Text = "Delete",
                    FontSize = 9, FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F87171")),
                    FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                    HorizontalAlignment = HorizontalAlignment.Center
                },
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2D0F0F")),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#5C1A1A")),
                BorderThickness = new Thickness(1),
                Cursor = System.Windows.Input.Cursors.Hand,
                ToolTip = "Remove from library",
                Tag = entry
            };
            delBtn.Click += (s, ev) =>
            {
                ev.Handled = true;
                var e = (ScriptEntry)((Button)s).Tag;
                bool confirmed = DarkMessageBox.Confirm(
                    $"Remove '{e.Name}' from the library?\n\nThis only removes it from the list — the script file itself is not deleted.",
                    "Remove Script");
                if (!confirmed) return;
                // Delete associated Windows Task and launcher .bat
                if (!string.IsNullOrEmpty(e.WinTaskName))
                    WindowsTaskScheduler.DeleteTask(e.WinTaskName);
                WindowsTaskScheduler.DeleteLauncher(e.ScriptPath);
                _library.Remove(e);
                SaveLibrary();
                RefreshLibraryUI();
            };
            btnStack.Children.Add(delBtn);

            Grid.SetColumn(btnStack, 1);
            outer.Children.Add(btnStack);

            card.Child = outer;

            // Click to load
            card.MouseLeftButtonUp += (s, _) => LoadLibraryEntry((ScriptEntry)((Border)s).Tag);

            // Hover effect
            card.MouseEnter += (s, _) =>
                ((Border)s).BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB"));
            card.MouseLeave += (s, _) =>
                ((Border)s).BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640"));

            return card;
        }

        private void OpenScheduleDialog(ScriptEntry entry)
        {
            var dlg = new ScheduleDialog(entry.Schedule) { Owner = this };
            if (dlg.ShowDialog() != true) return;

            // Remove old task and launcher if one existed
            if (!string.IsNullOrEmpty(entry.WinTaskName))
            {
                WindowsTaskScheduler.DeleteTask(entry.WinTaskName);
                entry.WinTaskName = null;
            }
            WindowsTaskScheduler.DeleteLauncher(entry.ScriptPath);

            entry.Schedule = dlg.Result;

            // Register new task if schedule is set and enabled
            if (entry.Schedule?.Enabled == true)
            {
                var (ok, taskName, err, logFile) = WindowsTaskScheduler.CreateTask(entry);
                if (ok)
                {
                    entry.WinTaskName      = taskName;
                    entry.ScheduledLogPath = logFile;
                    ScriptStatusLabel.Text = $"Schedule saved — Windows Task '{taskName}' created. Log: {logFile}";
                    ScriptStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
                }
                else
                {
                    DarkMessageBox.Show(
                        $"Schedule saved in library but Windows Task Scheduler registration failed:\n\n{err}\n\nThe script will still run while the app is open.",
                        "Task Scheduler Warning");
                }
            }
            else if (entry.Schedule == null)
            {
                ScriptStatusLabel.Text = "Schedule removed.";
                ScriptStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0xA8, 0xB3, 0xC8));
            }

            SaveLibrary();
            RefreshLibraryUI();
        }

        private bool _suppressScriptDirtyTracking = false;
        private ScriptEntry? _activeScriptEntry   = null;
        private bool         _scriptEntryDirty    = false;

        private void LoadLibraryEntry(ScriptEntry entry)
        {
            _suppressScriptDirtyTracking = true;
            _activeScriptEntry = entry;
            _scriptEntryDirty  = false;

            SetScriptFile(entry.ScriptPath);
            ScriptRuntimeBox.Text  = entry.Runtime;
            ScriptArgsBox.Text     = entry.Arguments;
            ScriptWorkDirBox.Text  = entry.WorkingDir;

            // Rebuild env vars
            ScriptEnvVarPanel.Children.Clear();
            foreach (var kv in entry.EnvVars)
            {
                ScriptAddEnvVar_Click(this, new RoutedEventArgs());
                var row = (Grid)ScriptEnvVarPanel.Children[^1];
                ((TextBox)row.Children[0]).Text = kv.Key;
                ((TextBox)row.Children[1]).Text = kv.Value;
            }

            ScriptStatusLabel.Text = $"Loaded: {entry.Name}";
            ScriptStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x60, 0xA5, 0xFA));
            ScriptUpdateLibraryBtn.Visibility = Visibility.Collapsed;

            _suppressScriptDirtyTracking = false;
        }

        private void MarkScriptEntryDirty()
        {
            if (_suppressScriptDirtyTracking) return;
            if (_activeScriptEntry == null) return;
            if (_scriptEntryDirty) return;
            _scriptEntryDirty = true;
            ScriptUpdateLibraryBtn.Visibility = Visibility.Visible;
            ScriptStatusLabel.Text = $"Unsaved changes — {_activeScriptEntry.Name}";
            ScriptStatusLabel.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FBBF24"));
        }

        private void ScriptUpdateLibrary_Click(object sender, RoutedEventArgs e)
        {
            if (_activeScriptEntry == null) return;

            _activeScriptEntry.ScriptPath  = _scriptFilePath ?? _activeScriptEntry.ScriptPath;
            _activeScriptEntry.Runtime     = ScriptRuntimeBox.Text.Trim();
            _activeScriptEntry.Arguments   = ScriptArgsBox.Text.Trim();
            _activeScriptEntry.WorkingDir  = ScriptWorkDirBox.Text.Trim();

            var envVars = new Dictionary<string, string>();
            foreach (var child in ScriptEnvVarPanel.Children)
            {
                if (child is Grid row && row.Children.Count >= 2)
                {
                    var k = (row.Children[0] as TextBox)?.Text?.Trim() ?? "";
                    var v = (row.Children[1] as TextBox)?.Text?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(k)) envVars[k] = v;
                }
            }
            _activeScriptEntry.EnvVars = envVars;

            SaveLibrary();
            RefreshLibraryUI();

            _scriptEntryDirty = false;
            ScriptUpdateLibraryBtn.Visibility = Visibility.Collapsed;
            ScriptStatusLabel.Text = $"Updated: {_activeScriptEntry.Name}";
            ScriptStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
        }

        private void LibrarySave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_scriptFilePath))
            {
                DarkMessageBox.Show("Select a script file first before saving to the library.",
                    "No Script");
                return;
            }

            // Show save dialog
            var dlg = new LibrarySaveDialog
            {
                Owner = this,
                SuggestedName = System.IO.Path.GetFileNameWithoutExtension(_scriptFilePath)
            };
            if (dlg.ShowDialog() != true) return;

            // Collect current env vars
            var envVars = new Dictionary<string, string>();
            foreach (var child in ScriptEnvVarPanel.Children)
            {
                if (child is Grid row && row.Children.Count >= 2)
                {
                    var k = (row.Children[0] as TextBox)?.Text?.Trim() ?? "";
                    var v = (row.Children[1] as TextBox)?.Text?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(k)) envVars[k] = v;
                }
            }

            var entry = new ScriptEntry
            {
                Name        = dlg.EntryName,
                Description = dlg.EntryDescription,
                ScriptPath  = _scriptFilePath!,
                Runtime     = ScriptRuntimeBox.Text.Trim(),
                Arguments   = ScriptArgsBox.Text.Trim(),
                WorkingDir  = ScriptWorkDirBox.Text.Trim(),
                EnvVars     = envVars,
                SavedAt     = DateTime.Now
            };

            _library.Add(entry);
            SaveLibrary();
            RefreshLibraryUI();

            ScriptStatusLabel.Text = $"Saved '{entry.Name}' to library.";
            ScriptStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
        }


        private string? _scriptFilePath = null;
        private System.Diagnostics.Process? _scriptProcess = null;
        private string? _saveLogPath = null;

        private void SaveLog_Toggled(object sender, RoutedEventArgs e)
        {
            bool on = SaveLogCheckbox.IsChecked == true;
            SaveLogPathPanel.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
            if (!on) { _saveLogPath = null; return; }
            if (string.IsNullOrEmpty(_scriptFilePath)) return;
            // Default path
            _saveLogPath = System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(_scriptFilePath) ?? "",
                System.IO.Path.GetFileNameWithoutExtension(_scriptFilePath) + "_run.log");
            SaveLogPathLabel.Text = _saveLogPath;
            SaveLogPathLabel.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#CBD5E1"));
        }

        private void SaveLogBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save log to...",
                Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*",
                DefaultExt = ".log",
                FileName = System.IO.Path.GetFileNameWithoutExtension(_scriptFilePath ?? "script") + "_run.log",
                OverwritePrompt = false
            };
            if (dlg.ShowDialog() == true)
            {
                _saveLogPath = dlg.FileName;
                SaveLogPathLabel.Text = _saveLogPath;
                SaveLogPathLabel.Foreground = new SolidColorBrush(
                    (Color)ColorConverter.ConvertFromString("#CBD5E1"));
            }
        }

        private void SaveLogAfterRun_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save log as...",
                Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*",
                DefaultExt = ".log",
                FileName = System.IO.Path.GetFileNameWithoutExtension(_scriptFilePath ?? "script") + "_run.log",
                OverwritePrompt = true,
                InitialDirectory = string.IsNullOrEmpty(_scriptFilePath) ? ""
                    : System.IO.Path.GetDirectoryName(_scriptFilePath) ?? ""
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                string header = $"========================================{Environment.NewLine}" +
                                $"Run: {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}" +
                                $"Script: {_scriptFilePath}{Environment.NewLine}" +
                                $"========================================{Environment.NewLine}";
                // Strip timestamps from displayed log for clean save
                System.IO.File.WriteAllText(dlg.FileName, header + ScriptLog.Text + Environment.NewLine);
                SaveLogAfterRunBtn.Visibility = Visibility.Collapsed;
                ScriptStatusLabel.Text = $"Log saved to {System.IO.Path.GetFileName(dlg.FileName)}";
                ScriptStatusLabel.Foreground = new SolidColorBrush(
                    (Color)ColorConverter.ConvertFromString("#4ADE80"));
            }
            catch (Exception ex)
            {
                DarkMessageBox.Show($"Could not save log:\n{ex.Message}", "Save Failed");
            }
        }
        private static readonly Dictionary<string, (string Runtime, string ArgsPrefix, string Label, string Color)> ScriptTypes = new()
        {
            { ".ps1",  ("powershell.exe", "-ExecutionPolicy Bypass -File",  "PowerShell", "#5B9CF6") },
            { ".py",   ("python",         "",                                "Python",     "#F6C94B") },
            { ".jar",  ("java",           "-jar",                            "Java",       "#F6804B") },
            { ".bat",  ("cmd.exe",        "/c",                              "Batch",      "#A8B3C8") },
            { ".cmd",  ("cmd.exe",        "/c",                              "Batch",      "#A8B3C8") },
            { ".js",   ("node",           "",                                "Node.js",    "#6BCB77") },
            { ".sh",   ("bash",           "",                                "Shell",      "#E07BE0") },
        };

        private void ScriptBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Select Script File",
                Filter = "All Scripts (*.ps1;*.py;*.jar;*.bat;*.cmd;*.js;*.sh)|*.ps1;*.py;*.jar;*.bat;*.cmd;*.js;*.sh" +
                         "|PowerShell (*.ps1)|*.ps1|Python (*.py)|*.py|Java (*.jar)|*.jar" +
                         "|Batch (*.bat;*.cmd)|*.bat;*.cmd|Node.js (*.js)|*.js|Shell (*.sh)|*.sh|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
                SetScriptFile(dlg.FileName);
        }


        private void ScriptFile_Drop(object sender, DragEventArgs e)
        {
            if (sender is Border b)
                b.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252D42"));
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files?.Length > 0) SetScriptFile(files[0]);
            }
        }

        private void ScriptFile_DragEnter(object sender, DragEventArgs e)
        {
            if (sender is Border b && e.Data.GetDataPresent(DataFormats.FileDrop))
                b.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB"));
        }

        private void ScriptFile_DragLeave(object sender, DragEventArgs e)
        {
            if (sender is Border b)
                b.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252D42"));
        }

        private void ScriptField_Changed(object sender, System.Windows.Controls.TextChangedEventArgs e)
            => MarkScriptEntryDirty();

        private void TrendsField_Changed(object sender, System.Windows.Controls.TextChangedEventArgs e)
            => MarkTrendsCustomerDirty();

        private void ScriptWorkDir_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Select working directory",
                Filter = "Folder|*.folder",
                FileName = "Select Folder",
                CheckFileExists = false,
                CheckPathExists = true,
                ValidateNames = false
            };
            if (dlg.ShowDialog() == true)
            {
                ScriptWorkDirBox.Text = System.IO.Path.GetDirectoryName(dlg.FileName) ?? string.Empty;
                MarkScriptEntryDirty();
            }
        }

        private void ScriptAddEnvVar_Click(object sender, RoutedEventArgs e)
        {
            var row = new Grid { Margin = new Thickness(0, 2, 0, 2) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(8) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var keyBox = new TextBox
            {
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0F1117")),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CBD5E1")),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2A2F3E")),
                FontSize = 11, Height = 26, Padding = new Thickness(6, 0, 6, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                FontFamily = new FontFamily("Consolas, Segoe UI Mono"),
                ToolTip = "Variable name, e.g. MY_ENV"
            };
            var valBox = new TextBox
            {
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0F1117")),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CBD5E1")),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2A2F3E")),
                FontSize = 11, Height = 26, Padding = new Thickness(6, 0, 6, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                FontFamily = new FontFamily("Consolas, Segoe UI Mono"),
                ToolTip = "Variable value"
            };
            var removeBtn = new Button
            {
                Content = "✕",
                Style = (Style)Resources["RemoveButtonStyle"],
                Margin = new Thickness(4, 0, 0, 0)
            };
            removeBtn.Click += (s, _) => ScriptEnvVarPanel.Children.Remove(row);

            Grid.SetColumn(keyBox,    0);
            Grid.SetColumn(valBox,    2);
            Grid.SetColumn(removeBtn, 3);
            row.Children.Add(keyBox);
            row.Children.Add(valBox);
            row.Children.Add(removeBtn);
            ScriptEnvVarPanel.Children.Add(row);
        }

        private void ScriptStop_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _scriptProcess?.Kill(entireProcessTree: true);
                AppendScriptLog("\n■ Process killed by user.", "#F87171");
            }
            catch { }
        }

        private void AppendScriptLog(string text, string colorHex = "#A8B3C8")
        {
            string ts   = DateTime.Now.ToString("HH:mm:ss");
            string line = $"[{ts}]  {text}";
            if (ScriptLog.Text.Length > 0) ScriptLog.Text += "\n";
            ScriptLog.Text += line;
            ScriptLog.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));
            ScriptLogScroller.ScrollToEnd();

            // Also append to log file if enabled
            if (SaveLogCheckbox.IsChecked == true && !string.IsNullOrEmpty(_saveLogPath))
            {
                try { System.IO.File.AppendAllText(_saveLogPath, line + Environment.NewLine); }
                catch { }
            }
        }

        private void WriteLogHeader()
        {
            if (SaveLogCheckbox.IsChecked != true || string.IsNullOrEmpty(_saveLogPath)) return;
            try
            {
                string header = $"========================================{Environment.NewLine}" +
                                $"Run: {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}" +
                                $"Script: {_scriptFilePath}{Environment.NewLine}" +
                                $"========================================{Environment.NewLine}";
                System.IO.File.AppendAllText(_saveLogPath, header);
            }
            catch { }
        }

        // ── Test Run Trends — DB ─────────────────────────────────────────────────

        private class DbTrendsCustomer
        {
            public string Id                { get; set; } = Guid.NewGuid().ToString("N")[..8];
            public string Name              { get; set; } = "";
            public string ApiHost           { get; set; } = "";
            public string ApiKey            { get; set; } = "";
            public string DownloadFolder    { get; set; } = "";
            public string ReportsFolder     { get; set; } = "";
            public int    FailWindow        { get; set; } = 3;
            public bool   WatchEnabled      { get; set; } = false;
            public int    WatchIntervalSecs { get; set; } = 300;
            public DateTime? LastGenerated  { get; set; } = null;
            public List<string> KnownRuns      { get; set; } = new();
            public int          MaxMonths      { get; set; } = 0;
            public bool         IncludeOldRuns { get; set; } = false;
        }

        private List<DbTrendsCustomer>  _dbTrendsLibrary        = new();
        private DbTrendsCustomer?       _activeDbTrendsCustomer = null;
        private List<string>            _dbFetchedRunNumbers    = new();
        private List<string>            _dbApiHosts             = new();
        private readonly Dictionary<string, System.Windows.Threading.DispatcherTimer> _dbWatchTimers = new();
        private bool                    _dbTrendsCustomerDirty  = false;
        private bool                    _suppressDbDirtyTracking = false;

        private void MarkDbTrendsCustomerDirty()
        {
            if (_suppressDbDirtyTracking) return;
            if (_activeDbTrendsCustomer == null) return;
            if (_dbTrendsCustomerDirty) return;
            _dbTrendsCustomerDirty = true;
            DbTrendsUpdateLibraryBtn.Visibility = Visibility.Visible;
            DbTrendsStatusLabel.Text = $"Unsaved changes — {_activeDbTrendsCustomer.Name}";
            DbTrendsStatusLabel.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FBBF24"));
        }

        private void LoadDbApiHostsToComboBox()
        {
            _dbApiHosts = AppDataManager.LoadDbApiHosts();
            DbTrendsApiHostBox.Items.Clear();
            foreach (var h in _dbApiHosts)
                DbTrendsApiHostBox.Items.Add(h);
            // Select the default host
            DbTrendsApiHostBox.Text = _dbApiHosts.Count > 0 ? _dbApiHosts[0] : "http://apso1wats4:8080";
            UpdateDbTrendsFetchUrlPreview();
        }

                private void RefreshDbTrendsLibraryUI()
        {
            DbTrendsLibraryPanel.Children.Clear();
            DbTrendsLibraryEmptyLabel.Visibility = _dbTrendsLibrary.Count == 0
                ? Visibility.Visible : Visibility.Collapsed;
            foreach (var c in _dbTrendsLibrary)
                DbTrendsLibraryPanel.Children.Add(BuildDbTrendsCard(c));
        }

        private UIElement BuildDbTrendsCard(DbTrendsCustomer customer)
        {
            var card = new Border
            {
                Background      = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0D1020")),
                BorderBrush     = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1A2030")),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(6),
                Padding         = new Thickness(10, 8, 10, 8),
                Margin          = new Thickness(2, 2, 2, 2),
                Cursor          = System.Windows.Input.Cursors.Hand,
                Tag             = customer
            };

            var nameBlock = new TextBlock
            {
                Text         = customer.Name,
                Foreground   = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CBD5E1")),
                FontSize     = 12.5,
                FontWeight   = FontWeights.SemiBold,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            var lastGen = customer.LastGenerated.HasValue
                ? $"Last: {customer.LastGenerated.Value:MMM d, yyyy}"
                : "Never generated";
            var subBlock = new TextBlock
            {
                Text       = lastGen,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4A5F88")),
                FontSize   = 10.5,
                Margin     = new Thickness(0, 2, 0, 0)
            };

            var deleteBtn = new Button
            {
                Content    = "\u2715",
                Style      = (Style)Resources["ActionButtonStyle"],
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640")),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B7A99")),
                FontSize   = 10,
                Width = 20, Height = 20,
                Tag = customer
            };
            deleteBtn.Click += (s, _) =>
            {
                _dbTrendsLibrary.Remove((DbTrendsCustomer)((Button)s).Tag);
                RefreshDbTrendsLibraryUI();
            };

            var headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(nameBlock, 0);
            Grid.SetColumn(deleteBtn, 1);
            headerGrid.Children.Add(nameBlock);
            headerGrid.Children.Add(deleteBtn);

            var sp = new StackPanel();
            sp.Children.Add(headerGrid);
            sp.Children.Add(subBlock);
            card.Child = sp;

            card.MouseLeftButtonUp += (s, _) => LoadDbTrendsCustomer((DbTrendsCustomer)((Border)s).Tag);
            return card;
        }

        private void LoadDbTrendsCustomer(DbTrendsCustomer c)
        {
            _suppressDbDirtyTracking = true;
            _activeDbTrendsCustomer = c;
            _dbTrendsCustomerDirty  = false;
            DbTrendsCustomerNameBox.Text = c.Name;
            DbTrendsFailWindowBox.Text   = c.FailWindow > 0 ? c.FailWindow.ToString() : "3";
            DbTrendsWatchIntervalBox.Text = c.WatchIntervalSecs > 0 ? c.WatchIntervalSecs.ToString() : "300";
            DbTrendsApiHostBox.Text         = string.IsNullOrEmpty(c.ApiHost) ? "http://apso1wats4:8080" : c.ApiHost;
            DbTrendsMaxMonthsBox.Text       = c.MaxMonths > 0 ? c.MaxMonths.ToString() : "0";
            DbTrendsIncludeOldRunsCheck.IsChecked = c.IncludeOldRuns;
            UpdateDbTrendsFetchUrlPreview();
            DbTrendsApiKeyBox.Password   = c.ApiKey;
            DbTrendsDownloadFolderLabel.Text = string.IsNullOrEmpty(c.DownloadFolder) ? "Not set" : c.DownloadFolder;
            DbTrendsDownloadFolderLabel.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                string.IsNullOrEmpty(c.DownloadFolder) ? "#4A5F88" : "#CBD5E1"));
            DbTrendsReportsFolderLabel.Text = string.IsNullOrEmpty(c.ReportsFolder) ? "Same as Download folder" : c.ReportsFolder;
            DbTrendsUpdateLibraryBtn.Visibility = Visibility.Collapsed;
            DbTrendsStatusLabel.Text = "";
            DbTrendsDownloadBtn.IsEnabled  = false;
            DbTrendsRunBtn.IsEnabled       = false;
            _dbFetchedRunNumbers.Clear();
            _suppressDbDirtyTracking = false;
        }

        private void DbTrendsField_Changed(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateDbTrendsFetchUrlPreview();
            MarkDbTrendsCustomerDirty();
        }

        private void DbTrendsApiHost_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateDbTrendsFetchUrlPreview();
            MarkDbTrendsCustomerDirty();
        }

        private void DbTrendsApiHostSave_Click(object sender, RoutedEventArgs e)
        {
            var host = DbTrendsApiHostBox.Text?.Trim().TrimEnd('/');
            if (string.IsNullOrEmpty(host) || host == "http://" || host == "https://")
            {
                DbTrendsStatusLabel.Text = "\u26a0 Enter a valid host before saving.";
                return;
            }
            var hosts = _dbApiHosts;
            if (!hosts.Contains(host, StringComparer.OrdinalIgnoreCase))
            {
                hosts.Add(host);
                AppDataManager.SaveDbApiHosts(hosts);
                DbTrendsApiHostBox.Items.Add(host);
                DbTrendsApiHostBox.Text = host;
                DbTrendsStatusLabel.Text = $"\u2714 Host saved: {host}";
            }
            else
            {
                DbTrendsStatusLabel.Text = $"Host already in list.";
            }
        }

        private void DbTrendsApiKey_Changed(object sender, RoutedEventArgs e) { MarkDbTrendsCustomerDirty(); }

        private void DbTrendsIncludeOldRuns_Changed(object sender, RoutedEventArgs e) { MarkDbTrendsCustomerDirty(); }

        private void UpdateDbTrendsFetchUrlPreview()
        {
            if (DbTrendsApiHostBox == null || DbTrendsFetchUrlPreview == null) return;
            var host     = DbTrendsApiHostBox.Text.TrimEnd('/');
            var customer = DbTrendsCustomerNameBox?.Text.Trim() ?? "";
            if (string.IsNullOrEmpty(host) || host == "http://" || host == "https://")
            {
                DbTrendsFetchUrlPreview.Text          = "—";
                DbTrendsFetchUrlPreview.Foreground    = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4A5F88"));
                DbTrendsDownloadUrlPreview.Text       = "—";
                DbTrendsDownloadUrlPreview.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4A5F88"));
                return;
            }
            var encodedCustomer = Uri.EscapeDataString(customer);
            DbTrendsFetchUrlPreview.Text          = $"{host}/spring/runNumber/getRunNumberByProjectName?projectName={encodedCustomer}";
            DbTrendsFetchUrlPreview.Foreground    = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#60A5FA"));
            DbTrendsDownloadUrlPreview.Text       = $"{host}/spring/export-to-excel?projectName={encodedCustomer}&runNumber={{run}}&result=final";
            DbTrendsDownloadUrlPreview.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A78BFA"));
        }

        private string BuildFetchUrl()
        {
            var host     = DbTrendsApiHostBox.Text.TrimEnd('/');
            var customer = DbTrendsCustomerNameBox.Text.Trim();
            return $"{host}/spring/runNumber/getRunNumberByProjectName?projectName={Uri.EscapeDataString(customer)}";
        }

        private string BuildDownloadUrl(string runNumber)
        {
            var host     = DbTrendsApiHostBox.Text.TrimEnd('/');
            var customer = DbTrendsCustomerNameBox.Text.Trim();
            return $"{host}/spring/export-to-excel?projectName={Uri.EscapeDataString(customer)}&runNumber={Uri.EscapeDataString(runNumber)}&result=final";
        }

        private void DbTrendsLibraryAdd_Click(object sender, RoutedEventArgs e)
        {
            var c = new DbTrendsCustomer { Name = "New Customer" };
            _dbTrendsLibrary.Add(c);
            RefreshDbTrendsLibraryUI();
            LoadDbTrendsCustomer(c);
        }

        private void DbTrendsGenerateAll_Click(object sender, RoutedEventArgs e)
        {
            var eligible = _dbTrendsLibrary
                .Where(c => !string.IsNullOrEmpty(c.DownloadFolder)
                         && System.IO.Directory.Exists(c.DownloadFolder))
                .ToList();

            if (eligible.Count == 0)
            {
                DarkMessageBox.Show(
                    "No DB customers are ready to generate.\n\nMake sure each customer has a valid Download folder set.",
                    "Generate All");
                return;
            }

            // Run each customer's generation directly from its saved data — do NOT call
            // LoadDbTrendsCustomer() here, which would reset UI state and race with the
            // background tasks that read from those same UI fields.
            foreach (var c in eligible)
                RunDbTrendsGenerationForCustomer(c);
        }

        private void DbTrendsUpdateLibrary_Click(object sender, RoutedEventArgs e)
        {
            if (_activeDbTrendsCustomer == null) return;
            _activeDbTrendsCustomer.Name           = DbTrendsCustomerNameBox.Text.Trim();
            _activeDbTrendsCustomer.ApiHost        = DbTrendsApiHostBox.Text.Trim();
            _activeDbTrendsCustomer.ApiKey         = DbTrendsApiKeyBox.Password;
            _activeDbTrendsCustomer.DownloadFolder = DbTrendsDownloadFolderLabel.Text == "Not set" ? "" : DbTrendsDownloadFolderLabel.Text;
            _activeDbTrendsCustomer.ReportsFolder  = DbTrendsReportsFolderLabel.Text == "Same as Download folder" ? "" : DbTrendsReportsFolderLabel.Text;
            if (int.TryParse(DbTrendsFailWindowBox.Text, out int fw)) _activeDbTrendsCustomer.FailWindow = fw;
            _activeDbTrendsCustomer.MaxMonths      = int.TryParse(DbTrendsMaxMonthsBox.Text.Trim(), out int dmm) && dmm > 0 ? dmm : 0;
            _activeDbTrendsCustomer.IncludeOldRuns = DbTrendsIncludeOldRunsCheck.IsChecked == true;
            _activeDbTrendsCustomer.WatchIntervalSecs = int.TryParse(DbTrendsWatchIntervalBox.Text.Trim(), out int dwi) && dwi > 0 ? dwi : 300;
            RefreshDbTrendsLibraryUI();
            _dbTrendsCustomerDirty = false;
            DbTrendsUpdateLibraryBtn.Visibility = Visibility.Collapsed;
            DbTrendsStatusLabel.Text = $"\u2714 '{_activeDbTrendsCustomer.Name}' saved to library.";
            DbTrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
        }

        private void DbTrendsDownloadFolderBrowse_Click(object sender, RoutedEventArgs e)
        {
            string f = BrowseFolder("Select Download folder (run files will be saved here)");
            if (string.IsNullOrEmpty(f)) return;
            DbTrendsDownloadFolderLabel.Text       = f;
            DbTrendsDownloadFolderLabel.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CBD5E1"));
            MarkDbTrendsCustomerDirty();
        }

        private void DbTrendsReportsFolderBrowse_Click(object sender, RoutedEventArgs e)
        {
            string f = BrowseFolder("Select Reports folder (where Trends.xlsx will be saved)");
            if (string.IsNullOrEmpty(f)) return;
            DbTrendsReportsFolderLabel.Text = f;
            MarkDbTrendsCustomerDirty();
        }

        private async void DbTrendsFetchRuns_Click(object sender, RoutedEventArgs e)
        {
            var host     = DbTrendsApiHostBox.Text.TrimEnd('/');
            var customer = DbTrendsCustomerNameBox.Text.Trim();

            if (string.IsNullOrEmpty(host) || host == "http://" || host == "https://")
            {
                DbTrendsStatusLabel.Text = "\u26a0 Please enter an API Host first.";
                return;
            }
            if (string.IsNullOrEmpty(customer))
            {
                DbTrendsStatusLabel.Text = "\u26a0 Please enter a Customer Name first.";
                return;
            }

            var url    = BuildFetchUrl();
            var apiKey = DbTrendsApiKeyBox.Password;

            DbTrendsFetchRunsBtn.IsEnabled = false;
            DbTrendsLogPanel.Visibility    = Visibility.Visible;
            DbTrendsProgress.Visibility    = Visibility.Visible;
            DbTrendsLog.Text               = $"Fetching run numbers from:\n  {url}\n";
            DbTrendsStatusLabel.Text       = "";

            try
            {
                using var http = new System.Net.Http.HttpClient();
                if (!string.IsNullOrEmpty(apiKey))
                    http.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

                var response = await http.GetAsync(url);
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                // Response is a JSON array of "RunNumber,ID" strings e.g. ["TST_MAR_26_1,20210", ...]
                // Strip the trailing ",<number>" — keep only the run number before the comma
                var raw = System.Text.Json.JsonSerializer.Deserialize<List<string>>(json) ?? new();
                _dbFetchedRunNumbers = raw
                    .Select(s => s.Contains(',') ? s[..s.LastIndexOf(',')] : s)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();

                DbTrendsLog.Text += $"\u2714 {_dbFetchedRunNumbers.Count} runs found:\n";
                foreach (var r in _dbFetchedRunNumbers)
                    DbTrendsLog.Text += $"  {r}\n";

                DbTrendsRunsFoundLabel.Text   = $"{_dbFetchedRunNumbers.Count} runs retrieved from API";
                DbTrendsDownloadBtn.IsEnabled = true;
                DbTrendsStatusLabel.Text      = $"\u2714 {_dbFetchedRunNumbers.Count} runs fetched. Click Download Files to proceed.";
            }
            catch (Exception ex)
            {
                DbTrendsLog.Text        += $"\u2716 Error: {ex.Message}\n";
                DbTrendsStatusLabel.Text = $"\u2716 Fetch failed: {ex.Message}";
            }
            finally
            {
                DbTrendsFetchRunsBtn.IsEnabled = true;
                DbTrendsProgress.Visibility    = Visibility.Collapsed;
            }
        }

        private async void DbTrendsDownload_Click(object sender, RoutedEventArgs e)
        {
            var downloadFolder = DbTrendsDownloadFolderLabel.Text;
            if (string.IsNullOrEmpty(downloadFolder) || downloadFolder == "Not set")
            {
                DbTrendsStatusLabel.Text = "\u26a0 Please set a Download Folder first.";
                return;
            }
            if (_dbFetchedRunNumbers.Count == 0)
            {
                DbTrendsStatusLabel.Text = "\u26a0 No run numbers available. Fetch runs first.";
                return;
            }

            var apiKey = DbTrendsApiKeyBox.Password;
            DbTrendsDownloadBtn.IsEnabled = false;
            DbTrendsLogPanel.Visibility   = Visibility.Visible;
            DbTrendsProgress.Visibility   = Visibility.Visible;

            // Determine which runs are missing from disk before downloading
            var customerNameForDl = DbTrendsCustomerNameBox.Text.Trim();
            var dlPrefix = customerNameForDl + "_";
            var onDisk = System.IO.Directory.Exists(downloadFolder)
                ? System.IO.Directory.GetFiles(downloadFolder, "*.xlsx")
                    .Select(f => System.IO.Path.GetFileNameWithoutExtension(f))
                    .Where(n => n.StartsWith(dlPrefix, StringComparison.OrdinalIgnoreCase))
                    .Select(n => n[dlPrefix.Length..])
                    .ToHashSet(StringComparer.OrdinalIgnoreCase)
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var runsToDownload = _dbFetchedRunNumbers.Where(r => !onDisk.Contains(r)).ToList();
            var skippedCount   = _dbFetchedRunNumbers.Count - runsToDownload.Count;

            DbTrendsLog.Text = skippedCount > 0
                ? $"Downloading {runsToDownload.Count} new file(s) to {downloadFolder}... ({skippedCount} already on disk, skipping)\n"
                : $"Downloading {runsToDownload.Count} file(s) to {downloadFolder}...\n";

            if (runsToDownload.Count == 0)
            {
                DbTrendsLog.Text        += $"\u2714 All {skippedCount} file(s) already present — nothing to download.\n";
                DbTrendsRunBtn.IsEnabled     = true;
                DbTrendsDownloadBtn.IsEnabled = true;
                DbTrendsProgress.Visibility  = Visibility.Collapsed;
                DbTrendsStatusLabel.Text = $"\u2714 All files already on disk. Click Generate Trends to produce the report.";
                return;
            }

            int success = 0, failed = 0;
            try
            {
                using var http = new System.Net.Http.HttpClient();
                if (!string.IsNullOrEmpty(apiKey))
                    http.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
                http.Timeout = TimeSpan.FromMinutes(5);

                foreach (var run in runsToDownload)
                {
                    var dlUrl    = BuildDownloadUrl(run);
                    var fileName = $"{customerNameForDl}_{run}.xlsx";
                    var target   = System.IO.Path.Combine(downloadFolder, fileName);
                    try
                    {
                        var bytes = await http.GetByteArrayAsync(dlUrl);
                        await System.IO.File.WriteAllBytesAsync(target, bytes);
                        DbTrendsLog.Text += $"  \u2193 {fileName}\n";
                        success++;
                    }
                    catch (Exception exInner)
                    {
                        DbTrendsLog.Text += $"  \u2716 {run} failed: {exInner.Message}\n";
                        failed++;
                    }
                }

                if (failed == 0)
                {
                    DbTrendsLog.Text        += $"\u2714 All {success} files downloaded.\n";
                    DbTrendsRunBtn.IsEnabled = true;
                    DbTrendsStatusLabel.Text = $"\u2714 {success} files downloaded. Click Generate Trends to produce the report.";
                }
                else
                {
                    DbTrendsLog.Text        += $"\u26a0 {success} downloaded, {failed} failed.\n";
                    DbTrendsRunBtn.IsEnabled = success > 0;
                    DbTrendsStatusLabel.Text = $"\u26a0 {success} succeeded, {failed} failed. Check log for details.";
                    DbTrendsDownloadBtn.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                DbTrendsLog.Text        += $"\u2716 Error: {ex.Message}\n";
                DbTrendsStatusLabel.Text = $"\u2716 Download failed: {ex.Message}";
                DbTrendsDownloadBtn.IsEnabled = true;
            }
            finally
            {
                DbTrendsProgress.Visibility = Visibility.Collapsed;
            }
        }

        private void DbTrendsRun_Click(object sender, RoutedEventArgs e)
            => RunDbTrendsGeneration();

        private void RunDbTrendsGeneration()
        {
            var name           = DbTrendsCustomerNameBox.Text.Trim();
            var downloadFolder = DbTrendsDownloadFolderLabel.Text;
            var reportsFolder  = DbTrendsReportsFolderLabel.Text == "Same as Download folder" || string.IsNullOrEmpty(DbTrendsReportsFolderLabel.Text)
                                    ? downloadFolder : DbTrendsReportsFolderLabel.Text;

            if (string.IsNullOrEmpty(name))
            { DarkMessageBox.Show("Enter a customer name.", "Required"); return; }

            if (string.IsNullOrEmpty(downloadFolder) || downloadFolder == "Not set" ||
                !System.IO.Directory.Exists(downloadFolder))
            { DarkMessageBox.Show("Select a valid Download folder.", "Required"); return; }

            // Use saved customer value when available; fall back to the text box
            int failWindow = (_activeDbTrendsCustomer?.FailWindow > 0)
                ? _activeDbTrendsCustomer.FailWindow
                : (int.TryParse(DbTrendsFailWindowBox.Text.Trim(), out int fwBox) && fwBox > 0 ? fwBox : 3);

            DbTrendsLogPanel.Visibility = Visibility.Visible;
            DbTrendsProgress.Visibility = Visibility.Visible;
            DbTrendsLog.Text            = "";
            DbTrendsRunBtn.IsEnabled    = false;
            DbTrendsStatusLabel.Text    = "Generating…";
            DbTrendsStatusLabel.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#60A5FA"));

            var customerName   = name;
            var runsFolder     = downloadFolder;
            var rptFolder      = reportsFolder;
            var fw             = failWindow;
            int maxMonths      = int.TryParse(DbTrendsMaxMonthsBox.Text.Trim(), out int mm) && mm > 0 ? mm : 0;
            bool includeOldRuns = DbTrendsIncludeOldRunsCheck.IsChecked == true;

            System.Threading.Tasks.Task.Run(() =>
            {
                Action<string> log = msg =>
                    Dispatcher.Invoke(() =>
                    {
                        DbTrendsLog.Text += msg + "\n";
                        // Auto-scroll
                        var sv = DbTrendsLog.Parent as ScrollViewer
                              ?? (DbTrendsLog.Parent as FrameworkElement)?.Parent as ScrollViewer;
                        sv?.ScrollToEnd();
                    });

                var (ok, outputPath, error) =
                    TestRunTrendsProcessor.Generate(log, runsFolder, customerName, rptFolder, fw, maxMonths, includeOldRuns);

                Dispatcher.Invoke(() =>
                {
                    DbTrendsProgress.Visibility = Visibility.Collapsed;
                    DbTrendsRunBtn.IsEnabled    = true;

                    if (ok)
                    {
                        if (_activeDbTrendsCustomer != null)
                        {
                            _activeDbTrendsCustomer.LastGenerated = DateTime.Now;
                            RefreshDbTrendsLibraryUI();
                        }

                        string shortName = System.IO.Path.GetFileName(outputPath);
                        DbTrendsStatusLabel.Text = $"\u2714 Saved to {shortName}";
                        DbTrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));

                        if (DarkMessageBox.Confirm($"Done:\n{outputPath}\n\nOpen it now?", "Trends Generated"))
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                { FileName = outputPath, UseShellExecute = true });
                    }
                    else
                    {
                        DbTrendsStatusLabel.Text = $"\u2716 Failed: {error}";
                        DbTrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71));
                    }
                });
            });
        }

        // Runs generation for an arbitrary customer object without touching any UI fields.
        // Used by GenerateAll so multiple background tasks never race on the same controls.
        private void RunDbTrendsGenerationForCustomer(DbTrendsCustomer c)
        {
            if (string.IsNullOrEmpty(c.DownloadFolder) || !System.IO.Directory.Exists(c.DownloadFolder))
                return;

            var runsFolder  = c.DownloadFolder;
            var rptFolder   = string.IsNullOrEmpty(c.ReportsFolder) ? runsFolder : c.ReportsFolder;
            var fw          = c.FailWindow > 0 ? c.FailWindow : 3;
            var maxMonths   = c.MaxMonths;
            var inclOld     = c.IncludeOldRuns;
            var custName    = c.Name;
            var custId      = c.Id;

            System.Threading.Tasks.Task.Run(() =>
            {
                var (ok, outputPath, error) =
                    TestRunTrendsProcessor.Generate(null, runsFolder, custName, rptFolder, fw, maxMonths, inclOld);

                Dispatcher.Invoke(() =>
                {
                    var entry = _dbTrendsLibrary.FirstOrDefault(x => x.Id == custId);
                    if (entry != null) entry.LastGenerated = DateTime.Now;
                    SaveDbTrendsLibrary();
                    RefreshDbTrendsLibraryUI();

                    // If this customer is currently active in the UI, update status labels
                    if (_activeDbTrendsCustomer?.Id == custId)
                    {
                        DbTrendsProgress.Visibility = Visibility.Collapsed;
                        DbTrendsRunBtn.IsEnabled    = true;
                        if (ok)
                        {
                            DbTrendsStatusLabel.Text = $"\u2714 Saved to {System.IO.Path.GetFileName(outputPath)}";
                            DbTrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
                        }
                        else
                        {
                            DbTrendsStatusLabel.Text = $"\u2716 Failed: {error}";
                            DbTrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71));
                        }
                    }
                });
            });
        }

        // ── DB Library persistence ──────────────────────────────────────────────

        private void SaveDbTrendsLibrary()
        {
            try
            {
                var dtos = _dbTrendsLibrary.Select(c => new
                {
                    c.Id, c.Name, c.ApiHost, c.ApiKey, c.DownloadFolder, c.ReportsFolder,
                    c.FailWindow, c.WatchEnabled, c.WatchIntervalSecs, c.LastGenerated, c.KnownRuns,
                    c.MaxMonths, c.IncludeOldRuns
                });
                var json = System.Text.Json.JsonSerializer.Serialize(dtos,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                var path = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(AppDataManager.TrendsLibraryPath)!,
                    "db_trends_library.json");
                System.IO.File.WriteAllText(path, json);
            }
            catch { }
        }

        private void LoadDbTrendsLibrary()
        {
            try
            {
                var path = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(AppDataManager.TrendsLibraryPath)!,
                    "db_trends_library.json");
                if (!System.IO.File.Exists(path)) return;
                var json = System.IO.File.ReadAllText(path);
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                foreach (var el in doc.RootElement.EnumerateArray())
                {
                    var c = new DbTrendsCustomer
                    {
                        Id             = el.TryGetProperty("Id",             out var v) ? v.GetString() ?? Guid.NewGuid().ToString("N")[..8] : Guid.NewGuid().ToString("N")[..8],
                        Name           = el.TryGetProperty("Name",           out v) ? v.GetString() ?? "" : "",
                        ApiHost        = el.TryGetProperty("ApiHost",        out v) ? v.GetString() ?? "" : "",
                        ApiKey         = el.TryGetProperty("ApiKey",         out v) ? v.GetString() ?? "" : "",
                        DownloadFolder = el.TryGetProperty("DownloadFolder", out v) ? v.GetString() ?? "" : "",
                        ReportsFolder  = el.TryGetProperty("ReportsFolder",  out v) ? v.GetString() ?? "" : "",
                        FailWindow     = el.TryGetProperty("FailWindow",     out v) ? v.GetInt32() : 3,
                        WatchEnabled   = el.TryGetProperty("WatchEnabled",   out v) && v.GetBoolean(),
                        WatchIntervalSecs = el.TryGetProperty("WatchIntervalSecs", out v) ? v.GetInt32() : 300,
                        LastGenerated  = el.TryGetProperty("LastGenerated",  out v) && v.ValueKind != System.Text.Json.JsonValueKind.Null
                                         ? v.GetDateTime() : (DateTime?)null,
                        KnownRuns      = el.TryGetProperty("KnownRuns", out v)
                                         ? v.EnumerateArray().Select(r => r.GetString() ?? "").Where(s => s.Length > 0).ToList()
                                         : new List<string>(),
                        MaxMonths      = el.TryGetProperty("MaxMonths",      out v) ? v.GetInt32() : 0,
                        IncludeOldRuns = el.TryGetProperty("IncludeOldRuns", out v) && v.GetBoolean()
                    };
                    _dbTrendsLibrary.Add(c);
                }
                RefreshDbTrendsLibraryUI();
            }
            catch { }
        }

        // ── DB Export / Import ───────────────────────────────────────────────────

        private void DbTrendsLibraryExport_Click(object sender, RoutedEventArgs e)
        {
            if (_dbTrendsLibrary.Count == 0)
            { DarkMessageBox.Show("The DB trends library is empty — nothing to export.", "Export Library"); return; }

            var dlg = new SaveFileDialog
            {
                Title = "Export DB Trends Library",
                Filter = "JSON backup (*.json)|*.json",
                FileName = $"DbTrendsLibrary_backup_{DateTime.Now:yyyyMMdd}.json"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var dtos = _dbTrendsLibrary.Select(c => new
                {
                    c.Id, c.Name, c.ApiHost, c.ApiKey, c.DownloadFolder, c.ReportsFolder,
                    c.FailWindow, c.WatchIntervalSecs, c.LastGenerated,
                    c.MaxMonths, c.IncludeOldRuns, c.KnownRuns
                });
                var json = System.Text.Json.JsonSerializer.Serialize(dtos,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                System.IO.File.WriteAllText(dlg.FileName, json);
                DarkMessageBox.Show($"Exported {_dbTrendsLibrary.Count} customer(s) to:\n{dlg.FileName}", "Export Complete");
            }
            catch (Exception ex) { DarkMessageBox.Show($"Export failed: {ex.Message}", "Export Failed"); }
        }

        private void DbTrendsLibraryImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Import DB Trends Library",
                Filter = "JSON backup (*.json)|*.json|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var json = System.IO.File.ReadAllText(dlg.FileName);
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                var existingNames = new HashSet<string>(_dbTrendsLibrary.Select(c => c.Name), StringComparer.OrdinalIgnoreCase);
                int added = 0;
                foreach (var el in doc.RootElement.EnumerateArray())
                {
                    string name = el.TryGetProperty("Name", out var v) ? v.GetString() ?? "" : "";
                    if (existingNames.Contains(name)) continue;
                    _dbTrendsLibrary.Add(new DbTrendsCustomer
                    {
                        Name           = name,
                        ApiHost        = el.TryGetProperty("ApiHost",        out v) ? v.GetString() ?? "" : "",
                        ApiKey         = el.TryGetProperty("ApiKey",         out v) ? v.GetString() ?? "" : "",
                        DownloadFolder = el.TryGetProperty("DownloadFolder", out v) ? v.GetString() ?? "" : "",
                        ReportsFolder  = el.TryGetProperty("ReportsFolder",  out v) ? v.GetString() ?? "" : "",
                        FailWindow     = el.TryGetProperty("FailWindow",     out v) ? v.GetInt32() : 3,
                        WatchIntervalSecs = el.TryGetProperty("WatchIntervalSecs", out v) ? v.GetInt32() : 300,
                        MaxMonths      = el.TryGetProperty("MaxMonths",      out v) ? v.GetInt32() : 0,
                        IncludeOldRuns = el.TryGetProperty("IncludeOldRuns", out v) && v.GetBoolean(),
                        KnownRuns      = el.TryGetProperty("KnownRuns", out v)
                                         ? v.EnumerateArray().Select(r => r.GetString() ?? "").Where(s => s.Length > 0).ToList()
                                         : new List<string>(),
                    });
                    existingNames.Add(name);
                    added++;
                }
                if (added == 0)
                { DarkMessageBox.Show("No new customers found — all entries already exist.", "Nothing to Import"); return; }

                SaveDbTrendsLibrary();
                RefreshDbTrendsLibraryUI();
                DbTrendsStatusLabel.Text = $"\u2714 Imported {added} new customer(s).";
                DbTrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
            }
            catch (Exception ex) { DarkMessageBox.Show($"Import failed: {ex.Message}", "Import Failed"); }
        }

        // ── DB Auto-Watch ────────────────────────────────────────────────────────

        private void DbTrendsWatchToggle_Click(object sender, RoutedEventArgs e)
        {
            if (_activeDbTrendsCustomer == null) return;
            if (_dbWatchTimers.ContainsKey(_activeDbTrendsCustomer.Id))
                DbStopCustomerWatch(_activeDbTrendsCustomer.Id);
            else
                DbStartCustomerWatch(_activeDbTrendsCustomer);
            DbRefreshWatchAllBtn();
        }

        private void DbTrendsWatchAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var c in _dbTrendsLibrary)
                if (!string.IsNullOrEmpty(c.ApiHost) && !string.IsNullOrEmpty(c.DownloadFolder)
                    && !_dbWatchTimers.ContainsKey(c.Id))
                    DbStartCustomerWatch(c, silent: true);
            DbRefreshWatchAllBtn();
        }

        private void DbTrendsStopAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var id in _dbWatchTimers.Keys.ToList())
                DbStopCustomerWatch(id);
            DbRefreshWatchAllBtn();
        }

        private void DbStartCustomerWatch(DbTrendsCustomer customer, bool silent = false)
        {
            if (_dbWatchTimers.ContainsKey(customer.Id)) return;
            if (string.IsNullOrEmpty(customer.ApiHost) || string.IsNullOrEmpty(customer.DownloadFolder))
            {
                if (!silent) DarkMessageBox.Show("Set an API Host and Download folder before starting auto-watch.", "Required");
                return;
            }

            int intervalSecs = customer.WatchIntervalSecs > 0 ? customer.WatchIntervalSecs : 300;
            var timer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(intervalSecs)
            };
            string capturedId = customer.Id;
            timer.Tick += async (_, _) =>
            {
                var c = _dbTrendsLibrary.FirstOrDefault(x => x.Id == capturedId);
                if (c == null) { DbStopCustomerWatch(capturedId); return; }
                await DbCustomerWatchTickAsync(c);
            };
            timer.Start();
            _dbWatchTimers[customer.Id] = timer;

            customer.WatchEnabled = true;
            SaveDbTrendsLibrary();

            if (_activeDbTrendsCustomer?.Id == customer.Id)
            {
                DbTrendsWatchToggleBtn.Content    = "⏹ Stop Watch";
                DbTrendsWatchToggleBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1F2D20"));
                DbTrendsWatchToggleBtn.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
                DbTrendsWatchStatusLabel.Text       = $"Watching · every {intervalSecs}s · polling now…";
                DbTrendsWatchStatusLabel.Visibility = Visibility.Visible;
            }

            // Kick off first poll immediately
            Dispatcher.BeginInvoke(new Action(async () =>
            {
                var c = _dbTrendsLibrary.FirstOrDefault(x => x.Id == capturedId);
                if (c != null) await DbCustomerWatchTickAsync(c);
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void DbStopCustomerWatch(string customerId)
        {
            if (_dbWatchTimers.TryGetValue(customerId, out var timer))
            {
                timer.Stop();
                _dbWatchTimers.Remove(customerId);
            }
            var c = _dbTrendsLibrary.FirstOrDefault(x => x.Id == customerId);
            if (c != null)
            {
                c.WatchEnabled = false;
                SaveDbTrendsLibrary();
                if (_activeDbTrendsCustomer?.Id == customerId)
                {
                    DbTrendsWatchToggleBtn.Content    = "\U0001F441 Auto-Watch";
                    DbTrendsWatchToggleBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640"));
                    DbTrendsWatchToggleBtn.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#60A5FA"));
                    DbTrendsWatchStatusLabel.Text       = "";
                    DbTrendsWatchStatusLabel.Visibility = Visibility.Collapsed;
                }
            }
        }

        private async Task DbCustomerWatchTickAsync(DbTrendsCustomer customer)
        {
            try
            {
                var host     = customer.ApiHost.TrimEnd('/');
                var encoded  = Uri.EscapeDataString(customer.Name);
                var url      = $"{host}/spring/runNumber/getRunNumberByProjectName?projectName={encoded}";

                using var http = new System.Net.Http.HttpClient { Timeout = TimeSpan.FromSeconds(30) };
                if (!string.IsNullOrEmpty(customer.ApiKey))
                    http.DefaultRequestHeaders.Add("Authorization", $"Bearer {customer.ApiKey}");

                var json = await http.GetStringAsync(url);
                var raw  = System.Text.Json.JsonSerializer.Deserialize<List<string>>(json) ?? new();
                var latest = raw.Select(s => s.Contains(',') ? s[..s.LastIndexOf(',')] : s)
                               .Where(s => !string.IsNullOrWhiteSpace(s))
                               .ToList();

                // Merge API results into KnownRuns so we never lose track of a run
                // even if the API stops returning it in future polls (e.g. it trims old runs).
                foreach (var r in latest)
                    if (!customer.KnownRuns.Contains(r, StringComparer.OrdinalIgnoreCase))
                        customer.KnownRuns.Add(r);

                // Determine which runs are missing from disk — check against ALL known runs,
                // not just the current API response. This means a file deleted locally is
                // always re-downloaded on the next tick, even if the API no longer lists it.
                var prefix = customer.Name + "_";
                var onDisk = System.IO.Directory.Exists(customer.DownloadFolder)
                    ? System.IO.Directory.GetFiles(customer.DownloadFolder, "*.xlsx")
                        .Select(f => System.IO.Path.GetFileNameWithoutExtension(f))
                        .Where(n => n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                        .Select(n => n[prefix.Length..])
                        .ToHashSet(StringComparer.OrdinalIgnoreCase)
                    : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Runs that should be on disk but aren't — could be brand-new (from API)
                // or previously downloaded and later deleted (from KnownRuns).
                var missingRuns = customer.KnownRuns
                    .Where(r => !onDisk.Contains(r))
                    .ToList();

                if (latest.Count == 0 && onDisk.Count == 0 && customer.KnownRuns.Count == 0)
                {
                    // API returned nothing and we have no history — nothing to do
                    if (_activeDbTrendsCustomer?.Id == customer.Id)
                        DbTrendsWatchStatusLabel.Text = $"Last poll: {DateTime.Now:HH:mm:ss} · download folder empty, skipping";
                    return;
                }
                else if (missingRuns.Count > 0)
                {
                    if (_activeDbTrendsCustomer?.Id == customer.Id)
                        DbTrendsWatchStatusLabel.Text = $"Missing {missingRuns.Count} file(s): {string.Join(", ", missingRuns)} · downloading…";

                    // Try to download each missing run — some may no longer be on the API
                    // (deleted remotely), so failures are non-fatal and logged only.
                    foreach (var run in missingRuns)
                    {
                        var dlUrl  = $"{host}/spring/export-to-excel?projectName={encoded}&runNumber={Uri.EscapeDataString(run)}&result=final";
                        var target = System.IO.Path.Combine(customer.DownloadFolder, $"{customer.Name}_{run}.xlsx");
                        try
                        {
                            var bytes = await http.GetByteArrayAsync(dlUrl);
                            await System.IO.File.WriteAllBytesAsync(target, bytes);
                        }
                        catch
                        {
                            // Run may no longer exist on the API — log via watch status if active
                            if (_activeDbTrendsCustomer?.Id == customer.Id)
                                Dispatcher.Invoke(() =>
                                    DbTrendsWatchStatusLabel.Text = $"Last poll: {DateTime.Now:HH:mm:ss} · {run} unavailable on API, skipped");
                        }
                    }
                }
                else
                {
                    // All known runs are present on disk — no downloads needed, still regenerate
                    if (_activeDbTrendsCustomer?.Id == customer.Id)
                        DbTrendsWatchStatusLabel.Text = $"Last poll: {DateTime.Now:HH:mm:ss} · no missing files, regenerating…";
                }

                // Re-generate trends using the customer's saved settings directly
                // (do NOT call LoadDbTrendsCustomer — it resets button states)
                int    dbWatchFw         = customer.FailWindow > 0 ? customer.FailWindow : 3;
                int    dbWatchMaxMonths  = customer.MaxMonths;
                bool   dbWatchIncludeOld = customer.IncludeOldRuns;
                string dbWatchDlFolder   = customer.DownloadFolder;
                string dbWatchRptFolder  = string.IsNullOrEmpty(customer.ReportsFolder)
                                           ? customer.DownloadFolder : customer.ReportsFolder;
                bool   isCurrent         = _activeDbTrendsCustomer?.Id == customer.Id;

                if (isCurrent)
                {
                    DbTrendsLogPanel.Visibility = Visibility.Visible;
                    DbTrendsProgress.Visibility = Visibility.Visible;
                    DbTrendsLog.Text            = "";
                }

                await System.Threading.Tasks.Task.Run(() =>
                {
                    Action<string>? wLog = isCurrent
                        ? msg => Dispatcher.Invoke(() =>
                          {
                              DbTrendsLog.Text += msg + "\n";
                              var sv = DbTrendsLog.Parent as ScrollViewer
                                    ?? (DbTrendsLog.Parent as FrameworkElement)?.Parent as ScrollViewer;
                              sv?.ScrollToEnd();
                          })
                        : null;

                    var (ok, outputPath, error) = TestRunTrendsProcessor.Generate(
                        wLog, dbWatchDlFolder, customer.Name, dbWatchRptFolder,
                        dbWatchFw, dbWatchMaxMonths, dbWatchIncludeOld);

                    Dispatcher.Invoke(() =>
                    {
                        if (isCurrent) DbTrendsProgress.Visibility = Visibility.Collapsed;
                        customer.LastGenerated = DateTime.Now;
                        SaveDbTrendsLibrary();
                        RefreshDbTrendsLibraryUI();

                        if (isCurrent)
                        {
                            // Re-enable buttons so user can still interact after watch run
                            DbTrendsDownloadBtn.IsEnabled = true;
                            DbTrendsRunBtn.IsEnabled      = true;

                            if (ok)
                            {
                                string shortName = System.IO.Path.GetFileName(outputPath);
                                DbTrendsStatusLabel.Text = $"\u2714 Auto-updated {DateTime.Now:HH:mm:ss} \u2192 {shortName}";
                                DbTrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
                                DbTrendsWatchStatusLabel.Text  = $"Updated {DateTime.Now:HH:mm:ss} · {missingRuns.Count} file(s) downloaded";
                            }
                            else
                            {
                                DbTrendsStatusLabel.Text = $"\u2716 Watch generate failed: {error}";
                                DbTrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71));
                                DbTrendsWatchStatusLabel.Text  = $"Error {DateTime.Now:HH:mm:ss}: {error}";
                            }
                        }
                    });
                });
            }
            catch (Exception ex)
            {
                if (_activeDbTrendsCustomer?.Id == customer.Id)
                    DbTrendsWatchStatusLabel.Text = $"Poll error {DateTime.Now:HH:mm:ss}: {ex.Message}";
            }
        }

        private void DbRefreshWatchAllBtn()
        {
            int watching = _dbWatchTimers.Count;
            DbTrendsWatchAllBadge.Text       = watching > 0 ? $"({watching})" : "";
            DbTrendsWatchAllBadge.Visibility = watching > 0 ? Visibility.Visible : Visibility.Collapsed;
            DbTrendsStopAllBtn.IsEnabled     = watching > 0;
            DbTrendsStopAllBadge.Text        = watching > 0 ? $"({watching})" : "";
            DbTrendsStopAllBadge.Visibility  = watching > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        // ── Test Run Trends ─────────────────────────────────────────────────────

        private static readonly string TrendsLibraryPath = AppDataManager.TrendsLibraryPath;

        private class TrendsCustomer
        {
            public string Id          { get; set; } = Guid.NewGuid().ToString("N")[..8];
            public string Name        { get; set; } = "";
            public string RunsFolder  { get; set; } = "";
            public string ReportsFolder { get; set; } = "";
            public DateTime? LastGenerated { get; set; } = null;
            public string?   LastOutput    { get; set; } = null;
            /// <summary>Per-customer fail window. 0 = use the global setting.</summary>
            public int       FailWindow        { get; set; } = 0;
            /// <summary>Per-customer watch interval in seconds. 0 = use the global setting.</summary>
            public int       WatchIntervalSecs { get; set; } = 0;
            /// <summary>
            /// Persisted watch state — true if auto-watch was active for this customer
            /// when the app last saved.  Used to restore watches after restart/reboot.
            /// </summary>
            public bool      WatchEnabled      { get; set; } = false;
            public int       MaxMonths         { get; set; } = 0;
            public bool      IncludeOldRuns    { get; set; } = false;
        }

        private List<TrendsCustomer> _trendsLibrary = new();
        private string? _trendsRunsFolder    = null;
        private string? _trendsReportsFolder = null;

        private void LoadTrendsLibrary()
        {
            var dtos = AppDataManager.LoadTrendsLibrary();
            _trendsLibrary = dtos.Select(d => new TrendsCustomer
            {
                Id            = d.Id,
                Name          = d.Name,
                RunsFolder    = d.RunsFolder,
                ReportsFolder = d.ReportsFolder,
                LastGenerated = d.LastGenerated,
                LastOutput    = d.LastOutput,
                FailWindow    = d.FailWindow,
                WatchIntervalSecs = d.WatchIntervalSecs,
                WatchEnabled      = d.WatchEnabled,
            }).ToList();
            RefreshTrendsLibraryUI();
        }

        private void SaveTrendsLibrary()
        {
            var dtos = _trendsLibrary.Select(c => new AppDataManager.TrendsCustomerDto
            {
                Id            = c.Id,
                Name          = c.Name,
                RunsFolder    = c.RunsFolder,
                ReportsFolder = c.ReportsFolder,
                LastGenerated = c.LastGenerated,
                LastOutput    = c.LastOutput,
                FailWindow    = c.FailWindow,
                WatchIntervalSecs = c.WatchIntervalSecs,
                WatchEnabled      = c.WatchEnabled,
            }).ToList();
            AppDataManager.SaveTrendsLibrary(dtos);
        }

        private void RefreshTrendsLibraryUI()
        {
            TrendsLibraryPanel.Children.Clear();
            TrendsLibraryEmptyLabel.Visibility = _trendsLibrary.Count == 0
                ? Visibility.Visible : Visibility.Collapsed;

            foreach (var customer in _trendsLibrary.OrderBy(c => c.Name))
                TrendsLibraryPanel.Children.Add(BuildTrendsCard(customer));
        }

        private UIElement BuildTrendsCard(TrendsCustomer customer)
        {
            var card = new Border
            {
                Background      = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0D1020")),
                BorderBrush     = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640")),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(6),
                Margin          = new Thickness(2, 2, 2, 2),
                Padding         = new Thickness(10, 8, 10, 8),
                Cursor          = System.Windows.Input.Cursors.Hand,
                Tag             = customer
            };

            var outer = new Grid();
            outer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            outer.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var stack = new StackPanel();

            // Customer name
            stack.Children.Add(new TextBlock
            {
                Text       = customer.Name,
                FontSize   = 12.5, FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E8F0")),
                FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                Margin     = new Thickness(0, 0, 0, 3)
            });

            // Runs folder
            stack.Children.Add(new TextBlock
            {
                Text       = "Runs: " + (string.IsNullOrEmpty(customer.RunsFolder) ? "—"
                    : System.IO.Path.GetFileName(customer.RunsFolder.TrimEnd('\\', '/')) == ""
                        ? customer.RunsFolder : "…\\" + System.IO.Path.GetFileName(customer.RunsFolder.TrimEnd('\\', '/'))),
                FontSize   = 10.5, Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B7A99")),
                FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                ToolTip    = customer.RunsFolder
            });

            // Reports folder
            stack.Children.Add(new TextBlock
            {
                Text       = "Reports: " + (string.IsNullOrEmpty(customer.ReportsFolder) ? "same as Runs"
                    : "…\\" + System.IO.Path.GetFileName(customer.ReportsFolder.TrimEnd('\\', '/'))),
                FontSize   = 10.5, Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B7A99")),
                FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                ToolTip    = customer.ReportsFolder
            });

            // Last generated
            if (customer.LastGenerated.HasValue)
                stack.Children.Add(new TextBlock
                {
                    Text       = $"Last: {customer.LastGenerated.Value:dd MMM yyyy HH:mm}",
                    FontSize   = 10, Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4A5F88")),
                    FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                    Margin     = new Thickness(0, 3, 0, 0)
                });

            Grid.SetColumn(stack, 0);
            outer.Children.Add(stack);

            // Delete button
            var delBtn = new Button
            {
                Content         = new TextBlock { Text = "Del", FontSize = 9, FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F87171")),
                    FontFamily = new FontFamily("Segoe UI Variable, Segoe UI") },
                Width = 38, Height = 20,
                Background      = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2D0F0F")),
                BorderBrush     = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#5C1A1A")),
                BorderThickness = new Thickness(1),
                Cursor          = System.Windows.Input.Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Top,
                Margin          = new Thickness(4, 0, 0, 0),
                Tag             = customer
            };
            delBtn.Click += (s, ev) =>
            {
                ev.Handled = true;
                var c = (TrendsCustomer)((Button)s).Tag;
                if (!DarkMessageBox.Confirm($"Remove '{c.Name}' from library?", "Remove Customer")) return;
                _trendsLibrary.Remove(c);
                SaveTrendsLibrary();
                RefreshTrendsLibraryUI();
            };
            Grid.SetColumn(delBtn, 1);
            outer.Children.Add(delBtn);

            card.Child = outer;
            card.MouseLeftButtonUp += (s, _) => LoadTrendsCustomer((TrendsCustomer)((Border)s).Tag);
            card.MouseEnter += (s, _) => ((Border)s).BorderBrush =
                new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB"));
            card.MouseLeave += (s, _) => ((Border)s).BorderBrush =
                new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640"));

            // Right-click context menu
            var menu = new ContextMenu();
            menu.Style = null;
            var menuGenerate = new MenuItem { Header = "Generate Trends" };
            menuGenerate.Click += (s, ev) =>
            {
                LoadTrendsCustomer(customer);
                RunTrendsGeneration(silent: false);
            };
            var menuOpen = new MenuItem
            {
                Header    = "Open Last Output",
                IsEnabled = !string.IsNullOrEmpty(customer.LastOutput)
                         && System.IO.File.Exists(customer.LastOutput)
            };
            menuOpen.Click += (s, ev) =>
            {
                if (!string.IsNullOrEmpty(customer.LastOutput) && System.IO.File.Exists(customer.LastOutput))
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        { FileName = customer.LastOutput, UseShellExecute = true });
            };
            var menuExport = new MenuItem { Header = "Export this customer..." };
            menuExport.Click += (s, ev) => ExportSingleCustomer(customer);
            var menuDelete = new MenuItem { Header = "Delete" };
            menuDelete.Click += (s, ev) =>
            {
                if (!DarkMessageBox.Confirm($"Remove '{customer.Name}' from library?", "Remove Customer")) return;
                _trendsLibrary.Remove(customer);
                SaveTrendsLibrary();
                RefreshTrendsLibraryUI();
            };
            menu.Items.Add(menuGenerate);
            menu.Items.Add(menuOpen);
            menu.Items.Add(new Separator());
            menu.Items.Add(menuExport);
            menu.Items.Add(new Separator());
            menu.Items.Add(menuDelete);
            card.ContextMenu = menu;

            return card;
        }

        private TrendsCustomer? _activeTrendsCustomer = null;
        private bool            _trendsCustomerDirty  = false;
        private bool            _suppressDirtyTracking = false;

        private void LoadTrendsCustomer(TrendsCustomer c)
        {
            _suppressDirtyTracking = true;
            _activeTrendsCustomer  = c;
            _trendsCustomerDirty   = false;

            TrendsCustomerNameBox.Text = c.Name;
            _trendsRunsFolder    = c.RunsFolder;
            _trendsReportsFolder = c.ReportsFolder;
            // Load per-customer fail window; 0 means "use global" so leave box showing global value
            TrendsFailWindowBox.Text = c.FailWindow > 0
                ? c.FailWindow.ToString()
                : _settings.TrendsFailWindow;
            TrendsMaxMonthsBox.Text         = c.MaxMonths > 0 ? c.MaxMonths.ToString() : "0";
            TrendsWatchIntervalBox.Text     = c.WatchIntervalSecs > 0 ? c.WatchIntervalSecs.ToString() : "60";
            TrendsIncludeOldRunsCheck.IsChecked = c.IncludeOldRuns;
            TrendsRunsFolderLabel.Text = string.IsNullOrEmpty(c.RunsFolder) ? "Not set" : c.RunsFolder;
            TrendsRunsFolderLabel.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                string.IsNullOrEmpty(c.RunsFolder) ? "#4A5F88" : "#CBD5E1"));
            TrendsReportsFolderLabel.Text = string.IsNullOrEmpty(c.ReportsFolder) ? "Same as Runs folder" : c.ReportsFolder;
            TrendsReportsFolderLabel.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                string.IsNullOrEmpty(c.ReportsFolder) ? "#4A5F88" : "#CBD5E1"));
            TrendsUpdateLibraryBtn.Visibility = Visibility.Collapsed;

            // Clear any stale status message from the previously loaded customer
            // (e.g. "Unsaved changes — OtherCustomer" must not linger after switching)
            TrendsStatusLabel.Text = "";

            // Sync the per-customer Watch toggle button to the actual watch state of
            // the newly-loaded customer — avoids carrying over the previous customer's
            // button state (e.g. Encova "Stop Watch" bleeding into Intersnack).
            bool isWatched = _watchTimers.ContainsKey(c.Id);
            TrendsWatchToggleBtn.Content    = isWatched ? "⏹ Stop Watch" : "👁 Auto-Watch";
            TrendsWatchToggleBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                isWatched ? "#1F2D20" : "#1E2640"));
            TrendsWatchToggleBtn.Foreground = isWatched
                ? new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80))
                : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#60A5FA"));
            TrendsWatchStatusLabel.Text       = isWatched ? TrendsWatchStatusLabel.Text : "";
            TrendsWatchStatusLabel.Visibility = isWatched ? Visibility.Visible : Visibility.Collapsed;

            UpdateTrendsFilesCount();

            _suppressDirtyTracking = false;
        }

        private void MarkTrendsCustomerDirty()
        {
            if (_suppressDirtyTracking) return;
            if (_activeTrendsCustomer == null) return;
            if (_trendsCustomerDirty) return;
            _trendsCustomerDirty = true;
            TrendsUpdateLibraryBtn.Visibility = Visibility.Visible;
            TrendsStatusLabel.Text = $"Unsaved changes — {_activeTrendsCustomer.Name}";
            TrendsStatusLabel.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FBBF24"));
        }

        private void TrendsUpdateLibrary_Click(object sender, RoutedEventArgs e)
        {
            if (_activeTrendsCustomer == null) return;

            var existing = _trendsLibrary.FirstOrDefault(c =>
                c.Id == _activeTrendsCustomer.Id);
            if (existing == null) return;

            existing.Name          = TrendsCustomerNameBox.Text.Trim();
            existing.RunsFolder    = _trendsRunsFolder    ?? "";
            existing.ReportsFolder = _trendsReportsFolder ?? "";
            existing.FailWindow    = int.TryParse(TrendsFailWindowBox.Text.Trim(), out int ufv) && ufv >= 1 ? ufv : 0;
            existing.MaxMonths     = int.TryParse(TrendsMaxMonthsBox.Text.Trim(), out int umm) && umm > 0 ? umm : 0;
            existing.IncludeOldRuns = TrendsIncludeOldRunsCheck.IsChecked == true;
            existing.WatchIntervalSecs = int.TryParse(TrendsWatchIntervalBox.Text.Trim(), out int uwi) && uwi > 0 ? uwi : 60;

            SaveTrendsLibrary();
            RefreshTrendsLibraryUI();

            _activeTrendsCustomer = existing;
            _trendsCustomerDirty  = false;
            TrendsUpdateLibraryBtn.Visibility = Visibility.Collapsed;
            TrendsStatusLabel.Text = $"Updated: {existing.Name}";
            TrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
        }

        private void TrendsLibraryAdd_Click(object sender, RoutedEventArgs e)
        {
            string name = TrendsCustomerNameBox.Text.Trim();
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(_trendsRunsFolder))
            {
                DarkMessageBox.Show("Enter a customer name and select the Runs folder first.",
                    "Missing Info");
                return;
            }
            int addFw = int.TryParse(TrendsFailWindowBox.Text.Trim(), out int afv) && afv >= 1 ? afv : 0;
            int addMm = int.TryParse(TrendsMaxMonthsBox.Text.Trim(), out int amm) && amm > 0 ? amm : 0;
            int addWi = int.TryParse(TrendsWatchIntervalBox.Text.Trim(), out int awi) && awi > 0 ? awi : 60;
            bool addOld = TrendsIncludeOldRunsCheck.IsChecked == true;
            var existing = _trendsLibrary.FirstOrDefault(c =>
                c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.RunsFolder     = _trendsRunsFolder;
                existing.ReportsFolder  = _trendsReportsFolder ?? "";
                existing.FailWindow     = addFw;
                existing.MaxMonths      = addMm;
                existing.IncludeOldRuns = addOld;
                existing.WatchIntervalSecs = addWi;
            }
            else
            {
                _trendsLibrary.Add(new TrendsCustomer
                {
                    Name          = name,
                    RunsFolder    = _trendsRunsFolder,
                    ReportsFolder = _trendsReportsFolder ?? "",
                    FailWindow    = addFw,
                    WatchIntervalSecs = addWi,
                });
            }
            SaveTrendsLibrary();
            RefreshTrendsLibraryUI();
            TrendsStatusLabel.Text = $"'{name}' saved to library.";
            TrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
        }

        private string BrowseFolder(string title)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = title, Filter = "Folder|*.folder", FileName = "Select Folder",
                CheckFileExists = false, CheckPathExists = true, ValidateNames = false
            };
            return dlg.ShowDialog() == true
                ? System.IO.Path.GetDirectoryName(dlg.FileName) ?? ""
                : "";
        }

        private void TrendsRunsFolderBrowse_Click(object sender, RoutedEventArgs e)
        {
            string f = BrowseFolder("Select Runs folder (contains monthly Excel files)");
            if (string.IsNullOrEmpty(f)) return;
            _trendsRunsFolder = f;
            TrendsRunsFolderLabel.Text = f;
            TrendsRunsFolderLabel.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#CBD5E1"));
            if (string.IsNullOrWhiteSpace(TrendsCustomerNameBox.Text))
                TrendsCustomerNameBox.Text = System.IO.Path.GetFileName(
                    System.IO.Path.GetDirectoryName(f.TrimEnd('\\', '/')) ?? f);
            UpdateTrendsFilesCount();
            MarkTrendsCustomerDirty();
        }

        private void TrendsReportsFolderBrowse_Click(object sender, RoutedEventArgs e)
        {
            string f = BrowseFolder("Select Reports folder (where Trends.xlsx will be saved)");
            if (string.IsNullOrEmpty(f)) return;
            _trendsReportsFolder = f;
            TrendsReportsFolderLabel.Text = f;
            TrendsReportsFolderLabel.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#CBD5E1"));
            MarkTrendsCustomerDirty();
        }

        private void UpdateTrendsFilesCount()
        {
            if (string.IsNullOrEmpty(_trendsRunsFolder) || !System.IO.Directory.Exists(_trendsRunsFolder))
            { TrendsFilesFoundLabel.Text = ""; return; }
            int count = System.IO.Directory.GetFiles(_trendsRunsFolder, "*.xlsx")
                .Count(f => !System.IO.Path.GetFileName(f).EndsWith("_Trends.xlsx",
                    StringComparison.OrdinalIgnoreCase));
            TrendsFilesFoundLabel.Text = $"{count} run file(s) found in Runs folder.";
        }

        private void TrendsRun_Click(object sender, RoutedEventArgs e)
            => RunTrendsGeneration(silent: false);

        /// <summary>
        /// Core generation logic shared by the manual button and the auto-watch timer.
        /// When <paramref name="silent"/> is true the "Open it now?" prompt is suppressed.
        /// </summary>
        private void RunTrendsGeneration(bool silent, Action<bool>? onComplete = null)
        {
            string name = TrendsCustomerNameBox.Text.Trim();
            if (string.IsNullOrEmpty(name))
            { if (!silent) DarkMessageBox.Show("Enter a customer name.", "Required"); return; }
            if (string.IsNullOrEmpty(_trendsRunsFolder) || !System.IO.Directory.Exists(_trendsRunsFolder))
            { if (!silent) DarkMessageBox.Show("Select a valid Runs folder.", "Required"); return; }

            if (!int.TryParse(TrendsFailWindowBox.Text.Trim(), out int failWindow) || failWindow < 1)
                failWindow = 3;
            // If the active customer has its own override, use that
            if (_activeTrendsCustomer?.FailWindow > 0)
                failWindow = _activeTrendsCustomer.FailWindow;

            string reportsFolder = string.IsNullOrEmpty(_trendsReportsFolder)
                ? _trendsRunsFolder : _trendsReportsFolder;

            TrendsLogPanel.Visibility = Visibility.Visible;
            TrendsProgress.Visibility = Visibility.Visible;
            if (!silent) TrendsLog.Text = "";
            TrendsRunBtn.IsEnabled = false;
            TrendsStatusLabel.Text = silent ? "Auto-generating…" : "Generating...";
            TrendsStatusLabel.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#60A5FA"));

            string runsFolder    = _trendsRunsFolder;
            string customerName  = name;
            int    fw            = failWindow;
            int    maxMonths     = int.TryParse(TrendsMaxMonthsBox.Text.Trim(), out int mm) && mm > 0 ? mm : 0;
            bool   includeOldRuns = TrendsIncludeOldRunsCheck.IsChecked == true;

            System.Threading.Tasks.Task.Run(() =>
            {
                Action<string> log = msg =>
                    Dispatcher.Invoke(() =>
                    {
                        TrendsAppendLog(msg);
                    });

                var (ok, outputPath, error) =
                    TestRunTrendsProcessor.Generate(log, runsFolder, customerName, reportsFolder, fw, maxMonths, includeOldRuns);

                Dispatcher.Invoke(() =>
                {
                    TrendsProgress.Visibility = Visibility.Collapsed;
                    TrendsRunBtn.IsEnabled    = true;

                    if (ok)
                    {
                        var entry = _trendsLibrary.FirstOrDefault(c =>
                            c.Name.Equals(customerName, StringComparison.OrdinalIgnoreCase));
                        if (entry != null)
                        { entry.LastGenerated = DateTime.Now; entry.LastOutput = outputPath; }
                        SaveTrendsLibrary();
                        RefreshTrendsLibraryUI();

                        string shortName = System.IO.Path.GetFileName(outputPath);
                        TrendsStatusLabel.Text = silent
                            ? $"Auto-updated {DateTime.Now:HH:mm:ss} → {shortName}"
                            : $"Saved to {shortName}";
                        TrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));

                        if (!silent && DarkMessageBox.Confirm($"Done:\n{outputPath}\n\nOpen it now?", "Trends Generated"))
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                { FileName = outputPath, UseShellExecute = true });
                    }
                    else
                    {
                        TrendsStatusLabel.Text = $"Failed: {error}";
                        TrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71));
                    }

                    // Invoke completion callback — passes success flag so caller
                    // only writes the manifest on success and clears the generating lock
                    onComplete?.Invoke(ok);
                });
            });
        }

        // ── Generate All ──────────────────────────────────────────────────────

        /// <summary>
        /// Regenerates trends for every valid library customer in parallel.
        /// Customers already being generated (by a watch tick) are skipped.
        /// The button label shows live progress (e.g. "Generating 2/5…") and
        /// reverts to "⚡ Generate All" once all tasks complete.
        /// </summary>
        private async void TrendsGenerateAll_Click(object sender, RoutedEventArgs e)
        {
            var eligible = _trendsLibrary
                .Where(c => !string.IsNullOrEmpty(c.RunsFolder)
                         && System.IO.Directory.Exists(c.RunsFolder)
                         && !_generatingIds.Contains(c.Id))
                .ToList();

            if (eligible.Count == 0)
            {
                DarkMessageBox.Show(
                    "No customers are ready to generate.\n\nMake sure each customer has a valid Runs folder set.",
                    "Generate All");
                return;
            }

            // Disable button and show progress label
            TrendsGenerateAllBtn.IsEnabled = false;
            TrendsGenerateAllBtn.Content   = $"Generating 0/{eligible.Count}…";
            TrendsGenerateAllBtn.Foreground =
                new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FBBF24"));

            int completed = 0;
            int succeeded = 0;
            var failedNames = new System.Collections.Generic.List<string>();
            var tasks = new System.Collections.Generic.List<System.Threading.Tasks.Task>();

            foreach (var customer in eligible)
            {
                // Snapshot everything needed — captures correct values per iteration
                var c            = customer;
                string runs      = c.RunsFolder;
                string name      = c.Name;
                string reports   = string.IsNullOrEmpty(c.ReportsFolder) ? runs : c.ReportsFolder;
                int fw           = c.FailWindow > 0
                    ? c.FailWindow
                    : (int.TryParse(_settings.TrendsFailWindow, out int gfw) && gfw >= 1 ? gfw : 3);

                _generatingIds.Add(c.Id);

                var task = System.Threading.Tasks.Task.Run(() =>
                {
                    // No UI log delegate — concurrent multi-customer output would interleave
                    // Everything goes to the app log file if logging is enabled
                    var (ok, outputPath, error) =
                        TestRunTrendsProcessor.Generate(null, runs, name, reports, fw);

                    Dispatcher.Invoke(() =>
                    {
                        _generatingIds.Remove(c.Id);
                        completed++;

                        if (ok)
                        {
                            succeeded++;
                            var entry = _trendsLibrary.FirstOrDefault(x => x.Id == c.Id);
                            if (entry != null)
                            { entry.LastGenerated = DateTime.Now; entry.LastOutput = outputPath; }
                        }
                        else
                        {
                            failedNames.Add(name);
                        }

                        // Update button progress label
                        TrendsGenerateAllBtn.Content =
                            completed < eligible.Count
                                ? $"Generating {completed}/{eligible.Count}…"
                                : "⚡ Generate All";
                    });
                });

                tasks.Add(task);
            }

            // Await all tasks off the UI thread so the window stays responsive
            await System.Threading.Tasks.Task.WhenAll(tasks);

            // All done — update library and show result
            SaveTrendsLibrary();
            RefreshTrendsLibraryUI();

            TrendsGenerateAllBtn.IsEnabled = true;
            TrendsGenerateAllBtn.Content   = "⚡ Generate All";
            TrendsGenerateAllBtn.Foreground =
                new SolidColorBrush((Color)ColorConverter.ConvertFromString("#34D399"));

            string summary = succeeded == eligible.Count
                ? $"All {succeeded} customer(s) generated successfully."
                : $"{succeeded}/{eligible.Count} succeeded.";
            if (failedNames.Count > 0)
                summary += $"\nFailed: {string.Join(", ", failedNames)}";

            TrendsStatusLabel.Text = summary;
            TrendsStatusLabel.Foreground = new SolidColorBrush(
                failedNames.Count == 0
                    ? Color.FromRgb(0x4A, 0xDE, 0x80)
                    : Color.FromRgb(0xFB, 0xBF, 0x24));
        }

        // ── Auto-watch (multi-customer) ───────────────────────────────────────

        // Each customer gets its own timer. Key = customer Id.
        private readonly Dictionary<string, System.Windows.Threading.DispatcherTimer>
            _watchTimers = new();
        // Track which customer is currently generating to avoid overlaps per-customer
        private readonly HashSet<string> _generatingIds = new();

        private TrayManager? _tray;

        private void InitTray()
        {
            try
            {
                _tray = new TrayManager(this)
                {
                    OnWatchAll          = WatchAll,
                    OnStopAll           = StopAll,
                    GetActiveWatchCount = () => _watchTimers.Count
                };
            }
            catch { /* tray unavailable — not fatal */ }
        }

        // Called from MainWindow() Loaded handler
        private void InitTrayOnLoad()
        {
            InitTray();
            // Restore per-customer watches from persisted WatchEnabled state.
            // Only customers that were actually being watched before the last
            // shutdown/crash are restarted — not all customers with a valid folder.
            if (_settings.TrendsAutoWatch)
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    foreach (var c in _trendsLibrary)
                    {
                        if (c.WatchEnabled &&
                            !string.IsNullOrEmpty(c.RunsFolder) &&
                            Directory.Exists(c.RunsFolder))
                            StartCustomerWatch(c, silent: true);
                    }

                    // Sync the Watch All / Stop All button to reflect restored watches
                    RefreshWatchAllBtn();

                    // If launched with --minimized (e.g. after server reboot via
                    // Task Scheduler), go straight to the system tray.
                    if (App.StartMinimized && _watchTimers.Count > 0)
                        _tray?.MinimizeToTray();
                }), System.Windows.Threading.DispatcherPriority.Loaded);

            // Restore DB watches unconditionally — DB watch state is persisted per-customer
            // and is independent of the TrendsAutoWatch global setting.
            Dispatcher.BeginInvoke(new Action(() =>
            {
                foreach (var c in _dbTrendsLibrary)
                {
                    if (c.WatchEnabled &&
                        !string.IsNullOrEmpty(c.ApiHost) &&
                        !string.IsNullOrEmpty(c.DownloadFolder))
                        DbStartCustomerWatch(c, silent: true);
                }
                DbRefreshWatchAllBtn();

                if (App.StartMinimized && _dbWatchTimers.Count > 0)
                    _tray?.MinimizeToTray();
            }), System.Windows.Threading.DispatcherPriority.Loaded);

            // Even without active watches, honour --minimized so the app
            // doesn't pop up a window on a headless server after reboot.
            if (App.StartMinimized && !_settings.TrendsAutoWatch)
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _tray?.MinimizeToTray();
                }), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private void TrendsWatchToggle_Click(object sender, RoutedEventArgs e)
        {
            string name = TrendsCustomerNameBox.Text.Trim();
            var customer = _trendsLibrary.FirstOrDefault(c =>
                c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            if (customer != null && _watchTimers.ContainsKey(customer.Id))
                StopCustomerWatch(customer.Id);
            else
                StartTrendsWatch();

            RefreshWatchAllBtn();
        }

        private void StartTrendsWatch(bool silent = false)
        {
            string name = TrendsCustomerNameBox.Text.Trim();
            if (string.IsNullOrEmpty(name))
            { if (!silent) DarkMessageBox.Show("Enter a customer name before starting auto-watch.", "Required"); return; }
            if (string.IsNullOrEmpty(_trendsRunsFolder) || !System.IO.Directory.Exists(_trendsRunsFolder))
            { if (!silent) DarkMessageBox.Show("Select a valid Runs folder before starting auto-watch.", "Required"); return; }

            // Find or create the customer object
            var customer = _trendsLibrary.FirstOrDefault(c =>
                c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (customer == null)
            {
                if (!silent) DarkMessageBox.Show("Save this customer to the library first before starting auto-watch.", "Not in Library");
                return;
            }

            StartCustomerWatch(customer, silent);
            RefreshWatchAllBtn();
        }

        private void StartCustomerWatch(TrendsCustomer customer, bool silent = false)
        {
            if (_watchTimers.ContainsKey(customer.Id)) return; // already watching

            TrendsManifest.Delete(customer.RunsFolder);

            int intervalSecs = customer.WatchIntervalSecs > 0
                ? customer.WatchIntervalSecs
                : (_settings.TrendsWatchIntervalSecs > 0 ? _settings.TrendsWatchIntervalSecs : 60);

            var timer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(intervalSecs)
            };
            // Capture only the Id — not the full customer object — so that if the
            // customer is removed from the library the next tick resolves null and
            // stops the timer, releasing the reference cleanly.
            string capturedId = customer.Id;
            timer.Tick += (_, _) =>
            {
                var c = _trendsLibrary.FirstOrDefault(x => x.Id == capturedId);
                if (c == null) { StopCustomerWatch(capturedId); return; }
                CustomerWatchTick(c);
            };
            timer.Start();
            _watchTimers[customer.Id] = timer;

            // If this is the currently displayed customer, update the UI
            if (IsCurrentCustomer(customer))
            {
                TrendsWatchToggleBtn.Content    = "⏹ Stop Watch";
                TrendsWatchToggleBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1F2D20"));
                TrendsWatchToggleBtn.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
                TrendsWatchStatusLabel.Text     = $"Watching · every {intervalSecs}s · scanning now…";
                TrendsWatchStatusLabel.Visibility = Visibility.Visible;
            }

            _settings.TrendsAutoWatch = true;
            AppDataManager.SaveSettings(_settings);

            // Persist per-customer watch state so it survives app/server restarts
            customer.WatchEnabled = true;
            SaveTrendsLibrary();

            // Auto-register Windows logon startup so the app restarts after a
            // server reboot — only registers once (idempotent).
            EnsureAutoStartRegistered();

            _tray?.UpdateTooltip(_watchTimers.Count);

            // Kick off first scan immediately (use capturedId for consistency)
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var c = _trendsLibrary.FirstOrDefault(x => x.Id == capturedId);
                if (c != null) CustomerWatchTick(c);
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void StopCustomerWatch(string customerId)
        {
            if (_watchTimers.TryGetValue(customerId, out var timer))
            {
                timer.Stop();
                _watchTimers.Remove(customerId);
            }

            var customer = _trendsLibrary.FirstOrDefault(c => c.Id == customerId);
            if (customer != null)
            {
                // Persist per-customer watch state
                customer.WatchEnabled = false;
                SaveTrendsLibrary();

                if (IsCurrentCustomer(customer))
                {
                    TrendsWatchToggleBtn.Content    = "👁 Auto-Watch";
                    TrendsWatchToggleBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640"));
                    TrendsWatchToggleBtn.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#60A5FA"));
                    TrendsWatchStatusLabel.Text     = "";
                    TrendsWatchStatusLabel.Visibility = Visibility.Collapsed;
                }
            }

            if (_watchTimers.Count == 0)
            {
                _settings.TrendsAutoWatch = false;
                AppDataManager.SaveSettings(_settings);

                // No active watchers — remove the Windows logon auto-start entry
                // so the app doesn't launch on reboot when nothing needs watching.
                EnsureAutoStartUnregistered();
            }
            _tray?.UpdateTooltip(_watchTimers.Count);
        }

        private void StopAll()
        {
            foreach (var id in _watchTimers.Keys.ToList())
                StopCustomerWatch(id);
            RefreshWatchAllBtn();
        }

        private void WatchAll()
        {
            foreach (var customer in _trendsLibrary)
            {
                if (!string.IsNullOrEmpty(customer.RunsFolder) &&
                    Directory.Exists(customer.RunsFolder) &&
                    !_watchTimers.ContainsKey(customer.Id))
                {
                    StartCustomerWatch(customer, silent: true);
                }
            }
            RefreshWatchAllBtn();
            // Update UI for current customer if they're in the library
            string name = TrendsCustomerNameBox.Text.Trim();
            var cur = _trendsLibrary.FirstOrDefault(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (cur != null && _watchTimers.ContainsKey(cur.Id))
            {
                TrendsWatchToggleBtn.Content    = "⏹ Stop Watch";
                TrendsWatchToggleBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1F2D20"));
                TrendsWatchToggleBtn.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
            }
        }

        private void WatchAllBtn_Click(object sender, RoutedEventArgs e) => WatchAll();

        private void StopAllBtn_Click(object sender, RoutedEventArgs e) => StopAll();

        private void RefreshWatchAllBtn()
        {
            int watching = _watchTimers.Count;
            int total    = _trendsLibrary.Count(c =>
                !string.IsNullOrEmpty(c.RunsFolder) && Directory.Exists(c.RunsFolder));
            int notWatching = total - watching;
            bool anyActive  = watching > 0;
            bool allActive  = total > 0 && watching >= total;

            if (TrendsWatchAllBtn == null || TrendsStopAllBtn == null) return;

            // ── Watch All button ────────────────────────────────────────────────
            // Enabled only when there are customers not yet being watched.
            // Badge shows how many are NOT watching when partially active.
            bool watchAllEnabled = notWatching > 0;
            TrendsWatchAllBtn.IsEnabled  = watchAllEnabled;
            TrendsWatchAllBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                watchAllEnabled ? "#1E2640" : "#111827"));
            TrendsWatchAllBtn.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                watchAllEnabled ? "#60A5FA" : "#374151"));

            if (TrendsWatchAllBadge != null)
            {
                // Show badge only in partial state — tells user how many will be started
                if (anyActive && !allActive)
                {
                    TrendsWatchAllBadge.Text       = $"(+{notWatching})";
                    TrendsWatchAllBadge.Visibility = Visibility.Visible;
                }
                else
                {
                    TrendsWatchAllBadge.Visibility = Visibility.Collapsed;
                }
            }

            // ── Stop All button ─────────────────────────────────────────────────
            // Enabled only when any watcher is running.
            // Badge always shows how many are currently active.
            TrendsStopAllBtn.IsEnabled  = anyActive;
            TrendsStopAllBtn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                anyActive ? "#1F2D20" : "#111827"));
            TrendsStopAllBtn.Foreground = anyActive
                ? new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80))
                : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#374151"));

            if (TrendsStopAllBadge != null)
            {
                if (anyActive)
                {
                    TrendsStopAllBadge.Text       = $"({watching}/{total})";
                    TrendsStopAllBadge.Visibility = Visibility.Visible;
                }
                else
                {
                    TrendsStopAllBadge.Visibility = Visibility.Collapsed;
                }
            }
        }

        private bool IsCurrentCustomer(TrendsCustomer c)
            => c.Name.Equals(TrendsCustomerNameBox.Text.Trim(), StringComparison.OrdinalIgnoreCase);

        private void CustomerWatchTick(TrendsCustomer customer)
        {
            if (_generatingIds.Contains(customer.Id)) return;
            if (!Directory.Exists(customer.RunsFolder))
            {
                StopCustomerWatch(customer.Id);
                return;
            }

            int intervalSecs = customer.WatchIntervalSecs > 0
                ? customer.WatchIntervalSecs
                : (_settings.TrendsWatchIntervalSecs > 0 ? _settings.TrendsWatchIntervalSecs : 60);

            bool isCurrent = IsCurrentCustomer(customer);
            if (isCurrent)
                TrendsWatchStatusLabel.Text =
                    $"Watching · every {intervalSecs}s · last scan: {DateTime.Now:HH:mm:ss}";

            var (changed, changeDesc) = TrendsManifest.HasChanged(customer.RunsFolder, customer.Name);
            if (!changed) return;

            _generatingIds.Add(customer.Id);
            if (isCurrent)
            {
                TrendsWatchStatusLabel.Text = $"{changeDesc} — regenerating…";
                TrendsWatchStatusLabel.Visibility = Visibility.Visible;
            }

            string runsFolder = customer.RunsFolder;
            string name       = customer.Name;
            string reportsFolder = string.IsNullOrEmpty(customer.ReportsFolder)
                ? runsFolder : customer.ReportsFolder;
            // Use per-customer settings if set, otherwise fall back to global/defaults
            int fw = customer.FailWindow > 0
                ? customer.FailWindow
                : (int.TryParse(_settings.TrendsFailWindow, out int gfw) && gfw >= 1 ? gfw : 3);
            int   watchMaxMonths    = customer.MaxMonths;
            bool  watchIncludeOld   = customer.IncludeOldRuns;

            // Run generation for this customer in background
            System.Threading.Tasks.Task.Run(() =>
            {
                Action<string>? log = isCurrent
                    ? msg => Dispatcher.Invoke(() =>
                    {
                        TrendsAppendLog(msg);
                    })
                    : null; // background customers: no UI log, no race condition

                var (ok, outputPath, error) =
                    TestRunTrendsProcessor.Generate(log, runsFolder, name, reportsFolder, fw, watchMaxMonths, watchIncludeOld);

                Dispatcher.Invoke(() =>
                {
                    if (ok)
                    {
                        TrendsManifest.Write(runsFolder, name);
                        var entry = _trendsLibrary.FirstOrDefault(c => c.Id == customer.Id);
                        if (entry != null)
                        { entry.LastGenerated = DateTime.Now; entry.LastOutput = outputPath; }
                        SaveTrendsLibrary();
                        RefreshTrendsLibraryUI();

                        if (isCurrent)
                        {
                            TrendsProgress.Visibility = Visibility.Collapsed;
                            TrendsRunBtn.IsEnabled    = true;
                            string shortName = System.IO.Path.GetFileName(outputPath);
                            TrendsStatusLabel.Text = $"Auto-updated {DateTime.Now:HH:mm:ss} → {shortName}";
                            TrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
                        }

                        // Show balloon tip with pass% and fail count
                        if (!isCurrent || !IsVisible)
                        {
                            _tray?.UpdateTooltip(_watchTimers.Count);
                            _tray?.ShowWatchResult(name, outputPath);
                        }
                    }
                    else
                    {
                        if (isCurrent)
                        {
                            TrendsStatusLabel.Text = $"Failed: {error}";
                            TrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71));
                            TrendsWatchStatusLabel.Text =
                                $"File busy — will retry in {intervalSecs}s";
                        }
                    }

                    _generatingIds.Remove(customer.Id);
                });
            });
        }

        // ── Settings — persist UI defaults ───────────────────────────────────

        private AppDataManager.AppSettings _settings = new();

        // ── Settings page handlers ────────────────────────────────────────────

        private void AppLogSettings_Changed(object sender, System.Windows.RoutedEventArgs e)
        {
            _settings.AppLogEnabled = AppLogEnabledCheck?.IsChecked == true;
            _settings.AppLogFolder  = AppLogFolderBox?.Text?.Trim() ?? "";
            AppDataManager.SaveSettings(_settings);
            AppLogger.Configure(_settings.AppLogEnabled, _settings.AppLogFolder);

            if (AppLogStatusLabel == null) return;
            if (_settings.AppLogEnabled && !string.IsNullOrEmpty(_settings.AppLogFolder))
            {
                AppLogStatusLabel.Text       = $"Logging enabled → {_settings.AppLogFolder}";
                AppLogStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
            }
            else if (_settings.AppLogEnabled && string.IsNullOrEmpty(_settings.AppLogFolder))
            {
                AppLogStatusLabel.Text       = "Select a log folder to activate logging.";
                AppLogStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0xFB, 0xBF, 0x24));
            }
            else
            {
                AppLogStatusLabel.Text       = "Logging disabled.";
                AppLogStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x7A, 0x99));
            }
        }

        private void AppLogFolderBrowse_Click(object sender, RoutedEventArgs e)
        {
            string folder = BrowseFolder("Select folder for log files");
            if (string.IsNullOrEmpty(folder)) return;
            if (AppLogFolderBox != null) AppLogFolderBox.Text = folder;
            AppLogSettings_Changed(sender, e);
        }

        private void LoadSettings()
        {
            _settings = AppDataManager.LoadSettings();

            // Apply to UI controls
            if (ClubOutputCheckbox     != null) ClubOutputCheckbox.IsChecked     = _settings.ConvertClubOutput;
            if (IncludeChartsCheckbox  != null) IncludeChartsCheckbox.IsChecked  = _settings.ConvertIncludeCharts;
            if (JTLClubOutputCheckbox  != null) JTLClubOutputCheckbox.IsChecked  = _settings.JtlClubOutput;
            if (JTLIncludeChartsCheckbox != null) JTLIncludeChartsCheckbox.IsChecked = _settings.JtlIncludeCharts;
            if (BLGProduceGraphsCheckbox != null) BLGProduceGraphsCheckbox.IsChecked = _settings.BlgProduceGraphs;
            if (BLGRadioDb != null && _settings.BlgServerType == "Db") BLGRadioDb.IsChecked = true;
            if (CmpModeSequential != null && _settings.CmpMode == "Sequential") CmpModeSequential.IsChecked = true;
            if (CmpSlaTextBox  != null) CmpSlaTextBox.Text  = _settings.CmpSlaMs;
            if (TrendsFailWindowBox != null) TrendsFailWindowBox.Text = _settings.TrendsFailWindow;
            if (NmonOutDirBox != null && !string.IsNullOrEmpty(_settings.LastNmonOutputDir))
                NmonOutDirBox.Text = _settings.LastNmonOutputDir;

            if (AppLogEnabledCheck != null) AppLogEnabledCheck.IsChecked = _settings.AppLogEnabled;
            if (AppLogFolderBox   != null) AppLogFolderBox.Text          = _settings.AppLogFolder;
            AppLogger.Configure(_settings.AppLogEnabled, _settings.AppLogFolder);

            // Restore auto-start checkbox; also sync with Task Scheduler reality
            // in case the task was manually deleted outside the app.
            if (_settings.AutoStartOnLogon && !WindowsTaskScheduler.IsAppAutoStartRegistered())
                _settings.AutoStartOnLogon = false;  // task was removed externally
            if (AutoStartOnLogonCheck != null)
                AutoStartOnLogonCheck.IsChecked = _settings.AutoStartOnLogon;

            // Auto-watch restore is handled by InitTrayOnLoad() after library is loaded
        }

        private void SaveSettings()
        {
            _settings.ConvertClubOutput    = ClubOutputCheckbox?.IsChecked      == true;
            _settings.ConvertIncludeCharts = IncludeChartsCheckbox?.IsChecked   == true;
            _settings.JtlClubOutput        = JTLClubOutputCheckbox?.IsChecked   == true;
            _settings.JtlIncludeCharts     = JTLIncludeChartsCheckbox?.IsChecked == true;
            _settings.BlgProduceGraphs     = BLGProduceGraphsCheckbox?.IsChecked == true;
            _settings.BlgServerType        = BLGRadioDb?.IsChecked == true ? "Db" : "App";
            _settings.CmpMode              = CmpModeSequential?.IsChecked == true ? "Sequential" : "AllVsBaseline";
            _settings.CmpSlaMs             = CmpSlaTextBox?.Text?.Trim() ?? "";
            _settings.TrendsFailWindow     = TrendsFailWindowBox?.Text?.Trim() ?? "3";
            _settings.LastNmonOutputDir    = NmonOutDirBox?.Text?.Trim() ?? "";
            _settings.AppLogEnabled        = AppLogEnabledCheck?.IsChecked == true;
            _settings.AppLogFolder         = AppLogFolderBox?.Text?.Trim() ?? "";
            // Note: AutoStartOnLogon is managed by EnsureAutoStartRegistered/Unregistered,
            // not saved here — it changes only when watchers start or stop.
            AppDataManager.SaveSettings(_settings);
            AppLogger.Configure(_settings.AppLogEnabled, _settings.AppLogFolder);
        }

        // Save settings whenever the window closes
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (_watchTimers.Count > 0)
            {
                e.Cancel = true;
                _tray?.MinimizeToTray();
                return;
            }
            base.OnClosing(e);
        }

        protected override void OnClosed(EventArgs e)
        {
            foreach (var timer in _watchTimers.Values) timer.Stop();
            _watchTimers.Clear();
            _tray?.Dispose();
            SaveSettings();
            AppLogger.Shutdown();     // flush and stop background log writer
            CleanOrphanTempFiles();   // remove any TrendRun_* leftovers in %TEMP%
            base.OnClosed(e);
        }

        /// <summary>
        /// Deletes orphaned TrendRun_* temp files left in %TEMP% by any
        /// ParseRunFile call that was interrupted before it could delete its copy.
        /// Safe to call at startup and shutdown — ignores any file that is still
        /// in use.
        /// </summary>
        private static void CleanOrphanTempFiles()
        {
            try
            {
                string tmp = Path.GetTempPath();
                foreach (var f in Directory.GetFiles(tmp, "TrendRun_*.xlsx"))
                {
                    try { File.Delete(f); } catch { /* in use or locked — skip */ }
                }
            }
            catch { }
        }

        // ── Auto-start on Windows logon (server resilience) ─────────────────
        //
        // Auto-start is managed automatically by the watcher lifecycle:
        //   • First watcher starts  → register Task Scheduler ONLOGON entry
        //   • All watchers stop     → unregister the entry
        // The Settings checkbox is read-only — it reflects the current state.

        /// <summary>
        /// Registers the Task Scheduler auto-start entry if not already present.
        /// Called when the first watcher starts.
        /// </summary>
        private void EnsureAutoStartRegistered()
        {
            if (WindowsTaskScheduler.IsAppAutoStartRegistered()) 
            {
                SyncAutoStartCheckbox(true);
                return;
            }

            var (ok, err) = WindowsTaskScheduler.RegisterAppAutoStart();
            if (ok)
            {
                _settings.AutoStartOnLogon = true;
                AppDataManager.SaveSettings(_settings);
                SyncAutoStartCheckbox(true);
                AppLogger.Write("AutoStart", "Registered auto-start on Windows logon.");
            }
            else
            {
                AppLogger.Write("AutoStart", $"Could not register auto-start: {err}");
                SyncAutoStartCheckbox(false);
            }
        }

        /// <summary>
        /// Removes the Task Scheduler auto-start entry if present.
        /// Called when the last watcher stops.
        /// </summary>
        private void EnsureAutoStartUnregistered()
        {
            if (!WindowsTaskScheduler.IsAppAutoStartRegistered())
            {
                SyncAutoStartCheckbox(false);
                return;
            }

            var (ok, err) = WindowsTaskScheduler.UnregisterAppAutoStart();
            if (ok)
            {
                _settings.AutoStartOnLogon = false;
                AppDataManager.SaveSettings(_settings);
                SyncAutoStartCheckbox(false);
            }
            else
            {
                AppLogger.Write("AutoStart", $"Could not unregister auto-start: {err}");
            }
        }

        /// <summary>Updates the read-only Settings checkbox to reflect the current state.</summary>
        private void SyncAutoStartCheckbox(bool registered)
        {
            if (AutoStartOnLogonCheck != null)
                AutoStartOnLogonCheck.IsChecked = registered;
        }

        // ── Export / Import — Script Library ─────────────────────────────────

        private void LibraryExport_Click(object sender, RoutedEventArgs e)
        {
            if (_library.Count == 0)
            {
                DarkMessageBox.Show("The script library is empty — nothing to export.", "Export Library");
                return;
            }
            var dlg = new SaveFileDialog
            {
                Title = "Export Script Library",
                Filter = "JSON backup (*.json)|*.json",
                FileName = $"ScriptLibrary_backup_{DateTime.Now:yyyyMMdd}.json"
            };
            if (dlg.ShowDialog() != true) return;

            bool ok = AppDataManager.ExportScriptLibrary(_library, dlg.FileName);
            if (ok)
                DarkMessageBox.Show($"Exported {_library.Count} script(s) to:\n{dlg.FileName}", "Export Complete");
            else
                DarkMessageBox.Show("Export failed — could not write the file.", "Export Failed");
        }

        private void LibraryImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Import Script Library",
                Filter = "JSON backup (*.json)|*.json|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() != true) return;

            var (imported, error) = AppDataManager.ImportScriptLibrary(dlg.FileName, _library);
            if (imported == null)
            {
                DarkMessageBox.Show($"Import failed:\n{error}", "Import Failed");
                return;
            }
            if (imported.Count == 0)
            {
                DarkMessageBox.Show("No new entries found — all scripts in the file already exist in your library.", "Nothing to Import");
                return;
            }

            bool confirmed = DarkMessageBox.Confirm(
                $"Import {imported.Count} new script(s) into the library?\n\nExisting scripts will not be affected.",
                "Confirm Import");
            if (!confirmed) return;

            _library.AddRange(imported);
            SaveLibrary();
            RefreshLibraryUI();
            ScriptStatusLabel.Text = $"Imported {imported.Count} new script(s).";
            ScriptStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
        }

        // ── Export / Import — Trends Library ─────────────────────────────────

        private void TrendsLibraryExport_Click(object sender, RoutedEventArgs e)
        {
            if (_trendsLibrary.Count == 0)
            {
                DarkMessageBox.Show("The trends library is empty — nothing to export.", "Export Library");
                return;
            }
            var dlg = new SaveFileDialog
            {
                Title = "Export Trends Library",
                Filter = "JSON backup (*.json)|*.json",
                FileName = $"TrendsLibrary_backup_{DateTime.Now:yyyyMMdd}.json"
            };
            if (dlg.ShowDialog() != true) return;

            var dtos = _trendsLibrary.Select(c => new AppDataManager.TrendsCustomerDto
            {
                Id = c.Id, Name = c.Name, RunsFolder = c.RunsFolder,
                ReportsFolder = c.ReportsFolder, LastGenerated = c.LastGenerated, LastOutput = c.LastOutput,
                FailWindow = c.FailWindow,
                WatchIntervalSecs = c.WatchIntervalSecs,
                MaxMonths         = c.MaxMonths,
                IncludeOldRuns    = c.IncludeOldRuns,
            }).ToList();

            bool ok = AppDataManager.ExportTrendsLibrary(dtos, dlg.FileName);
            if (ok)
                DarkMessageBox.Show($"Exported {_trendsLibrary.Count} customer(s) to:\n{dlg.FileName}", "Export Complete");
            else
                DarkMessageBox.Show("Export failed — could not write the file.", "Export Failed");
        }

        private void TrendsLibraryImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Import Trends Library",
                Filter = "JSON backup (*.json)|*.json|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() != true) return;

            var existingDtos = _trendsLibrary.Select(c => new AppDataManager.TrendsCustomerDto
            {
                Id = c.Id, Name = c.Name, RunsFolder = c.RunsFolder,
                ReportsFolder = c.ReportsFolder, LastGenerated = c.LastGenerated, LastOutput = c.LastOutput,
                FailWindow = c.FailWindow,
                WatchIntervalSecs = c.WatchIntervalSecs,
            }).ToList();

            var (imported, error) = AppDataManager.ImportTrendsLibrary(dlg.FileName, existingDtos);
            if (imported == null)
            {
                DarkMessageBox.Show($"Import failed:\n{error}", "Import Failed");
                return;
            }
            if (imported.Count == 0)
            {
                DarkMessageBox.Show("No new customers found — all entries in the file already exist in your library.", "Nothing to Import");
                return;
            }

            bool confirmed = DarkMessageBox.Confirm(
                $"Import {imported.Count} new customer(s) into the trends library?\n\nExisting customers will not be affected.",
                "Confirm Import");
            if (!confirmed) return;

            foreach (var dto in imported)
                _trendsLibrary.Add(new TrendsCustomer
                {
                    Id = dto.Id, Name = dto.Name, RunsFolder = dto.RunsFolder,
                    ReportsFolder = dto.ReportsFolder, LastGenerated = dto.LastGenerated,
                    LastOutput = dto.LastOutput, FailWindow = dto.FailWindow,
                    WatchIntervalSecs = dto.WatchIntervalSecs,
                    MaxMonths         = dto.MaxMonths,
                    IncludeOldRuns    = dto.IncludeOldRuns,
                });

            SaveTrendsLibrary();
            RefreshTrendsLibraryUI();
            TrendsStatusLabel.Text = $"Imported {imported.Count} new customer(s).";
            TrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
        }

        // ── Bulk Import from root folder ─────────────────────────────────────

        private void TrendsBulkImport_Click(object sender, RoutedEventArgs e)
        {
            // Step 1: browse for root customers folder
            string rootFolder = BrowseFolder("Select root folder containing customer subfolders");
            if (string.IsNullOrEmpty(rootFolder) || !Directory.Exists(rootFolder)) return;

            var subfolders = Directory.GetDirectories(rootFolder, "*", SearchOption.TopDirectoryOnly)
                .OrderBy(d => d)
                .ToList();

            if (subfolders.Count == 0)
            {
                DarkMessageBox.Show("No subfolders found in the selected folder.", "Bulk Import");
                return;
            }

            // Skip folders already in library (by name)
            var existingNames = new HashSet<string>(
                _trendsLibrary.Select(c => c.Name), StringComparer.OrdinalIgnoreCase);

            var toAdd = subfolders
                .Where(d => !existingNames.Contains(Path.GetFileName(d)))
                .ToList();

            if (toAdd.Count == 0)
            {
                DarkMessageBox.Show(
                    "All subfolders in that location are already in the library.",
                    "Bulk Import — Nothing New");
                return;
            }

            // Step 2: show the bulk-import options dialog
            var dialog = new BulkImportOptionsDialog(rootFolder, toAdd)
            {
                Owner = this
            };
            if (dialog.ShowDialog() != true) return;

            string reportsMode   = dialog.ReportsMode;     // "same" | "shared" | "subfolder"
            string sharedReports = dialog.SharedReportsFolder;

            int added = 0;
            foreach (var folder in toAdd)
            {
                string name = Path.GetFileName(folder);

                string reportsFolder = reportsMode switch
                {
                    "same"      => folder,
                    "shared"    => sharedReports,
                    "subfolder" => Path.Combine(sharedReports, name),
                    _           => folder
                };

                // Ensure per-customer subfolder exists when that mode is chosen
                if (reportsMode == "subfolder" && !Directory.Exists(reportsFolder))
                {
                    try { Directory.CreateDirectory(reportsFolder); } catch { /* non-fatal */ }
                }

                _trendsLibrary.Add(new TrendsCustomer
                {
                    Name          = name,
                    RunsFolder    = folder,
                    ReportsFolder = reportsFolder,
                    // Bulk import: no per-customer value yet; defaults to 0 (use global)
                });
                added++;
            }

            SaveTrendsLibrary();
            RefreshTrendsLibraryUI();

            int skipped = subfolders.Count - toAdd.Count;
            string msg = $"Added {added} customer(s) to the library.";
            if (skipped > 0) msg += $"\n{skipped} already existed and were skipped.";
            TrendsStatusLabel.Text = msg;
            TrendsStatusLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
        }

        // ── Export single customer ───────────────────────────────────────────

        private void ExportSingleCustomer(TrendsCustomer customer)
        {
            var dlg = new SaveFileDialog
            {
                Title    = $"Export {customer.Name}",
                Filter   = "JSON (*.json)|*.json",
                FileName = $"{customer.Name}_TrendsConfig_{DateTime.Now:yyyyMMdd}.json"
            };
            if (dlg.ShowDialog() != true) return;

            var dto = new AppDataManager.TrendsCustomerDto
            {
                Id = customer.Id, Name = customer.Name,
                RunsFolder = customer.RunsFolder, ReportsFolder = customer.ReportsFolder,
                LastGenerated = customer.LastGenerated, LastOutput = customer.LastOutput,
                FailWindow = customer.FailWindow, WatchIntervalSecs = customer.WatchIntervalSecs,
            };
            bool ok = AppDataManager.ExportTrendsLibrary(new List<AppDataManager.TrendsCustomerDto> { dto }, dlg.FileName);
            if (ok)
                DarkMessageBox.Show($"Exported {customer.Name} to:\n{dlg.FileName}", "Export Complete");
            else
                DarkMessageBox.Show("Export failed — could not write the file.", "Export Failed");
        }

        // ── Global Export / Import ────────────────────────────────────────────

        private void GlobalExport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Export All — Backup",
                Filter = "JSON backup (*.json)|*.json",
                FileName = $"PTU_backup_{DateTime.Now:yyyyMMdd_HHmm}.json"
            };
            if (dlg.ShowDialog() != true) return;

            SaveSettings(); // ensure in-memory settings are flushed first

            var dtos = _trendsLibrary.Select(c => new AppDataManager.TrendsCustomerDto
            {
                Id = c.Id, Name = c.Name, RunsFolder = c.RunsFolder,
                ReportsFolder = c.ReportsFolder, LastGenerated = c.LastGenerated, LastOutput = c.LastOutput,
                FailWindow = c.FailWindow,
                WatchIntervalSecs = c.WatchIntervalSecs,
            }).ToList();

            bool ok = AppDataManager.ExportAll(_library, dtos, _settings, dlg.FileName);
            if (ok)
                DarkMessageBox.Show(
                    $"Backup saved to:\n{dlg.FileName}\n\n" +
                    $"Contains: {_library.Count} script(s), {_trendsLibrary.Count} customer(s), all UI settings.",
                    "Backup Complete");
            else
                DarkMessageBox.Show("Backup failed — could not write the file.", "Backup Failed");
        }

        private void GlobalImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Import All — Restore from Backup",
                Filter = "JSON backup (*.json)|*.json|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() != true) return;

            var existingDtos = _trendsLibrary.Select(c => new AppDataManager.TrendsCustomerDto
            {
                Id = c.Id, Name = c.Name, RunsFolder = c.RunsFolder,
                ReportsFolder = c.ReportsFolder, LastGenerated = c.LastGenerated, LastOutput = c.LastOutput,
                FailWindow = c.FailWindow,
                WatchIntervalSecs = c.WatchIntervalSecs,
            }).ToList();

            var result = AppDataManager.ImportAll(dlg.FileName, _library, existingDtos);

            if (result.Error != null)
            {
                DarkMessageBox.Show($"Restore failed:\n{result.Error}", "Restore Failed");
                return;
            }

            // Build a human-readable summary for the confirmation popup
            var summary = new System.Text.StringBuilder();
            summary.AppendLine($"Backup contains:");
            summary.AppendLine($"  • {result.TotalScripts} script(s)  →  {result.NewScripts.Count} new (rest already exist)");
            summary.AppendLine($"  • {result.TotalCustomers} customer(s)  →  {result.NewCustomers.Count} new (rest already exist)");
            summary.AppendLine($"  • UI settings");
            summary.AppendLine();
            summary.AppendLine("Settings will be replaced. New library entries will be merged in.");
            summary.AppendLine("Existing entries will NOT be removed or overwritten.");
            summary.AppendLine();
            summary.Append("Continue?");

            bool confirmed = DarkMessageBox.Confirm(summary.ToString(), "Confirm Restore");
            if (!confirmed) return;

            // Apply scripts
            if (result.NewScripts.Count > 0)
            {
                _library.AddRange(result.NewScripts);
                SaveLibrary();
                RefreshLibraryUI();
            }

            // Apply customers
            if (result.NewCustomers.Count > 0)
            {
                foreach (var dto in result.NewCustomers)
                    _trendsLibrary.Add(new TrendsCustomer
                    {
                        Id = dto.Id, Name = dto.Name, RunsFolder = dto.RunsFolder,
                        ReportsFolder = dto.ReportsFolder, LastGenerated = dto.LastGenerated,
                        LastOutput = dto.LastOutput, FailWindow = dto.FailWindow,
                    });
                SaveTrendsLibrary();
                RefreshTrendsLibraryUI();
            }

            // Apply settings
            if (result.Settings != null)
            {
                _settings = result.Settings;
                AppDataManager.SaveSettings(_settings);
                LoadSettings(); // re-apply to UI controls
            }

            DarkMessageBox.Show(
                $"Restore complete.\n\n" +
                $"  • {result.NewScripts.Count} new script(s) added\n" +
                $"  • {result.NewCustomers.Count} new customer(s) added\n" +
                $"  • UI settings restored",
                "Restore Complete");
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

        private const int LogMaxLines = 500;   // cap per log panel — prevents unbounded string growth

        private void LogMsg(TextBlock log, string message, string colorHex = "#8B93A5")
        {
            string ts   = DateTime.Now.ToString("HH:mm:ss");
            string line = $"[{ts}]  {message}";

            if (log.Text.Length > 0)
            {
                // Count existing lines; trim oldest ~10% when we hit the cap
                string current  = log.Text;
                int    newlines = 0;
                for (int i = 0; i < current.Length; i++)
                    if (current[i] == '\n') newlines++;

                if (newlines >= LogMaxLines)
                {
                    int drop  = LogMaxLines / 10;
                    int start = 0;
                    for (int i = 0; i < drop && start < current.Length; i++)
                    {
                        int nl = current.IndexOf('\n', start);
                        if (nl < 0) break;
                        start = nl + 1;
                    }
                    current  = "… (older entries trimmed) …\n" + current[start..];
                    log.Text = current;
                }

                log.Text += "\n" + line;
            }
            else
            {
                log.Text = line;
            }

            log.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));
        }

        private void LogSuccess(TextBlock log, string message) => LogMsg(log, message, "#4ADE80");
        private void LogError(TextBlock log, string message) => LogMsg(log, message, "#F87171");
        private void LogInfo(TextBlock log, string message) => LogMsg(log, message, "#60A5FA");
        /// <summary>
        /// Appends a timestamped line to TrendsLog, trimming oldest entries
        /// when the panel exceeds LogMaxLines to prevent unbounded growth.
        /// </summary>
        private void TrendsAppendLog(string msg)
        {
            string line = $"[{DateTime.Now:HH:mm:ss}]  {msg}";
            if (TrendsLog.Text.Length > 0)
            {
                string current = TrendsLog.Text;
                int newlines   = 0;
                for (int i = 0; i < current.Length; i++)
                    if (current[i] == '\n') newlines++;

                if (newlines >= LogMaxLines)
                {
                    int drop  = LogMaxLines / 10;
                    int start = 0;
                    for (int i = 0; i < drop && start < current.Length; i++)
                    {
                        int nl = current.IndexOf('\n', start);
                        if (nl < 0) break;
                        start = nl + 1;
                    }
                    TrendsLog.Text = "… (older entries trimmed) …\n" + current[start..];
                }

                TrendsLog.Text += "\n" + line;
            }
            else
            {
                TrendsLog.Text = line;
            }
        }


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

        // ── AI File Chat page ─────────────────────────────────────────────────

        private readonly AiChatEngine _aiEngine = new();
        private CancellationTokenSource? _aiCts;
        private bool _aiStreaming = false;

        // ── Provider / API key management ─────────────────────────────────────

        private AiProviderType GetSelectedProvider()
        {
            var item = AiProviderCombo?.SelectedItem as ComboBoxItem;
            string tag = item?.Tag?.ToString() ?? "Claude";
            return tag switch
            {
                "ChatGPT" => AiProviderType.ChatGPT,
                "Gemini"  => AiProviderType.Gemini,
                _         => AiProviderType.Claude
            };
        }

        private string GetApiKeyForProvider(AiProviderType type) => type switch
        {
            AiProviderType.Claude  => _settings.AiClaudeApiKey,
            AiProviderType.ChatGPT => _settings.AiChatGptApiKey,
            AiProviderType.Gemini  => _settings.AiGeminiApiKey,
            _ => ""
        };

        private void SetApiKeyForProvider(AiProviderType type, string key)
        {
            switch (type)
            {
                case AiProviderType.Claude:  _settings.AiClaudeApiKey  = key; break;
                case AiProviderType.ChatGPT: _settings.AiChatGptApiKey = key; break;
                case AiProviderType.Gemini:  _settings.AiGeminiApiKey  = key; break;
            }
            AppDataManager.SaveSettings(_settings);
        }

        private void AiProvider_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (AiApiKeyBox == null) return;
            var type = GetSelectedProvider();
            string key = GetApiKeyForProvider(type);
            AiApiKeyBox.Text = string.IsNullOrEmpty(key) ? "" : MaskKey(key);
            AiApiKeyBox.Tag = key;  // store real key in Tag
            _settings.AiProvider = type.ToString();
            AppDataManager.SaveSettings(_settings);
        }

        private void AiApiKey_LostFocus(object sender, RoutedEventArgs e)
        {
            string text = AiApiKeyBox.Text.Trim();
            // If the user typed a new key (not the masked version), save it
            if (!string.IsNullOrEmpty(text) && !text.Contains("••••"))
            {
                var type = GetSelectedProvider();
                SetApiKeyForProvider(type, text);
                AiApiKeyBox.Tag = text;
                AiApiKeyBox.Text = MaskKey(text);
            }
        }

        private static string MaskKey(string key)
        {
            if (string.IsNullOrEmpty(key)) return "";
            if (key.Length <= 8) return "••••••••";
            return key[..4] + "••••••••" + key[^4..];
        }

        private void InitAiChat()
        {
            // Restore provider selection
            if (AiProviderCombo != null)
            {
                string saved = _settings.AiProvider ?? "Claude";
                foreach (ComboBoxItem item in AiProviderCombo.Items)
                {
                    if (item.Tag?.ToString() == saved)
                    { item.IsSelected = true; break; }
                }
            }

            // Restore API key display
            var provType = GetSelectedProvider();
            string apiKey = GetApiKeyForProvider(provType);
            if (AiApiKeyBox != null)
            {
                AiApiKeyBox.Text = string.IsNullOrEmpty(apiKey) ? "" : MaskKey(apiKey);
                AiApiKeyBox.Tag  = apiKey;
            }
        }

        // ── File management ───────────────────────────────────────────────────

        private void AiAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title       = "Select files for AI analysis",
                Multiselect = true,
                Filter      = "All Supported|*.xlsx;*.xls;*.csv;*.tsv;*.jtl;*.txt;*.log;*.json;*.xml;*.md;*.yaml;*.yml;*.pdf;*.docx|All Files|*.*"
            };
            if (dlg.ShowDialog() != true) return;

            foreach (string path in dlg.FileNames)
            {
                try
                {
                    var loaded = _aiEngine.AddFile(path);
                    AddAiSystemMessage($"Loaded: {loaded.Summary}");
                }
                catch (Exception ex)
                {
                    AddAiSystemMessage($"Failed to load {Path.GetFileName(path)}: {ex.Message}");
                }
            }
            RefreshAiFileStatus();
        }

        private void AiClear_Click(object sender, RoutedEventArgs e)
        {
            _aiCts?.Cancel();
            _aiEngine.Reset();
            AiChatMessagesPanel.Children.Clear();
            RefreshAiFileStatus();
        }

        private void RefreshAiFileStatus()
        {
            if (AiFileStatusLabel == null) return;
            int count = _aiEngine.Files.Count;
            int chars  = _aiEngine.TotalFileChars;
            AiFileStatusLabel.Text = count == 0
                ? "No files loaded"
                : $"{count} file(s) loaded · {chars:N0} chars · {_aiEngine.TotalChunks} chunk(s)";
            AiFileStatusLabel.Foreground = new SolidColorBrush(
                count > 0 ? Color.FromRgb(0x4A, 0xDE, 0x80) : Color.FromRgb(0x4A, 0x5F, 0x88));
        }

        // ── Send message ──────────────────────────────────────────────────────

        private void AiInput_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter && !_aiStreaming)
            {
                e.Handled = true;
                AiSend_Click(sender, e);
            }
        }

        private async void AiSend_Click(object sender, RoutedEventArgs e)
        {
            string message = AiInputBox.Text.Trim();
            if (string.IsNullOrEmpty(message)) return;
            if (_aiStreaming) return;

            var providerType = GetSelectedProvider();
            string apiKey = AiApiKeyBox.Tag as string ?? GetApiKeyForProvider(providerType);

            if (string.IsNullOrEmpty(apiKey))
            {
                AddAiSystemMessage("Please enter your API key first.");
                return;
            }

            // Add user bubble
            AddAiUserMessage(message);
            AiInputBox.Text = "";

            // Create assistant bubble (will be streamed into)
            var assistantBubble = AddAiAssistantMessage("");

            _aiStreaming = true;
            AiSendBtn.Content = "Stop";
            AiSendBtn.Background = new SolidColorBrush(Color.FromRgb(0xDC, 0x26, 0x26));
            AiInputBox.IsEnabled = false;
            _aiCts = new CancellationTokenSource();

            try
            {
                var provider = AiProviderFactory.Create(providerType, apiKey);

                await _aiEngine.SendAsync(
                    provider,
                    message,
                    token => Dispatcher.Invoke(() =>
                    {
                        AppendToAssistantBubble(assistantBubble, token);
                    }),
                    _aiCts.Token);
            }
            catch (OperationCanceledException)
            {
                AppendToAssistantBubble(assistantBubble, "\n\n[Stopped]");
            }
            catch (HttpRequestException ex)
            {
                string errMsg = ex.Message;
                // Try to extract a readable error from the API response
                if (errMsg.Contains("401") || errMsg.Contains("403"))
                    errMsg = "Invalid API key. Please check your key and try again.";
                else if (errMsg.Contains("429"))
                    errMsg = "Rate limited. Please wait a moment and try again.";
                else if (errMsg.Contains("500") || errMsg.Contains("502") || errMsg.Contains("503"))
                    errMsg = "The AI service is temporarily unavailable. Try again shortly.";

                AppendToAssistantBubble(assistantBubble, $"\n\n[Error: {errMsg}]");
            }
            catch (Exception ex)
            {
                AppendToAssistantBubble(assistantBubble, $"\n\n[Error: {ex.Message}]");
            }
            finally
            {
                _aiStreaming = false;
                AiSendBtn.Content = "Send";
                AiSendBtn.Background = new SolidColorBrush(Color.FromRgb(0x37, 0x63, 0xFF));
                AiInputBox.IsEnabled = true;
                AiInputBox.Focus();
            }
        }

        // ── Chat bubble builders ──────────────────────────────────────────────

        private static readonly SolidColorBrush UserBubbleBg   = new(Color.FromRgb(0x1E, 0x3A, 0x8A));
        private static readonly SolidColorBrush AssistBubbleBg = new(Color.FromRgb(0x14, 0x18, 0x2B));
        private static readonly SolidColorBrush SystemBubbleBg = new(Color.FromRgb(0x1A, 0x1A, 0x2E));
        private static readonly SolidColorBrush UserFg         = new(Color.FromRgb(0xBF, 0xDB, 0xFE));
        private static readonly SolidColorBrush AssistFg       = new(Color.FromRgb(0xCB, 0xD5, 0xE1));
        private static readonly SolidColorBrush SystemFg       = new(Color.FromRgb(0x6B, 0x7A, 0x99));

        private void AddAiUserMessage(string text)
        {
            var border = new Border
            {
                Background    = UserBubbleBg,
                CornerRadius  = new CornerRadius(12, 12, 2, 12),
                Padding       = new Thickness(14, 10, 14, 10),
                Margin        = new Thickness(80, 4, 0, 4),
                HorizontalAlignment = HorizontalAlignment.Right,
                MaxWidth      = 600,
            };
            border.Child = new TextBlock
            {
                Text         = text,
                Foreground   = UserFg,
                FontSize     = 13,
                FontFamily   = new FontFamily("Segoe UI Variable, Segoe UI"),
                TextWrapping = TextWrapping.Wrap,
            };
            AiChatMessagesPanel.Children.Add(border);
            AiChatScroll.ScrollToEnd();
        }

        private Border AddAiAssistantMessage(string text)
        {
            var border = new Border
            {
                Background    = AssistBubbleBg,
                CornerRadius  = new CornerRadius(12, 12, 12, 2),
                Padding       = new Thickness(14, 10, 14, 10),
                Margin        = new Thickness(0, 4, 80, 4),
                HorizontalAlignment = HorizontalAlignment.Left,
                MaxWidth      = 700,
                BorderBrush   = new SolidColorBrush(Color.FromRgb(0x1E, 0x25, 0x40)),
                BorderThickness = new Thickness(1),
            };
            border.Child = new TextBlock
            {
                Text         = text,
                Foreground   = AssistFg,
                FontSize     = 13,
                FontFamily   = new FontFamily("Segoe UI Variable, Segoe UI"),
                TextWrapping = TextWrapping.Wrap,
            };
            AiChatMessagesPanel.Children.Add(border);
            AiChatScroll.ScrollToEnd();
            return border;
        }

        private void AddAiSystemMessage(string text)
        {
            var border = new Border
            {
                Background    = SystemBubbleBg,
                CornerRadius  = new CornerRadius(8),
                Padding       = new Thickness(12, 6, 12, 6),
                Margin        = new Thickness(40, 2, 40, 2),
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            border.Child = new TextBlock
            {
                Text         = text,
                Foreground   = SystemFg,
                FontSize     = 11,
                FontFamily   = new FontFamily("Segoe UI Variable, Segoe UI"),
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center,
            };
            AiChatMessagesPanel.Children.Add(border);
            AiChatScroll.ScrollToEnd();
        }

        private void AppendToAssistantBubble(Border bubble, string token)
        {
            if (bubble.Child is TextBlock tb)
            {
                tb.Text += token;
                AiChatScroll.ScrollToEnd();
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
                DarkMessageBox.Show("Please select a CSV or JTL file.",
                    "Wrong File Type");
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
                DarkMessageBox.Show(
                    "Please select at least two files — the first is the baseline.",
                    "Not Enough Files");
                return;
            }

            var missing = paths.Where(f => !File.Exists(f)).ToList();
            if (missing.Count > 0)
            {
                DarkMessageBox.Show(
                    $"These files no longer exist:\n\n{string.Join("\n", missing.Select(Path.GetFileName))}",
                    "Missing Files");
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
                    DarkMessageBox.Show("SLA threshold must be a positive number (milliseconds).",
                        "Invalid Input");
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

    // ── Library Save Dialog ───────────────────────────────────────────────────

    public class LibrarySaveDialog : System.Windows.Window
    {
        public string EntryName        { get; private set; } = "";
        public string EntryDescription { get; private set; } = "";
        public string SuggestedName    { get; set; } = "";

        private System.Windows.Controls.TextBox _nameBox = new();
        private System.Windows.Controls.TextBox _descBox = new();

        public LibrarySaveDialog()
        {
            Title           = "Save to Library";
            Width           = 400;
            Height          = 260;
            SizeToContent   = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode      = ResizeMode.NoResize;
            Background      = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#0F1117"));

            var grid = new Grid { Margin = new Thickness(20) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            void AddLabel(int row, string text)
            {
                var lbl = new System.Windows.Controls.TextBlock
                {
                    Text = text, FontSize = 10, FontWeight = FontWeights.Bold,
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#6B7A99")),
                    Margin = new Thickness(0, row == 0 ? 0 : 10, 0, 4)
                };
                Grid.SetRow(lbl, row);
                grid.Children.Add(lbl);
            }

            System.Windows.Controls.TextBox MakeBox(int row)
            {
                var box = new System.Windows.Controls.TextBox
                {
                    Background = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#161B2A")),
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E2E8F0")),
                    BorderBrush = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#252D42")),
                    CaretBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.White),
                    Height = 30, Padding = new Thickness(8, 0, 8, 0),
                    VerticalContentAlignment = VerticalAlignment.Center, FontSize = 12
                };
                Grid.SetRow(box, row);
                grid.Children.Add(box);
                return box;
            }

            AddLabel(0, "NAME");
            _nameBox = MakeBox(1);

            AddLabel(2, "DESCRIPTION  (optional)");
            _descBox = MakeBox(3);

            // Buttons
            var btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            var saveBtn = new System.Windows.Controls.Button
            {
                Content = "Save", Width = 80, Height = 32, Margin = new Thickness(0, 0, 8, 0),
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#2563EB")),
                Foreground = System.Windows.Media.Brushes.White,
                FontWeight = FontWeights.SemiBold, FontSize = 12,
                BorderThickness = new Thickness(0)
            };
            saveBtn.Click += (_, _) =>
            {
                if (string.IsNullOrWhiteSpace(_nameBox.Text))
                {
                    _nameBox.BorderBrush = System.Windows.Media.Brushes.Red;
                    return;
                }
                EntryName        = _nameBox.Text.Trim();
                EntryDescription = _descBox.Text.Trim();
                DialogResult = true;
            };
            var cancelBtn = new System.Windows.Controls.Button
            {
                Content = "Cancel", Width = 70, Height = 32,
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#374151")),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#9CA3AF")),
                FontSize = 12, BorderThickness = new Thickness(0)
            };
            cancelBtn.Click += (_, _) => DialogResult = false;
            btnRow.Children.Add(saveBtn);
            btnRow.Children.Add(cancelBtn);
            Grid.SetRow(btnRow, 4);
            grid.Children.Add(btnRow);

            Content = grid;
            Loaded += (_, _) =>
            {
                _nameBox.Text = SuggestedName;
                _nameBox.SelectAll();
                _nameBox.Focus();
            };
        }
    }

    // ── Schedule Dialog ───────────────────────────────────────────────────────

    public class ScheduleDialog : System.Windows.Window
    {
        public MainWindow.ScriptSchedule? Result { get; private set; }

        private System.Windows.Controls.ComboBox   _typeBox    = new();
        private System.Windows.Controls.TextBox    _timeBox    = new();
        private System.Windows.Controls.TextBox    _logPathBox = new();
        private System.Windows.Controls.ComboBox   _dayBox     = new();
        private System.Windows.Controls.DatePicker _datePicker = new();
        private System.Windows.Controls.CheckBox   _enabledBox = new();
        private System.Windows.Controls.StackPanel _oncePanel  = new();
        private System.Windows.Controls.StackPanel _weeklyPanel = new();

        public ScheduleDialog(MainWindow.ScriptSchedule? existing)
        {
            Title  = "Set Schedule";
            Width  = 400;
            SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            Background = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#0F1117"));

            var outer = new System.Windows.Controls.StackPanel { Margin = new Thickness(20) };

            System.Windows.Controls.TextBlock Lbl(string t) => new()
            {
                Text = t, FontSize = 10, FontWeight = FontWeights.Bold,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#6B7A99")),
                Margin = new Thickness(0, 10, 0, 4)
            };

            // Enabled toggle — prominent banner, hard to miss
            bool isEnabled = existing?.Enabled ?? true;

            var enabledBorder = new Border
            {
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(0, 0, 0, 4),
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(
                        isEnabled ? "#0D2A1A" : "#2A0D0D")),
                BorderThickness = new Thickness(1),
                BorderBrush = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(
                        isEnabled ? "#16A34A" : "#991B1B")),
                Cursor = System.Windows.Input.Cursors.Hand
            };

            _enabledBox = new System.Windows.Controls.CheckBox
            {
                IsChecked = isEnabled,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(
                        isEnabled ? "#4ADE80" : "#F87171")),
                FontSize = 12.5, FontWeight = FontWeights.SemiBold,
                Cursor = System.Windows.Input.Cursors.Hand,
                Content = isEnabled ? "✓  Schedule enabled — task will run at the set time" : "✗  Schedule disabled — task will not run"
            };

            // Update border + label colours on toggle
            _enabledBox.Checked += (_, _) =>
            {
                enabledBorder.Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#0D2A1A"));
                enabledBorder.BorderBrush = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#16A34A"));
                _enabledBox.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#4ADE80"));
                _enabledBox.Content = "✓  Schedule enabled — task will run at the set time";
            };
            _enabledBox.Unchecked += (_, _) =>
            {
                enabledBorder.Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#2A0D0D"));
                enabledBorder.BorderBrush = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#991B1B"));
                _enabledBox.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F87171"));
                _enabledBox.Content = "✗  Schedule disabled — task will not run";
            };

            enabledBorder.Child = _enabledBox;
            enabledBorder.MouseLeftButtonUp += (_, _) => { _enabledBox.IsChecked = !_enabledBox.IsChecked; };
            outer.Children.Add(enabledBorder);

            // Type
            outer.Children.Add(Lbl("SCHEDULE TYPE"));
            _typeBox = new System.Windows.Controls.ComboBox
            {
                Height = 30, Foreground = System.Windows.Media.Brushes.Black
            };
            foreach (var t in new[] { "Once", "Daily", "Weekly" }) _typeBox.Items.Add(t);
            _typeBox.SelectedItem = existing?.Type ?? "Once";
            if (_typeBox.SelectedIndex < 0) _typeBox.SelectedIndex = 0;
            _typeBox.SelectionChanged += (_, _) => UpdateVisibility();
            outer.Children.Add(_typeBox);

            // Time of day — immersive row with +/- minute buttons
            outer.Children.Add(Lbl("TIME  (HH:mm, 24-hour)"));

            string defaultTime = existing?.TimeOfDay ?? DateTime.Now.ToString("HH:mm");
            _timeBox = new System.Windows.Controls.TextBox
            {
                Height = 36, Text = defaultTime,
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#161B2A")),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E2E8F0")),
                BorderBrush = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#252D42")),
                Padding = new Thickness(10, 0, 10, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                CaretBrush = System.Windows.Media.Brushes.White,
                FontFamily = new FontFamily("Consolas, Segoe UI Mono"),
                FontSize = 16, FontWeight = FontWeights.SemiBold,
                Width = 90, TextAlignment = TextAlignment.Center
            };

            System.Windows.Controls.Button MakeTimeBtn(string label, int minuteDelta)
            {
                var btn = new System.Windows.Controls.Button
                {
                    Content = label, Width = 36, Height = 36,
                    FontSize = 13, FontWeight = FontWeights.Bold,
                    Background = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1E2640")),
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#60A5FA")),
                    BorderThickness = new Thickness(1),
                    BorderBrush = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#252D42")),
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                btn.Click += (_, _) =>
                {
                    if (TimeSpan.TryParse(_timeBox.Text.Trim(), out var t))
                    {
                        var newTime = DateTime.Today.Add(t).AddMinutes(minuteDelta).TimeOfDay;
                        _timeBox.Text = newTime.ToString(@"hh\:mm");
                    }
                };
                return btn;
            }

            var timeRow = new System.Windows.Controls.StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 0)
            };
            timeRow.Children.Add(MakeTimeBtn("−5",  -5));
            timeRow.Children.Add(MakeTimeBtn("−1",  -1));
            timeRow.Children.Add(_timeBox);
            timeRow.Children.Add(MakeTimeBtn("+1",  +1));
            timeRow.Children.Add(MakeTimeBtn("+5",  +5));

            // "Now" shortcut
            var nowBtn = new System.Windows.Controls.Button
            {
                Content = "Now", Height = 36, Padding = new Thickness(10, 0, 10, 0),
                Margin = new Thickness(8, 0, 0, 0),
                FontSize = 11, FontWeight = FontWeights.SemiBold,
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#0D2140")),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#7DD3FC")),
                BorderThickness = new Thickness(1),
                BorderBrush = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1E3A6E")),
                Cursor = System.Windows.Input.Cursors.Hand,
                ToolTip = "Set to current time"
            };
            nowBtn.Click += (_, _) => _timeBox.Text = DateTime.Now.ToString("HH:mm");
            timeRow.Children.Add(nowBtn);

            outer.Children.Add(timeRow);

            // Past-time hint — shown inline below the time field when time is in the past
            var timeHint = new System.Windows.Controls.TextBlock
            {
                Text = "⚠  This time is in the past — the task will not run today.",
                FontSize = 10.5,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#FCD34D")),
                FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                Margin = new Thickness(0, 4, 0, 0),
                Visibility = Visibility.Collapsed,
                TextWrapping = TextWrapping.Wrap
            };
            outer.Children.Add(timeHint);

            // Check on every text change
            _timeBox.TextChanged += (_, _) =>
            {
                if (TimeSpan.TryParse(_timeBox.Text.Trim(), out var t))
                {
                    bool isPast = _typeBox.SelectedItem?.ToString() == "Once"
                        ? (_datePicker.SelectedDate ?? DateTime.Today) == DateTime.Today && t < DateTime.Now.TimeOfDay
                        : t < DateTime.Now.TimeOfDay;
                    timeHint.Visibility = isPast ? Visibility.Visible : Visibility.Collapsed;
                }
            };

            // Once: date — defaults to today
            _oncePanel = new System.Windows.Controls.StackPanel();
            _oncePanel.Children.Add(Lbl("DATE"));
            DateTime preDate = DateTime.Today;  // default today, not tomorrow
            if (existing?.RunOnce != null && DateTime.TryParse(existing.RunOnce, out var pd)) preDate = pd;
            _datePicker = new System.Windows.Controls.DatePicker
            {
                SelectedDate = preDate, Height = 30,
                DisplayDateStart = DateTime.Today  // prevent selecting past dates
            };
            _oncePanel.Children.Add(_datePicker);
            outer.Children.Add(_oncePanel);

            // Log file path
            outer.Children.Add(Lbl("LOG FILE PATH  (blank = script folder, appends each run)"));
            var logRow = new System.Windows.Controls.Grid();
            logRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            logRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            _logPathBox = new System.Windows.Controls.TextBox
            {
                Height = 30,
                Text = existing?.LogFilePath ?? "",
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#161B2A")),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#CBD5E1")),
                BorderBrush = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#252D42")),
                Padding = new Thickness(8, 0, 8, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                FontFamily = new FontFamily("Consolas, Segoe UI Mono"),
                FontSize = 11,
                CaretBrush = System.Windows.Media.Brushes.White
            };
            System.Windows.Controls.Grid.SetColumn(_logPathBox, 0);

            var browseLogBtn = new System.Windows.Controls.Button
            {
                Content = "…", Width = 30, Height = 30, Margin = new Thickness(6, 0, 0, 0),
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#374151")),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#B0BAC8")),
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand
            };
            browseLogBtn.Click += (_, _) =>
            {
                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    Title = "Choose log file location",
                    Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*",
                    DefaultExt = ".log",
                    FileName = "scheduled.log"
                };
                if (dlg.ShowDialog() == true)
                    _logPathBox.Text = dlg.FileName;
            };
            System.Windows.Controls.Grid.SetColumn(browseLogBtn, 1);

            logRow.Children.Add(_logPathBox);
            logRow.Children.Add(browseLogBtn);
            outer.Children.Add(logRow);

            // Weekly: day of week
            _weeklyPanel = new System.Windows.Controls.StackPanel { Visibility = Visibility.Collapsed };
            _weeklyPanel.Children.Add(Lbl("DAY OF WEEK"));
            _dayBox = new System.Windows.Controls.ComboBox
            {
                Height = 30, Foreground = System.Windows.Media.Brushes.Black
            };
            foreach (var d in new[] { "Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday" })
                _dayBox.Items.Add(d);
            _dayBox.SelectedIndex = existing?.DayOfWeek ?? 1;
            _weeklyPanel.Children.Add(_dayBox);
            outer.Children.Add(_weeklyPanel);

            // Buttons
            var btnRow = new System.Windows.Controls.StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 16, 0, 0)
            };

            var saveBtn = new System.Windows.Controls.Button
            {
                Content = "Save", Width = 80, Height = 32, Margin = new Thickness(0, 0, 8, 0),
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#2563EB")),
                Foreground = System.Windows.Media.Brushes.White,
                FontWeight = FontWeights.SemiBold, FontSize = 12, BorderThickness = new Thickness(0)
            };
            saveBtn.Click += (_, _) =>
            {
                if (!TimeSpan.TryParse(_timeBox.Text.Trim(), out var t))
                {
                    _timeBox.BorderBrush = System.Windows.Media.Brushes.Red; return;
                }
                var type = _typeBox.SelectedItem?.ToString() ?? "Once";
                Result = new MainWindow.ScriptSchedule
                {
                    Type      = type,
                    TimeOfDay = _timeBox.Text.Trim(),
                    DayOfWeek = _dayBox.SelectedIndex,
                    RunOnce   = type == "Once"
                        ? (_datePicker.SelectedDate ?? DateTime.Today).Date.Add(t).ToString("o")
                        : null,
                    Enabled     = _enabledBox.IsChecked == true,
                    LogFilePath = string.IsNullOrWhiteSpace(_logPathBox.Text) ? null : _logPathBox.Text.Trim()
                };
                DialogResult = true;
            };

            var clearBtn = new System.Windows.Controls.Button
            {
                Content = "Remove", Width = 80, Height = 32, Margin = new Thickness(0, 0, 8, 0),
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#3D1F1F")),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F87171")),
                FontSize = 12, BorderThickness = new Thickness(0)
            };
            clearBtn.Click += (_, _) => { Result = null; DialogResult = true; };

            var cancelBtn = new System.Windows.Controls.Button
            {
                Content = "Cancel", Width = 70, Height = 32,
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#374151")),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#9CA3AF")),
                FontSize = 12, BorderThickness = new Thickness(0)
            };
            cancelBtn.Click += (_, _) => DialogResult = false;

            btnRow.Children.Add(saveBtn);
            btnRow.Children.Add(clearBtn);
            btnRow.Children.Add(cancelBtn);
            outer.Children.Add(btnRow);
            Content = outer;

            UpdateVisibility();
        }

        private void UpdateVisibility()
        {
            string type = _typeBox.SelectedItem?.ToString() ?? "Once";
            _oncePanel.Visibility   = type == "Once"   ? Visibility.Visible : Visibility.Collapsed;
            _weeklyPanel.Visibility = type == "Weekly" ? Visibility.Visible : Visibility.Collapsed;
        }
    }
    // ── Dark Message Box ──────────────────────────────────────────────────────

    public static class DarkMessageBox
    {
        private static System.Windows.Window? _owner;
        public static void SetOwner(System.Windows.Window w) => _owner = w;

        public static void Show(string message, string title = "Performance Test Utilities")
        {
            var dlg = new DarkMsgDialog(title, message, false) { Owner = _owner };
            dlg.ShowDialog();
        }

        public static bool Confirm(string message, string title = "Confirm")
        {
            var dlg = new DarkMsgDialog(title, message, true) { Owner = _owner };
            dlg.ShowDialog();
            return dlg.Confirmed;
        }
    }

    public class DarkMsgDialog : System.Windows.Window
    {
        public bool Confirmed { get; private set; } = false;

        public DarkMsgDialog(string title, string message, bool isConfirm)
        {
            Title  = title;
            Width  = 420;
            SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            Background = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#0F1117"));
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1E2640"));
            BorderThickness = new Thickness(1);

            var root = new Grid { Margin = new Thickness(24, 20, 24, 20) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var titleBlock = new System.Windows.Controls.TextBlock
            {
                Text = title, FontSize = 14, FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E2E8F0")),
                FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(titleBlock, 0);
            root.Children.Add(titleBlock);

            var msgBlock = new System.Windows.Controls.TextBlock
            {
                Text = message, FontSize = 12.5,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#A8B3C8")),
                FontFamily = new FontFamily("Segoe UI Variable, Segoe UI"),
                TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 20)
            };
            Grid.SetRow(msgBlock, 1);
            root.Children.Add(msgBlock);

            var btnRow = new System.Windows.Controls.StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            if (isConfirm)
            {
                var yesBtn = MakeBtn("Yes", "#2563EB", "#FFFFFF");
                yesBtn.Click += (_, _) => { Confirmed = true; DialogResult = true; };
                var noBtn  = MakeBtn("No",  "#374151", "#9CA3AF");
                noBtn.Click += (_, _) => { Confirmed = false; DialogResult = false; };
                btnRow.Children.Add(yesBtn);
                btnRow.Children.Add(noBtn);
            }
            else
            {
                var okBtn = MakeBtn("OK", "#2563EB", "#FFFFFF");
                okBtn.Click += (_, _) => DialogResult = true;
                btnRow.Children.Add(okBtn);
            }

            Grid.SetRow(btnRow, 2);
            root.Children.Add(btnRow);
            Content = root;
        }

        private static System.Windows.Controls.Button MakeBtn(string text, string bg, string fg)
            => new()
            {
                Content = text, Width = 80, Height = 34,
                Margin = new Thickness(8, 0, 0, 0),
                FontSize = 12, FontWeight = FontWeights.SemiBold,
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand,
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(bg)),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(fg)),
                FontFamily = new FontFamily("Segoe UI Variable, Segoe UI")
            };
    }

}