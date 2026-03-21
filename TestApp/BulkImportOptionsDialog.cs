using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace TestApp
{
    /// <summary>
    /// Dark-themed dialog shown during Bulk Import.
    /// Lets the user choose how the Reports output folder is assigned
    /// across all newly discovered customer folders.
    /// </summary>
    internal class BulkImportOptionsDialog : Window
    {
        // ── Outputs read by the caller ────────────────────────────────────────
        /// <summary>"same" | "shared" | "subfolder"</summary>
        public string ReportsMode { get; private set; } = "same";
        public string SharedReportsFolder { get; private set; } = "";

        // ── UI refs ───────────────────────────────────────────────────────────
        private RadioButton _rbSame      = null!;
        private RadioButton _rbShared    = null!;
        private RadioButton _rbSubfolder = null!;
        private TextBox     _sharedBox   = null!;
        private Button      _browseBtn   = null!;
        private Button      _okBtn       = null!;
        private TextBlock   _errorLabel  = null!;

        // ── Colours (matching app palette) ───────────────────────────────────
        private static readonly Color BgDeep   = Color.FromRgb(0x0A, 0x0D, 0x18);
        private static readonly Color BgPanel  = Color.FromRgb(0x10, 0x14, 0x28);
        private static readonly Color BgInput  = Color.FromRgb(0x1A, 0x20, 0x38);
        private static readonly Color Border   = Color.FromRgb(0x2A, 0x33, 0x55);
        private static readonly Color FgPrimary= Color.FromRgb(0xE2, 0xE8, 0xFF);
        private static readonly Color FgMuted  = Color.FromRgb(0x6B, 0x7A, 0x99);
        private static readonly Color Accent   = Color.FromRgb(0xA7, 0x8B, 0xFA);  // violet for bulk-import
        private static readonly Color AccentGo = Color.FromRgb(0x37, 0x63, 0xFF);

        public BulkImportOptionsDialog(string rootFolder, List<string> foldersToAdd)
        {
            Title           = "Bulk Import — Reports Output";
            Width           = 520;
            SizeToContent   = SizeToContent.Height;
            ResizeMode      = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background      = new SolidColorBrush(BgDeep);
            Foreground      = new SolidColorBrush(FgPrimary);
            FontFamily      = new FontFamily("Segoe UI Variable, Segoe UI, sans-serif");
            FontSize        = 13;

            Content = BuildLayout(rootFolder, foldersToAdd);
        }

        private UIElement BuildLayout(string rootFolder, List<string> foldersToAdd)
        {
            var root = new StackPanel { Margin = new Thickness(24, 20, 24, 20) };

            // ── Header ────────────────────────────────────────────────────────
            root.Children.Add(new TextBlock
            {
                Text       = "Bulk Import",
                FontSize   = 17,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Accent),
                Margin     = new Thickness(0, 0, 0, 4)
            });

            root.Children.Add(new TextBlock
            {
                Text         = $"Found {foldersToAdd.Count} new customer folder(s) in:",
                Foreground   = new SolidColorBrush(FgMuted),
                FontSize     = 11.5,
                TextWrapping = TextWrapping.Wrap,
                Margin       = new Thickness(0, 0, 0, 2)
            });

            root.Children.Add(new TextBlock
            {
                Text         = rootFolder,
                Foreground   = new SolidColorBrush(FgPrimary),
                FontSize     = 11,
                TextWrapping = TextWrapping.Wrap,
                Margin       = new Thickness(0, 0, 0, 16)
            });

            // ── Section label ────────────────────────────────────────────────
            root.Children.Add(new TextBlock
            {
                Text       = "Where should trend reports be saved?",
                FontWeight = FontWeights.SemiBold,
                Margin     = new Thickness(0, 0, 0, 10)
            });

            // ── Radio: Same as Runs ───────────────────────────────────────────
            _rbSame = MakeRadio(
                "Same as Runs folder  (report saved alongside run files)",
                isChecked: true);
            root.Children.Add(_rbSame);

            // ── Radio: One shared folder ──────────────────────────────────────
            _rbShared = MakeRadio("One shared folder for all customers");
            root.Children.Add(_rbShared);

            // ── Radio: Per-customer subfolder ────────────────────────────────
            _rbSubfolder = MakeRadio(
                "Per-customer subfolder inside a base path  (base\\CustomerName\\)");
            root.Children.Add(_rbSubfolder);

            // ── Shared/base folder picker (shown for last two options) ────────
            var pickerPanel = new Grid { Margin = new Thickness(22, 8, 0, 4) };
            pickerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            pickerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            _sharedBox = new TextBox
            {
                IsReadOnly        = true,
                Background        = new SolidColorBrush(BgInput),
                Foreground        = new SolidColorBrush(FgPrimary),
                BorderBrush       = new SolidColorBrush(Border),
                BorderThickness   = new Thickness(1),
                Padding           = new Thickness(8, 5, 8, 5),
                FontSize          = 11.5,
                VerticalContentAlignment = VerticalAlignment.Center,
                Text              = ""
            };
            Grid.SetColumn(_sharedBox, 0);

            _browseBtn = new Button
            {
                Content         = "…",
                Width           = 32,
                Margin          = new Thickness(6, 0, 0, 0),
                Background      = new SolidColorBrush(Color.FromRgb(0x1E, 0x26, 0x40)),
                Foreground      = new SolidColorBrush(FgPrimary),
                BorderBrush     = new SolidColorBrush(Border),
                BorderThickness = new Thickness(1),
                Cursor          = System.Windows.Input.Cursors.Hand
            };
            _browseBtn.Click += BrowseShared_Click;
            Grid.SetColumn(_browseBtn, 1);

            pickerPanel.Children.Add(_sharedBox);
            pickerPanel.Children.Add(_browseBtn);

            // Hint label under picker
            var pickerHint = new TextBlock
            {
                FontSize   = 10.5,
                Foreground = new SolidColorBrush(FgMuted),
                Margin     = new Thickness(22, 2, 0, 0),
                Text       = "Browse to select the output folder."
            };

            root.Children.Add(pickerPanel);
            root.Children.Add(pickerHint);

            // Show/hide picker based on radio selection
            void UpdatePickerVisibility()
            {
                bool needPicker = _rbShared.IsChecked == true || _rbSubfolder.IsChecked == true;
                pickerPanel.Visibility = needPicker ? Visibility.Visible : Visibility.Collapsed;
                pickerHint.Visibility  = needPicker ? Visibility.Visible : Visibility.Collapsed;

                pickerHint.Text = _rbSubfolder.IsChecked == true
                    ? "Browse to select the base path. A subfolder per customer will be created inside it."
                    : "Browse to select the shared output folder used for all customers.";
            }

            _rbSame.Checked      += (_, _) => UpdatePickerVisibility();
            _rbShared.Checked    += (_, _) => UpdatePickerVisibility();
            _rbSubfolder.Checked += (_, _) => UpdatePickerVisibility();
            UpdatePickerVisibility();

            // ── Error label ───────────────────────────────────────────────────
            _errorLabel = new TextBlock
            {
                Foreground   = new SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71)),
                FontSize     = 11,
                Visibility   = Visibility.Collapsed,
                Margin       = new Thickness(0, 10, 0, 0),
                TextWrapping = TextWrapping.Wrap
            };
            root.Children.Add(_errorLabel);

            // ── Buttons ───────────────────────────────────────────────────────
            var btnRow = new StackPanel
            {
                Orientation       = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin            = new Thickness(0, 20, 0, 0)
            };

            var cancelBtn = new Button
            {
                Content         = "Cancel",
                Width           = 90,
                Height          = 32,
                Margin          = new Thickness(0, 0, 10, 0),
                Background      = new SolidColorBrush(Color.FromRgb(0x1E, 0x26, 0x40)),
                Foreground      = new SolidColorBrush(FgMuted),
                BorderBrush     = new SolidColorBrush(Border),
                BorderThickness = new Thickness(1),
                Cursor          = System.Windows.Input.Cursors.Hand
            };
            cancelBtn.Click += (_, _) => { DialogResult = false; Close(); };

            _okBtn = new Button
            {
                Content         = "Import",
                Width           = 90,
                Height          = 32,
                Background      = new SolidColorBrush(AccentGo),
                Foreground      = new SolidColorBrush(Colors.White),
                BorderThickness = new Thickness(0),
                FontWeight      = FontWeights.SemiBold,
                Cursor          = System.Windows.Input.Cursors.Hand
            };
            _okBtn.Click += OkBtn_Click;

            btnRow.Children.Add(cancelBtn);
            btnRow.Children.Add(_okBtn);
            root.Children.Add(btnRow);

            return root;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private RadioButton MakeRadio(string text, bool isChecked = false)
        {
            return new RadioButton
            {
                Content         = text,
                IsChecked       = isChecked,
                Foreground      = new SolidColorBrush(FgPrimary),
                FontSize        = 12,
                Margin          = new Thickness(0, 0, 0, 6),
                GroupName       = "ReportsMode"
            };
        }

        private void BrowseShared_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title            = "Select Output Folder",
                Filter           = "Folder|*.folder",
                FileName         = "Select Folder",
                CheckFileExists  = false,
                CheckPathExists  = true,
                ValidateNames    = false
            };
            if (dlg.ShowDialog() != true) return;
            string folder = Path.GetDirectoryName(dlg.FileName) ?? dlg.FileName;
            _sharedBox.Text = folder;
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            bool needFolder = _rbShared.IsChecked == true || _rbSubfolder.IsChecked == true;

            if (needFolder && string.IsNullOrWhiteSpace(_sharedBox.Text))
            {
                _errorLabel.Text       = "Please browse and select an output folder.";
                _errorLabel.Visibility = Visibility.Visible;
                return;
            }

            if (needFolder && !Directory.Exists(_sharedBox.Text))
            {
                _errorLabel.Text       = "The selected folder does not exist. Please choose a valid path.";
                _errorLabel.Visibility = Visibility.Visible;
                return;
            }

            ReportsMode          = _rbSame.IsChecked      == true ? "same"
                                 : _rbShared.IsChecked    == true ? "shared"
                                 : "subfolder";
            SharedReportsFolder  = _sharedBox.Text.Trim();

            DialogResult = true;
            Close();
        }
    }
}
