using System;
using System.Reflection;
using System.Windows;

namespace TestApp
{
    /// <summary>
    /// System tray icon implemented via reflection against System.Windows.Forms.
    /// No compile-time WinForms reference — avoids all WPF/WinForms namespace collisions.
    /// </summary>
    public sealed class TrayManager : IDisposable
    {
        private readonly Window  _mainWindow;
        private          object? _notifyIcon;      // System.Windows.Forms.NotifyIcon
        private          object? _contextMenu;     // System.Windows.Forms.ContextMenuStrip
        private          Type?   _notifyIconType;
        private          Assembly? _formsAsm;

        public Action? OnWatchAll            { get; set; }
        public Action? OnStopAll             { get; set; }
        public Func<int>? GetActiveWatchCount { get; set; }

        public TrayManager(Window mainWindow)
        {
            _mainWindow = mainWindow;
            try { Init(); } catch { /* tray unavailable — non-fatal */ }
        }

        private void Init()
        {
            _formsAsm = Assembly.Load("System.Windows.Forms");
            _notifyIconType = _formsAsm.GetType("System.Windows.Forms.NotifyIcon")!;
            var menuType    = _formsAsm.GetType("System.Windows.Forms.ContextMenuStrip")!;
            var itemType    = _formsAsm.GetType("System.Windows.Forms.ToolStripMenuItem")!;
            var sepType     = _formsAsm.GetType("System.Windows.Forms.ToolStripSeparator")!;

            _notifyIcon = Activator.CreateInstance(_notifyIconType)!;

            // Set icon from the exe
            try
            {
                var drawingAsm  = Assembly.Load("System.Drawing");
                var iconType    = drawingAsm.GetType("System.Drawing.Icon")!;
                var exePath     = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var icon        = iconType.GetMethod("ExtractAssociatedIcon",
                                    new[] { typeof(string) })!.Invoke(null, new object[] { exePath });
                _notifyIconType.GetProperty("Icon")!.SetValue(_notifyIcon, icon);
            }
            catch { }

            SetProp(_notifyIcon, "Text", "Performance Test Utilities");
            SetProp(_notifyIcon, "Visible", false);

            // ContextMenuStrip
            _contextMenu = Activator.CreateInstance(menuType)!;
            var items = menuType.GetProperty("Items")!.GetValue(_contextMenu)!;
            var addMethod = items.GetType().GetMethod("Add", new[] { _formsAsm.GetType("System.Windows.Forms.ToolStripItem")! })!;

            object MakeItem(string text, Action onClick)
            {
                var item = Activator.CreateInstance(itemType, text)!;
                var clickEvt = itemType.GetEvent("Click")!;
                EventHandler handler = (_, _) => onClick();
                clickEvt.AddEventHandler(item, handler);
                return item;
            }

            addMethod.Invoke(items, new[] { MakeItem("Open / Show", ShowWindow) });
            addMethod.Invoke(items, new[] { Activator.CreateInstance(sepType)! });
            addMethod.Invoke(items, new[] { MakeItem("▶  Watch All Customers", () => OnWatchAll?.Invoke()) });
            addMethod.Invoke(items, new[] { MakeItem("⏹  Stop All Watches",    () => OnStopAll?.Invoke()) });
            addMethod.Invoke(items, new[] { Activator.CreateInstance(sepType)! });
            addMethod.Invoke(items, new[] { MakeItem("Exit", () =>
            {
                OnStopAll?.Invoke();
                SetProp(_notifyIcon!, "Visible", false);
                Application.Current.Shutdown();
            })});

            SetProp(_notifyIcon, "ContextMenuStrip", _contextMenu);

            // Wire DoubleClick
            var clickEvt = _notifyIconType.GetEvent("DoubleClick")!;
            EventHandler dblClick = (_, _) => ShowWindow();
            clickEvt.AddEventHandler(_notifyIcon, dblClick);
        }

        public void MinimizeToTray()
        {
            if (_notifyIcon == null) return;
            int count = GetActiveWatchCount?.Invoke() ?? 0;
            string tip = count > 0
                ? $"Performance Test Utilities — watching {count} customer(s)"
                : "Performance Test Utilities";
            SetProp(_notifyIcon, "Text",    tip);
            SetProp(_notifyIcon, "Visible", true);

            // ShowBalloonTip(int timeout, string title, string text, ToolTipIcon icon)
            var toolTipIconType = _formsAsm?.GetType("System.Windows.Forms.ToolTipIcon");
            var infoVal = toolTipIconType != null
                ? Enum.Parse(toolTipIconType, "Info") : (object)1;
            _notifyIconType?.GetMethod("ShowBalloonTip",
                new[] { typeof(int), typeof(string), typeof(string), toolTipIconType! })?
                .Invoke(_notifyIcon, new[]
                {
                    (object)2000,
                    "Running in background",
                    count > 0
                        ? $"Auto-watch active for {count} customer(s). Double-click to restore."
                        : "Double-click the tray icon to restore.",
                    infoVal
                });

            _mainWindow.Dispatcher.Invoke(() => _mainWindow.Hide());
        }

        public void ShowWindow()
        {
            _mainWindow.Dispatcher.Invoke(() =>
            {
                _mainWindow.Show();
                _mainWindow.WindowState = WindowState.Normal;
                _mainWindow.Activate();
            });
            if (_notifyIcon != null) SetProp(_notifyIcon, "Visible", false);
        }

        public void UpdateTooltip(int activeWatches)
        {
            if (_notifyIcon == null) return;
            if (!(bool)(_notifyIconType?.GetProperty("Visible")?.GetValue(_notifyIcon) ?? false)) return;
            SetProp(_notifyIcon, "Text", activeWatches > 0
                ? $"Performance Test Utilities — watching {activeWatches} customer(s)"
                : "Performance Test Utilities");
        }

        private static void SetProp(object obj, string prop, object? value)
            => obj.GetType().GetProperty(prop)?.SetValue(obj, value);

        public void Dispose()
        {
            if (_notifyIcon == null) return;
            SetProp(_notifyIcon, "Visible", false);
            (_notifyIcon as IDisposable)?.Dispose();
            _notifyIcon = null;
        }
    }
}
