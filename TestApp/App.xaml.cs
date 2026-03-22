using System.Linq;
using System.Windows;

namespace TestApp
{
    public partial class App : Application
    {
        /// <summary>
        /// When true the main window starts hidden and goes straight to the
        /// system tray.  Set by the <c>--minimized</c> command-line argument
        /// which is used by the auto-start Task Scheduler entry after a
        /// server reboot.
        /// </summary>
        public static bool StartMinimized { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            StartMinimized = e.Args.Any(a =>
                a.Equals("--minimized", System.StringComparison.OrdinalIgnoreCase) ||
                a.Equals("/minimized",  System.StringComparison.OrdinalIgnoreCase));
        }
    }
}
