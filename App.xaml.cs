using System;
using System.Threading;
using System.Windows;

namespace GSheetAutoConverter
{
    public partial class App : System.Windows.Application
    {
        private static Mutex? _mutex;
        private static EventWaitHandle? _showWindowEvent;
        private static Thread? _eventWaitThread;
        private static bool _ownsMutex;

        protected override void OnStartup(StartupEventArgs e)
        {
            const string mutexName = "STORM_AUTO_FORMAT_SingleInstance_Mutex";
            const string eventName = "STORM_AUTO_FORMAT_ShowWindow_Event";

            bool createdNew;
            _mutex = new Mutex(true, mutexName, out createdNew);
            _showWindowEvent = new EventWaitHandle(false, EventResetMode.AutoReset, eventName);

            if (!createdNew)
            {
                // Signal running instance to restore window and bring to front
                try
                {
                    _showWindowEvent.Set();
                }
                catch { }

                // Exit secondary instance
                Environment.Exit(0);
                return;
            }

            _ownsMutex = true;

            base.OnStartup(e);

            // Create and show main window (UI interface displays on startup and reboot)
            var mainWindow = new MainWindow();
            mainWindow.Show();
            mainWindow.Activate();
            mainWindow.Focus();

            // Listen for secondary app launches to restore window
            _eventWaitThread = new Thread(() =>
            {
                while (_showWindowEvent != null && _showWindowEvent.WaitOne())
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        mainWindow.ShowAndRestoreWindow();
                    }));
                }
            })
            {
                IsBackground = true
            };
            _eventWaitThread.Start();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            if (_ownsMutex)
            {
                try
                {
                    _mutex?.ReleaseMutex();
                }
                catch { }
            }
            _showWindowEvent?.Close();
            base.OnExit(e);
        }
    }
}
