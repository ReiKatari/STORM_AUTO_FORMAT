using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Threading;
using GSheetAutoConverter.Models;
using GSheetAutoConverter.Services;
using Win32OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Win32SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using Forms = System.Windows.Forms;
using WpfDragEventArgs = System.Windows.DragEventArgs;
using WpfApplication = System.Windows.Application;
using WpfButton = System.Windows.Controls.Button;

namespace GSheetAutoConverter
{
    public partial class MainWindow : Window
    {
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

        private readonly AppSettingsService _settingsService = new();
        private readonly AutoStartService _autoStartService = new();
        private readonly GSheetConverterService _converterService = new();

        private AppSettings _settings;
        private Forms.NotifyIcon? _notifyIcon;
        private DispatcherTimer? _countdownTimer;

        private int _remainingSeconds;
        private DateTime? _lastScheduledExecutionTime;
        private bool _isConverting;
        private bool _isExplicitExit;
        private bool _isInitialized;

        public MainWindow()
        {
            InitializeComponent();
            _settings = _settingsService.LoadSettings();

            Loaded += MainWindow_Loaded;

            LoadSettingsToUi();
            InitializeTimer();

            _isInitialized = true;

            LogMessage("🟢 STORM AUTO FORMAT запущен. Готов к автоматической конвертации.");

            // Perform initial conversion on launch if GSheet file is set
            if (!string.IsNullOrWhiteSpace(_settings.GSheetFilePath))
            {
                Dispatcher.BeginInvoke(new Action(async () => await PerformConversionAsync()), DispatcherPriority.Background);
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            EnableDarkTitleBar();
            SetWindowIconSafely();
            InitializeTrayIcon();
        }

        private void EnableDarkTitleBar()
        {
            try
            {
                var helper = new WindowInteropHelper(this);
                int value = 1;
                DwmSetWindowAttribute(helper.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to set dark title bar: {ex.Message}");
            }
        }

        private void SetWindowIconSafely()
        {
            try
            {
                using var icon = GetAppIcon();
                if (icon != null)
                {
                    var bitmapSource = Imaging.CreateBitmapSourceFromHIcon(
                        icon.Handle,
                        Int32Rect.Empty,
                        System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());

                    Icon = bitmapSource;
                    if (ImgHeaderIcon != null)
                    {
                        ImgHeaderIcon.Source = bitmapSource;
                    }
                }
            }
            catch { }
        }

        private System.Drawing.Icon GetAppIcon()
        {
            try
            {
                string exePath = Process.GetCurrentProcess().MainModule?.FileName ?? "";
                if (!string.IsNullOrEmpty(exePath) && File.Exists(exePath))
                {
                    var icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
                    if (icon != null) return icon;
                }
            }
            catch { }

            try
            {
                string iconPath = Path.Combine(AppContext.BaseDirectory, "app.ico");
                if (File.Exists(iconPath))
                {
                    return new System.Drawing.Icon(iconPath);
                }
            }
            catch { }

            return System.Drawing.SystemIcons.Application;
        }

        private void InitializeTrayIcon()
        {
            try
            {
                if (_notifyIcon != null)
                {
                    _notifyIcon.Visible = false;
                    _notifyIcon.Dispose();
                }

                _notifyIcon = new Forms.NotifyIcon
                {
                    Icon = GetAppIcon(),
                    Text = "STORM AUTO FORMAT",
                    Visible = true
                };

                var contextMenu = new Forms.ContextMenuStrip();

                var openItem = new Forms.ToolStripMenuItem("Показать окно STORM AUTO FORMAT", null, (s, e) => ShowAndRestoreWindow());
                openItem.Font = new Font(openItem.Font, System.Drawing.FontStyle.Bold);

                var convertItem = new Forms.ToolStripMenuItem("⚡ Сконвертировать сейчас", null, async (s, e) =>
                {
                    await Dispatcher.InvokeAsync(async () => await PerformConversionAsync());
                });

                var autoStartItem = new Forms.ToolStripMenuItem("Запускать с Windows")
                {
                    Checked = _autoStartService.IsAutoStartEnabled()
                };
                autoStartItem.Click += (s, e) =>
                {
                    bool newState = !autoStartItem.Checked;
                    if (_autoStartService.SetAutoStart(newState))
                    {
                        autoStartItem.Checked = newState;
                        ChkAutoStart.IsChecked = newState;
                        _settings.AutoStartWithWindows = newState;
                        SaveCurrentUiSettings();
                        LogMessage(newState ? "Включен автозапуск с Windows." : "Автозапуск с Windows отключен.");
                    }
                };

                var exitItem = new Forms.ToolStripMenuItem("Выход", null, (s, e) => ExitApplication());

                contextMenu.Items.Add(openItem);
                contextMenu.Items.Add(convertItem);
                contextMenu.Items.Add(new Forms.ToolStripSeparator());
                contextMenu.Items.Add(autoStartItem);
                contextMenu.Items.Add(new Forms.ToolStripSeparator());
                contextMenu.Items.Add(exitItem);

                _notifyIcon.ContextMenuStrip = contextMenu;

                // Left click or double click restores the window
                _notifyIcon.MouseClick += (s, e) =>
                {
                    if (e.Button == Forms.MouseButtons.Left)
                    {
                        ShowAndRestoreWindow();
                    }
                };
                _notifyIcon.DoubleClick += (s, e) => ShowAndRestoreWindow();
            }
            catch (Exception ex)
            {
                LogMessage($"⚠️ Предупреждение иконки трея: {ex.Message}");
            }
        }

        private void LoadSettingsToUi()
        {
            TxtGSheetPath.Text = _settings.GSheetFilePath;
            TxtXlsxPath.Text = _settings.OutputXlsxPath;
            TxtIntervalMinutes.Text = _settings.SyncIntervalMinutes.ToString();
            TxtScheduledTimes.Text = _settings.ScheduledTimes;

            if (_settings.SyncMode == "ScheduledTime")
            {
                RadModeScheduled.IsChecked = true;
                PnlScheduledMode.Visibility = Visibility.Visible;
                PnlIntervalMode.Visibility = Visibility.Collapsed;
            }
            else
            {
                RadModeInterval.IsChecked = true;
                PnlIntervalMode.Visibility = Visibility.Visible;
                PnlScheduledMode.Visibility = Visibility.Collapsed;
            }

            ChkAutoStart.IsChecked = _autoStartService.IsAutoStartEnabled();
            ChkMinimizeOnClose.IsChecked = _settings.MinimizeToTrayOnClose;
            TxtGoogleCookie.Text = _settings.GoogleAuthCookie;

            UpdateOutputSuggestionIfEmpty();
        }

        private void SaveCurrentUiSettings()
        {
            if (!_isInitialized) return;

            _settings.GSheetFilePath = TxtGSheetPath.Text.Trim();
            _settings.OutputXlsxPath = TxtXlsxPath.Text.Trim();

            _settings.SyncMode = RadModeScheduled.IsChecked == true ? "ScheduledTime" : "Interval";

            if (int.TryParse(TxtIntervalMinutes.Text, out int interval) && interval > 0)
            {
                _settings.SyncIntervalMinutes = interval;
            }

            _settings.ScheduledTimes = TxtScheduledTimes.Text.Trim();
            _settings.AutoStartWithWindows = ChkAutoStart.IsChecked == true;
            _settings.MinimizeToTrayOnClose = ChkMinimizeOnClose.IsChecked == true;
            _settings.GoogleAuthCookie = TxtGoogleCookie.Text.Trim();

            _settingsService.SaveSettings(_settings);
        }

        private void InitializeTimer()
        {
            _remainingSeconds = GetIntervalInSeconds();

            _countdownTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _countdownTimer.Tick += CountdownTimer_Tick;
            _countdownTimer.Start();
        }

        private async void CountdownTimer_Tick(object? sender, EventArgs e)
        {
            if (_isConverting) return;

            if (RadModeScheduled.IsChecked == true)
            {
                // --- Scheduled Time Mode (Exact Times) ---
                var nextTarget = GetNextScheduledTime();
                if (nextTarget.HasValue)
                {
                    TimeSpan remaining = nextTarget.Value - DateTime.Now;
                    if (remaining.TotalSeconds <= 0)
                    {
                        // Check to prevent double execution in same minute
                        if (!_lastScheduledExecutionTime.HasValue || (DateTime.Now - _lastScheduledExecutionTime.Value).TotalSeconds > 50)
                        {
                            _lastScheduledExecutionTime = DateTime.Now;
                            await PerformConversionAsync();
                        }
                    }

                    // Display countdown to target time
                    if (remaining.TotalSeconds < 0) remaining = TimeSpan.Zero;

                    LblCountdownPrefix.Text = $"🕒 Выгрузка в {nextTarget.Value:HH:mm:ss} (через): ";
                    if (remaining.TotalHours >= 1)
                    {
                        TxtCountdown.Text = $"{(int)remaining.TotalHours:D2}:{remaining.Minutes:D2}:{remaining.Seconds:D2}";
                    }
                    else
                    {
                        TxtCountdown.Text = $"{remaining.Minutes:D2}:{remaining.Seconds:D2}";
                    }
                }
                else
                {
                    LblCountdownPrefix.Text = "🕒 Укажите время (HH:mm): ";
                    TxtCountdown.Text = "--:--";
                }
            }
            else
            {
                // --- Interval Mode (Every N minutes) ---
                LblCountdownPrefix.Text = "⏱ Следующее авто-обновление через: ";
                _remainingSeconds--;

                if (_remainingSeconds <= 0)
                {
                    _remainingSeconds = GetIntervalInSeconds();
                    await PerformConversionAsync();
                }

                int mins = _remainingSeconds / 60;
                int secs = _remainingSeconds % 60;
                TxtCountdown.Text = $"{mins:D2}:{secs:D2}";
            }
        }

        private DateTime? GetNextScheduledTime()
        {
            string input = TxtScheduledTimes.Text.Trim();
            if (string.IsNullOrWhiteSpace(input)) return null;

            var timeParts = input.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var scheduledTimeSpans = new List<TimeSpan>();

            foreach (var part in timeParts)
            {
                if (TimeSpan.TryParse(part.Trim(), out var parsedTime))
                {
                    scheduledTimeSpans.Add(parsedTime);
                }
            }

            if (scheduledTimeSpans.Count == 0) return null;

            DateTime now = DateTime.Now;
            DateTime? earliestNext = null;

            foreach (var timeSpan in scheduledTimeSpans)
            {
                DateTime targetToday = DateTime.Today.Add(timeSpan);
                if (targetToday > now)
                {
                    if (earliestNext == null || targetToday < earliestNext)
                    {
                        earliestNext = targetToday;
                    }
                }
            }

            // If all today's times passed, pick earliest time tomorrow
            if (earliestNext == null)
            {
                var minSpan = scheduledTimeSpans.OrderBy(t => t).First();
                earliestNext = DateTime.Today.AddDays(1).Add(minSpan);
            }

            return earliestNext;
        }

        private int GetIntervalInSeconds()
        {
            int mins = _settings.SyncIntervalMinutes;
            if (mins <= 0) mins = 5;
            return mins * 60;
        }

        private async Task PerformConversionAsync()
        {
            if (_isConverting) return;

            string gsheetInput = TxtGSheetPath.Text.Trim();
            if (string.IsNullOrWhiteSpace(gsheetInput))
            {
                TxtStatus.Text = "Ожидание указания файла .gsheet";
                StatusDot.Background = System.Windows.Media.Brushes.Orange;
                return;
            }

            _isConverting = true;
            BtnConvertNow.IsEnabled = false;
            TxtStatus.Text = "Выполняется конвертация...";
            StatusDot.Background = System.Windows.Media.Brushes.Yellow;

            string targetXlsx = TxtXlsxPath.Text.Trim();

            LogMessage($"⏳ Запуск конвертации: {Path.GetFileName(gsheetInput)}...");

            try
            {
                var result = await _converterService.ConvertGSheetToXlsxAsync(
                    gsheetInput,
                    targetXlsx,
                    _settings.GoogleAuthCookie
                );

                if (result.Success)
                {
                    string sizeKb = (result.FileSizeBytes / 1024.0).ToString("N1");
                    string msg = $"✅ Сохранено: {Path.GetFileName(result.OutputPath)} ({sizeKb} KB) в {DateTime.Now:HH:mm:ss}";
                    LogMessage(msg);

                    TxtStatus.Text = $"Активно (последнее обновление: {DateTime.Now:HH:mm:ss})";
                    StatusDot.Background = System.Windows.Media.Brushes.LimeGreen;

                    // Send Tray Notification
                    _notifyIcon?.ShowBalloonTip(3000, "STORM AUTO FORMAT", $"Файл {Path.GetFileName(result.OutputPath)} успешно обновлен ({sizeKb} KB)", Forms.ToolTipIcon.Info);
                }
                else
                {
                    LogMessage($"❌ Ошибка: {result.ErrorMessage}");
                    TxtStatus.Text = "Ошибка при конвертации";
                    StatusDot.Background = System.Windows.Media.Brushes.Red;

                    _notifyIcon?.ShowBalloonTip(5000, "STORM AUTO FORMAT - Ошибка", result.ErrorMessage, Forms.ToolTipIcon.Error);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Исключение: {ex.Message}");
                TxtStatus.Text = "Критическая ошибка";
                StatusDot.Background = System.Windows.Media.Brushes.Red;
            }
            finally
            {
                _isConverting = false;
                BtnConvertNow.IsEnabled = true;
                _remainingSeconds = GetIntervalInSeconds();
            }
        }

        private void SyncMode_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;

            if (RadModeScheduled.IsChecked == true)
            {
                PnlScheduledMode.Visibility = Visibility.Visible;
                PnlIntervalMode.Visibility = Visibility.Collapsed;
            }
            else
            {
                PnlIntervalMode.Visibility = Visibility.Visible;
                PnlScheduledMode.Visibility = Visibility.Collapsed;
                _remainingSeconds = GetIntervalInSeconds();
            }

            SaveCurrentUiSettings();
        }

        private void BtnPresetInterval_Click(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.Tag != null)
            {
                TxtIntervalMinutes.Text = btn.Tag.ToString();
            }
        }

        private void BtnPresetTime_Click(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.Tag != null)
            {
                TxtScheduledTimes.Text = btn.Tag.ToString();
            }
        }

        private void LogMessage(string message)
        {
            string logLine = $"[{DateTime.Now:HH:mm:ss}] {message}";
            LstLog.Items.Insert(0, logLine);

            // Keep max 200 items in view
            while (LstLog.Items.Count > 200)
            {
                LstLog.Items.RemoveAt(LstLog.Items.Count - 1);
            }
        }

        public void ShowAndRestoreWindow()
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
            Focus();
            Topmost = true;
            Topmost = false;
        }

        private void ExitApplication()
        {
            _isExplicitExit = true;
            SaveCurrentUiSettings();
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
            }
            WpfApplication.Current.Shutdown();
        }

        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
            {
                Hide();
                _notifyIcon?.ShowBalloonTip(2000, "STORM AUTO FORMAT", "Программа свернута в трей.", Forms.ToolTipIcon.Info);
            }
            base.OnStateChanged(e);
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            SaveCurrentUiSettings();

            if (!_isExplicitExit && (_settings.MinimizeToTrayOnClose))
            {
                e.Cancel = true;
                Hide();
                _notifyIcon?.ShowBalloonTip(2000, "STORM AUTO FORMAT", "Программа продолжает работу в системном трее.", Forms.ToolTipIcon.Info);
                return;
            }

            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
            }
            base.OnClosing(e);
        }

        private void BtnBrowseGSheet_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Win32OpenFileDialog
            {
                Filter = "Google Sheets (*.gsheet)|*.gsheet|Ярлыки (*.url;*.gsheet)|*.url;*.gsheet|Все файлы (*.*)|*.*",
                Title = "Выберите файл Google Sheets (.gsheet)"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                TxtGSheetPath.Text = openFileDialog.FileName;
            }
        }

        private void BtnBrowseXlsx_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new Win32SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                Title = "Выберите путь для сохранения файла Excel (.xlsx)",
                FileName = Path.GetFileName(TxtXlsxPath.Text)
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                TxtXlsxPath.Text = saveFileDialog.FileName;
            }
        }

        private async void BtnConvertNow_Click(object sender, RoutedEventArgs e)
        {
            await PerformConversionAsync();
        }

        private void BtnMinimizeToTray_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            LstLog.Items.Clear();
        }

        private void Window_DragOver(object sender, WpfDragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                e.Effects = System.Windows.DragDropEffects.Copy;
            }
            else
            {
                e.Effects = System.Windows.DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void Window_Drop(object sender, WpfDragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                string[]? files = e.Data.GetData(System.Windows.DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0)
                {
                    string droppedFile = files[0];
                    TxtGSheetPath.Text = droppedFile;
                    LogMessage($"📥 Перетащен файл: {Path.GetFileName(droppedFile)}");
                }
            }
        }

        private void TxtGSheetPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;

            UpdateOutputSuggestionIfEmpty();
            SaveCurrentUiSettings();
        }

        private void TxtXlsxPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;

            SaveCurrentUiSettings();
        }

        private void UpdateOutputSuggestionIfEmpty()
        {
            if (string.IsNullOrWhiteSpace(TxtXlsxPath.Text) && !string.IsNullOrWhiteSpace(TxtGSheetPath.Text))
            {
                string suggested = _converterService.SuggestOutputPath(TxtGSheetPath.Text);
                if (!string.IsNullOrEmpty(suggested))
                {
                    TxtXlsxPath.Text = suggested;
                }
            }
        }

        private void TxtIntervalMinutes_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;

            if (int.TryParse(TxtIntervalMinutes.Text, out int interval) && interval > 0)
            {
                _remainingSeconds = GetIntervalInSeconds();
            }
            SaveCurrentUiSettings();
        }

        private void TxtScheduledTimes_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;

            SaveCurrentUiSettings();
        }

        private void ChkAutoStart_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;

            bool isChecked = ChkAutoStart.IsChecked == true;
            if (_autoStartService.SetAutoStart(isChecked))
            {
                SaveCurrentUiSettings();
                LogMessage(isChecked ? "Автозапуск с Windows включен." : "Автозапуск с Windows отключен.");
            }
        }

        private void ChkMinimizeOnClose_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;

            SaveCurrentUiSettings();
        }

        private void TxtGoogleCookie_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;

            SaveCurrentUiSettings();
        }
    }
}
