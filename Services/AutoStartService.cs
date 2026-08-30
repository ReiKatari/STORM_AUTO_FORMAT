using System;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace GSheetAutoConverter.Services
{
    public class AutoStartService
    {
        private const string AppName = "STORM AUTO FORMAT";
        private const string RunRegistryPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";

        public bool IsAutoStartEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunRegistryPath, false);
                var value = key?.GetValue(AppName) as string;
                return !string.IsNullOrEmpty(value);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error checking auto-start registry: {ex.Message}");
                return false;
            }
        }

        public bool SetAutoStart(bool enable)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunRegistryPath, true);
                if (key == null) return false;

                if (enable)
                {
                    string executablePath = Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty;
                    if (string.IsNullOrEmpty(executablePath) || executablePath.EndsWith("dotnet.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        executablePath = Path.Combine(AppContext.BaseDirectory, "StormAutoFormat.exe");
                    }
                    
                    // Launch app directly on Windows startup so interface is displayed
                    string command = $"\"{executablePath}\"";
                    key.SetValue(AppName, command);
                }
                else
                {
                    if (key.GetValue(AppName) != null)
                    {
                        key.DeleteValue(AppName, false);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting auto-start registry: {ex.Message}");
                return false;
            }
        }
    }
}
