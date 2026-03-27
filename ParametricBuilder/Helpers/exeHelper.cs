using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ParametricBuilder.Helpers
{
    public static class exeHelper
    {
        public static async Task RunAsAdministrator(string exePath)
        {
            var processInfo = new ProcessStartInfo
            {
                FileName = exePath,
                //Verb = "runas", // Requests admin privileges
                //UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Normal
            };

            try
            {
                using (var process = Process.Start(processInfo))
                {
                    if (process != null)
                    {
                        await Task.Run(() => process.WaitForExit()); // Waits for process to exit asynchronously
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to start process: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
