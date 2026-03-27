using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

public static class ExcelHelper
{
    // P/Invoke to check if a process has a visible main window
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    public static void CloseOrphanedExcelProcesses()
    {
        var excelProcesses = Process.GetProcessesByName("EXCEL");

        foreach (var process in excelProcesses)
        {
            try
            {
                // Check if the process has a main window handle (visible Excel session)
                if (process.MainWindowHandle == IntPtr.Zero || !IsWindowVisible(process.MainWindowHandle))
                {
                    // No visible window => background process (likely orphaned from automation)
                    process.Kill();
                    Console.WriteLine($"Killed background Excel process (PID: {process.Id})");
                }
                else
                {
                    Console.WriteLine($"Skipped visible Excel process (PID: {process.Id})");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error handling Excel process (PID: {process.Id}): {ex.Message}");
            }
        }
    }
}
