using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.IO;
using System.Reflection;

namespace ParametricBuilder
{
    public static class CadAiLogger
    {
        public static event EventHandler<string> LogMessageWritten;

        public static Logger ConfigureLogger()
        {
            // Get the directory path of the DLL
            string dllPath = Assembly.GetExecutingAssembly().Location;
            string dllDirectory = Path.GetDirectoryName(dllPath);

            // Create a log file name with the current date
            string logFileName = $"logger_{DateTime.Now:yyyyMMdd}.log";
            string logFilePath = Path.Combine(dllDirectory, "logs", logFileName); // Create a 'logs' folder

            // Configure NLog logger
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                FileName = logFilePath, // Use the DLL directory with the log file name
                Layout = "${longdate}|${level:uppercase=true}|${logger}|${message}|${exception:format=tostring}"
            };
            config.AddTarget("file", fileTarget);
            config.LoggingRules.Add(new LoggingRule("*", LogLevel.Debug, fileTarget));

            var customEventTarget = new CustomEventTarget
            {
                Layout = "${longdate}|${level:uppercase=true}|${logger}|${message}|${exception:format=tostring}"
            };
            config.AddTarget("customEvent", customEventTarget);
            config.LoggingRules.Add(new LoggingRule("*", LogLevel.Debug, customEventTarget));

            LogManager.Configuration = config;

            CustomEventTarget.LogMessageWritten += (sender, logMessage) =>
            {
                LogMessageWritten?.Invoke(sender, logMessage);
            };

            return LogManager.GetCurrentClassLogger();
        }

    }

    public sealed class CustomEventTarget : TargetWithLayout
    {
        public static event EventHandler<string> LogMessageWritten;

        protected override void Write(LogEventInfo logEvent)
        {
            var logMessage = this.Layout.Render(logEvent);
            LogMessageWritten?.Invoke(this, logMessage);
        }
    }
}
