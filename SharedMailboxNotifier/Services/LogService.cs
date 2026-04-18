using System;
using System.Diagnostics;
using SharedMailboxNotifier.Resources;

namespace SharedMailboxNotifier.Services
{
    /// <summary>
    /// Logging service that writes to Windows Event Log.
    /// 
    /// Priority order:
    /// 1. "SharedMailboxNotifier" source in "Application" log (created by MSI installer)
    /// 2. Office sources in "OAlerts" log (fallback for debugging without installation)
    /// 3. "Application" log with generic source (universal fallback)
    /// </summary>
    public static class LogService
    {
        private const string ApplicationLogName = "Application";
        private static string _activeSourceName;
        private static string _activeLogName;
        private static volatile bool _initialized;
        private static bool _canWrite;
        private static readonly object _logLock = new object();

        private static readonly string[] ProposedSourceNames = new[]
        {
            "Outlook Shared Mailbox Notifier", // Try our own source (created by MSI)
            "Microsoft Office 16 Alerts",
            "Microsoft Office 15 Alerts",
            "Microsoft Office 14 Alerts",
            "Microsoft Office Alerts",
            "Microsoft Office 16",
            "Microsoft Office 15",
            "Microsoft Office 14",
            "Microsoft Office",
            "Application" // Fallback to Application source
        };

        /// <summary>
        /// Initializes the logging service.
        /// </summary>
        public static void Initialize()
        {
            if (_initialized)
                return;

            lock (_logLock)
            {
                if (_initialized)
                    return;

                try
                {

                    foreach (string sourceName in ProposedSourceNames)
                    {
                        string proposedLogName = AllowedEventLogName(sourceName);
                        if (!String.IsNullOrWhiteSpace(proposedLogName))
                        {
                            Debug.WriteLine("[LogService] Using EventLog " + proposedLogName);
                            Debug.WriteLine("[LogService] Using Source " + sourceName);
                            if (!sourceName.Equals(ProposedSourceNames[0]))
                            {
                                Debug.WriteLine("[LogService] WARNING: Using fallback source, event messages may display incorrectly");
                            }
                            _activeLogName = proposedLogName;
                            _activeSourceName = sourceName;
                            _canWrite = true;
                            _initialized = true;
                            return;
                        }
                    }

                    // No one log was found
                    _canWrite = false;
                    Debug.WriteLine("[LogService] Cannot find accessible log: logging disabled");
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[LogService] Initialization failed, because " + ex.Message);
                    _canWrite = false;
                }

                _initialized = true;
            }
        }

        private static string AllowedEventLogName(string SourceName)
        {

            string LogName = "";
            // Test write
            try
            {
                if (EventLog.SourceExists(SourceName))
                {
                    LogName = EventLog.LogNameFromSourceName(SourceName, ".");
                    using (var log = new EventLog(LogName))
                    {
                        log.Source = SourceName;
                        // Don't actually write, just verify we can
                    }
                    return LogName;
                }
                else { return ""; }
            }
            catch (System.Security.SecurityException)
            {
                // Access denied - can't check source existence
                return "";
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[LogService] Error checking EventLog " + LogName + " and source " + SourceName + ", because " + ex.Message);
                return "";
            }
        }

        public static void Info(string message)
        {
            WriteEntry(message, EventLogEntryType.Information, 1000);
        }

        public static void Warning(string message)
        {
            WriteEntry(message, EventLogEntryType.Warning, 2000);
        }

        public static void Error(string message)
        {
            WriteEntry(message, EventLogEntryType.Error, 3000);
        }

        public static void Error(string message, Exception ex)
        {
            var fullMessage = string.Format("{0}\r\n\r\nException: {1}\r\n{2}",
                message, ex.Message, ex.StackTrace);
            WriteEntry(fullMessage, EventLogEntryType.Error, 3000);
        }

        public static void LogStartup(int sharedMailboxCount)
        {
            var message = string.Format(
                Strings.LogStartupSuccess + "\r\n" +
                Strings.LogMonitoringCount,
                sharedMailboxCount);
            WriteEntry(message, EventLogEntryType.Information, 1001);
        }

        public static void LogShutdown()
        {
            WriteEntry(Strings.LogShutdown, EventLogEntryType.Information, 1002);
        }

        public static void LogNewMail(string mailboxName, string sender, string subject)
        {
            var message = string.Format(
                Strings.LogNewMail + "\r\n" +
                Strings.LogMailbox + "\r\n" +
                Strings.LogFrom + "\r\n" +
                Strings.LogSubject,
                mailboxName, sender, subject);
            WriteEntry(message, EventLogEntryType.Information, 1003);
        }

        private static void WriteEntry(string message, EventLogEntryType type, int eventId)
        {
            // Always write to Debug output
            var prefix = type == EventLogEntryType.Error ? "[ERROR] " :
                         type == EventLogEntryType.Warning ? "[WARN] " : "[INFO] ";
            Debug.WriteLine("[SharedMailboxNotifier] " + prefix + message);

            if (!_initialized)
            {
                Initialize();
            }

            if (!_canWrite)
                return;

            lock (_logLock)
            {
                try
                {
                    EventLog.WriteEntry(_activeSourceName, message, type, eventId);
                }
                catch (System.Exception ex)
                {
                    // Don't throw - logging should never break the app
                    // Just disable further attempts if it fails
                    Debug.WriteLine("[LogService] Write failed: " + ex.Message);

                    // Try fallback to Application
                    if (_activeLogName != ApplicationLogName)
                    {
                        try
                        {
                            _activeSourceName = ApplicationLogName;
                            _activeLogName = ApplicationLogName;
                            EventLog.WriteEntry(_activeSourceName, message, type, eventId);
                        }
                        catch
                        {
                            _canWrite = false;
                        }
                    }
                    else
                    {
                        _canWrite = false;
                    }
                }
            }
        }
    }
}
