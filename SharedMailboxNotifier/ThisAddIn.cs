using System;
using System.Diagnostics;
using SharedMailboxNotifier.Services;
using SharedMailboxNotifier.UI;
using SharedMailboxNotifier.Resources;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SharedMailboxNotifier
{
    /// <summary>
    /// Main entry point for the Outlook VSTO Add-in.
    /// Initializes the shared mailbox monitor when Outlook starts.
    /// </summary>
    public partial class ThisAddIn
    {
        private SharedMailboxMonitor _monitor;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Debug.WriteLine("[SharedMailboxNotifier] Add-in starting...");
            LogService.Initialize();
            CategoryService.InitializeSelectedCategories(Application);

            // Subscribe to settings page event first — settings should work
            // even if monitor initialization fails
            Application.OptionsPagesAdd += Application_OptionsPagesAdd;

            try
            {
                _monitor = new SharedMailboxMonitor(Application);
                _monitor.StatusChanged += OnMonitorStatusChanged;
                _monitor.Initialize();

                Debug.WriteLine("[SharedMailboxNotifier] Add-in started successfully");
            }
            catch (System.Exception ex)
            {
                LogService.Error(Strings.LogFailedToStart, ex);
            }
        }

        private void OnMonitorStatusChanged(object sender, MonitorStatusEventArgs e)
        {
            if (e.IsRunning)
            {
                LogService.LogStartup(e.SharedMailboxCount);
            }

            Debug.WriteLine(string.Format("[SharedMailboxNotifier] Status: {0} (Running: {1}, Mailboxes: {2})",
                e.Message, e.IsRunning, e.SharedMailboxCount));
        }

        private void Application_OptionsPagesAdd(Outlook.PropertyPages Pages)
        {
            try
            {
                Pages.Add(new SettingsPage(Application), "Shared Mailbox Notifier");
                Debug.WriteLine("[SharedMailboxNotifier] Settings page added");
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SharedMailboxNotifier] Error adding settings page: " + ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Debug.WriteLine("[SharedMailboxNotifier] Add-in shutting down...");

            Application.OptionsPagesAdd -= Application_OptionsPagesAdd;

            try
            {
                if (_monitor != null)
                {
                    _monitor.StatusChanged -= OnMonitorStatusChanged;
                    _monitor.Dispose();
                    _monitor = null;
                }

                LogService.LogShutdown();
            }
            catch (System.Exception ex)
            {
                LogService.Error(Strings.LogErrorShutdown, ex);
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
