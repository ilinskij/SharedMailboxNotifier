using System;
using System.Diagnostics;
using Microsoft.Win32;
using SharedMailboxNotifier.Resources;

namespace SharedMailboxNotifier.Services
{
    /// <summary>
    /// Service for reading and writing add-in settings from registry.
    /// Settings are stored in HKCU\Software\SharedMailboxNotifier
    /// Also manages Outlook notification settings.
    /// </summary>
    public static class SettingsService
    {
        private const string RegistryPath = @"Software\SharedMailboxNotifier";

        // Setting names
        private const string KeyMonitorMode = "MonitorMode";
        private const string KeyEnableSound = "EnableSound";
        private const string KeySavedDesktopAlerts = "SavedDesktopAlerts"; // Backup of user's Outlook setting
        private const string KeyRoundAppLogo = "RoundAppLogo";
        private const string KeySearchContactPhotos = "SearchContactPhotos";
        private const string KeySelectedCategories = "SelectedCategories";

        // Default values
        private const int DefaultMonitorMode = 1; // SharedOnly
        private const bool DefaultEnableSound = true;
        private const bool DefaultRoundAppLogo = false; // Square by default
        private const bool DefaultSearchContactPhotos = true; // On by default — contact search can be slow

        // Outlook registry paths (for different versions)
        private static readonly string[] OutlookPreferencesPaths = new[]
        {
            @"Software\Microsoft\Office\16.0\Outlook\Preferences", // Office 2016/2019/365
            @"Software\Microsoft\Office\15.0\Outlook\Preferences", // Office 2013
            @"Software\Microsoft\Office\14.0\Outlook\Preferences", // Office 2010
        };

        // Outlook registry keys
        private const string OutlookNewMailDesktopAlerts = "NewmailDesktopAlerts";
        private const string OutlookPlaySound = "PlaySound";

        /// <summary>
        /// Monitor mode: 0 = All mailboxes, 1 = Shared mailboxes only
        /// </summary>
        public enum MonitorModeEnum
        {
            AllMailboxes = 0,
            SharedMailboxesOnly = 1
        }

        /// <summary>
        /// Gets or sets the monitor mode.
        /// </summary>
        public static MonitorModeEnum MonitorMode
        {
            get
            {
                var value = GetInt(KeyMonitorMode, DefaultMonitorMode);
                return (MonitorModeEnum)value;
            }
            set
            {
                SetInt(KeyMonitorMode, (int)value);
            }
        }

        /// <summary>
        /// Gets or sets whether sound notifications are enabled.
        /// </summary>
        public static bool EnableSound
        {
            get { return GetBool(KeyEnableSound, DefaultEnableSound); }
            set { SetBool(KeyEnableSound, value); }
        }

        /// <summary>
        /// Gets or sets whether the app logo should be cropped to a circle.
        /// </summary>
        public static bool RoundAppLogo
        {
            get { return GetBool(KeyRoundAppLogo, DefaultRoundAppLogo); }
            set { SetBool(KeyRoundAppLogo, value); }
        }

        /// <summary>
        /// Gets or sets whether to search local Contacts for sender photos.
        /// When enabled, may slow down notifications for large address books.
        /// </summary>
        public static bool SearchContactPhotos
        {
            get { return GetBool(KeySearchContactPhotos, DefaultSearchContactPhotos); }
            set { SetBool(KeySearchContactPhotos, value); }
        }

        /// <summary>
        /// Gets or sets the user's selected category names for toast notifications.
        /// Stored as comma-separated string (same format as Outlook's Categories property).
        /// Returns null if no selection has been saved yet (first run).
        /// </summary>
        public static string SelectedCategories
        {
            get { return GetString(KeySelectedCategories, null); }
            set { SetString(KeySelectedCategories, value); }
        }

        /// <summary>
        /// Returns true if the user has ever configured category selection.
        /// When false, CategoryService should auto-populate with first N categories.
        /// </summary>
        public static bool HasSelectedCategories
        {
            get { return SelectedCategories != null; }
        }

        #region Outlook Notification Control

        /// <summary>
        /// Disables Outlook's built-in desktop alerts.
        /// Changes take effect after Outlook restart.
        /// </summary>
        /// <returns>True if successful</returns>
        public static bool DisableOutlookDesktopAlerts()
        {
            return SetOutlookNotificationSetting(OutlookNewMailDesktopAlerts, 0);
        }

        /// <summary>
        /// Enables Outlook's built-in desktop alerts.
        /// Changes take effect after Outlook restart.
        /// </summary>
        /// <returns>True if successful</returns>
        public static bool EnableOutlookDesktopAlerts()
        {
            return SetOutlookNotificationSetting(OutlookNewMailDesktopAlerts, 1);
        }

        /// <summary>
        /// Disables Outlook's new mail sound.
        /// Changes take effect after Outlook restart.
        /// </summary>
        /// <returns>True if successful</returns>
        public static bool DisableOutlookSound()
        {
            return SetOutlookNotificationSetting(OutlookPlaySound, 0);
        }

        /// <summary>
        /// Enables Outlook's new mail sound.
        /// Changes take effect after Outlook restart.
        /// </summary>
        /// <returns>True if successful</returns>
        public static bool EnableOutlookSound()
        {
            return SetOutlookNotificationSetting(OutlookPlaySound, 1);
        }

        /// <summary>
        /// Gets current state of Outlook desktop alerts.
        /// </summary>
        public static bool IsOutlookDesktopAlertsEnabled()
        {
            return GetOutlookNotificationSetting(OutlookNewMailDesktopAlerts, 1) != 0;
        }

        /// <summary>
        /// Gets current state of Outlook sound.
        /// </summary>
        public static bool IsOutlookSoundEnabled()
        {
            return GetOutlookNotificationSetting(OutlookPlaySound, 1) != 0;
        }

        /// <summary>
        /// Applies Outlook notification settings based on current monitor mode.
        /// Call this when settings are saved.
        /// When switching to "All mailboxes" - saves user's current setting before disabling.
        /// When switching to "Shared only" - restores user's saved setting.
        /// </summary>
        /// <returns>True if Outlook restart is required</returns>
        public static bool ApplyOutlookNotificationSettings()
        {
            var mode = MonitorMode;
            bool changed = false;

            if (mode == MonitorModeEnum.AllMailboxes)
            {
                // "All mailboxes" mode - disable Outlook's own notifications
                // to avoid duplicates
                
                // First, save the current user's setting (only if not already saved)
                int? savedValue = GetSavedDesktopAlerts();
                if (savedValue == null)
                {
                    // No saved value yet - save current Outlook setting
                    int currentValue = GetOutlookNotificationSetting(OutlookNewMailDesktopAlerts, 1);
                    SetSavedDesktopAlerts(currentValue);
                    Debug.WriteLine(string.Format("[SettingsService] Saved user's DesktopAlerts setting: {0}", currentValue));
                    LogService.Info(string.Format(Strings.LogOutlookAlertsSaved,
                        currentValue == 1 ? "enabled" : "disabled"));
                }

                // Now disable Outlook alerts
                if (IsOutlookDesktopAlertsEnabled())
                {
                    DisableOutlookDesktopAlerts();
                    changed = true;
                    LogService.Info(Strings.LogOutlookAlertsDisabled);
                }
            }
            else
            {
                // "Shared only" mode - restore Outlook's notifications
                // for personal mailboxes
                
                // Restore saved value (or default to 1 = enabled)
                int? savedValue = GetSavedDesktopAlerts();
                int valueToRestore = savedValue ?? 1; // Default: enabled
                
                int currentValue = GetOutlookNotificationSetting(OutlookNewMailDesktopAlerts, 1);
                if (currentValue != valueToRestore)
                {
                    SetOutlookNotificationSetting(OutlookNewMailDesktopAlerts, valueToRestore);
                    changed = true;
                    Debug.WriteLine(string.Format("[SettingsService] Restored user's DesktopAlerts setting: {0}", valueToRestore));
                    LogService.Info(string.Format(Strings.LogOutlookAlertsRestored,
                        valueToRestore == 1 ? "enabled" : "disabled"));
                }

                // Clear saved value - it's been restored
                if (savedValue != null)
                {
                    ClearSavedDesktopAlerts();
                }
            }

            // Handle sound separately based on our setting
            if (EnableSound)
            {
                if (!IsOutlookSoundEnabled())
                {
                    EnableOutlookSound();
                    changed = true;
                    LogService.Info(string.Format(Strings.LogOutlookSoundChanged, "enabled"));
                }
            }
            else
            {
                if (IsOutlookSoundEnabled())
                {
                    DisableOutlookSound();
                    changed = true;
                    LogService.Info(string.Format(Strings.LogOutlookSoundChanged, "disabled"));
                }
            }

            return changed;
        }

        /// <summary>
        /// Gets the saved desktop alerts value, or null if not saved.
        /// </summary>
        private static int? GetSavedDesktopAlerts()
        {
            try
            {
                using (var key = OpenKey())
                {
                    if (key != null)
                    {
                        var value = key.GetValue(KeySavedDesktopAlerts);
                        if (value is int)
                        {
                            return (int)value;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsService] Failed to read saved alerts: " + ex.Message);
            }
            return null;
        }

        /// <summary>
        /// Saves the desktop alerts value.
        /// </summary>
        private static void SetSavedDesktopAlerts(int value)
        {
            SetInt(KeySavedDesktopAlerts, value);
        }

        /// <summary>
        /// Clears the saved desktop alerts value.
        /// </summary>
        private static void ClearSavedDesktopAlerts()
        {
            try
            {
                using (var key = OpenKey(writable: true))
                {
                    if (key != null)
                    {
                        key.DeleteValue(KeySavedDesktopAlerts, throwOnMissingValue: false);
                        Debug.WriteLine("[SettingsService] Cleared saved DesktopAlerts setting");
                    }
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsService] Failed to clear saved alerts: " + ex.Message);
            }
        }

        private static bool SetOutlookNotificationSetting(string valueName, int value)
        {
            bool success = false;

            foreach (var path in OutlookPreferencesPaths)
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(path, writable: true))
                    {
                        if (key != null)
                        {
                            key.SetValue(valueName, value, RegistryValueKind.DWord);
                            success = true;
                            Debug.WriteLine(string.Format("[SettingsService] Set {0}={1} in {2}", valueName, value, path));
                            break; // Found the right version, stop
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine(string.Format("[SettingsService] Failed to set {0} in {1}: {2}", valueName, path, ex.Message));
                }
            }

            return success;
        }

        private static int GetOutlookNotificationSetting(string valueName, int defaultValue)
        {
            foreach (var path in OutlookPreferencesPaths)
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(path))
                    {
                        if (key != null)
                        {
                            var value = key.GetValue(valueName);
                            if (value is int)
                            {
                                return (int)value;
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine(string.Format("[SettingsService] Failed to read {0} from {1}: {2}", valueName, path, ex.Message));
                }
            }

            return defaultValue;
        }

        #endregion

        #region Add-in Settings Registry Helpers

        private static RegistryKey OpenKey(bool writable = false)
        {
            try
            {
                if (writable)
                {
                    return Registry.CurrentUser.CreateSubKey(RegistryPath);
                }
                else
                {
                    return Registry.CurrentUser.OpenSubKey(RegistryPath);
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsService] Failed to open registry key: " + ex.Message);
                return null;
            }
        }

        private static int GetInt(string name, int defaultValue)
        {
            try
            {
                using (var key = OpenKey())
                {
                    if (key != null)
                    {
                        var value = key.GetValue(name);
                        if (value is int)
                        {
                            return (int)value;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsService] Failed to read " + name + ": " + ex.Message);
            }

            return defaultValue;
        }

        private static void SetInt(string name, int value)
        {
            try
            {
                using (var key = OpenKey(writable: true))
                {
                    if (key != null)
                    {
                        key.SetValue(name, value, RegistryValueKind.DWord);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsService] Failed to write " + name + ": " + ex.Message);
            }
        }

        private static bool GetBool(string name, bool defaultValue)
        {
            return GetInt(name, defaultValue ? 1 : 0) != 0;
        }

        private static void SetBool(string name, bool value)
        {
            SetInt(name, value ? 1 : 0);
        }

        private static string GetString(string name, string defaultValue)
        {
            try
            {
                using (var key = OpenKey())
                {
                    if (key != null)
                    {
                        var value = key.GetValue(name) as string;
                        if (value != null)
                        {
                            return value;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsService] Failed to read " + name + ": " + ex.Message);
            }

            return defaultValue;
        }

        private static void SetString(string name, string value)
        {
            try
            {
                using (var key = OpenKey(writable: true))
                {
                    if (key != null)
                    {
                        if (value != null)
                        {
                            key.SetValue(name, value, RegistryValueKind.String);
                        }
                        else
                        {
                            key.DeleteValue(name, throwOnMissingValue: false);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsService] Failed to write " + name + ": " + ex.Message);
            }
        }

        #endregion
    }
}
