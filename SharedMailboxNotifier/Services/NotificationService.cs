using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using SharedMailboxNotifier.Resources;
using Microsoft.Office.Interop.Outlook;

namespace SharedMailboxNotifier.Services
{
    /// <summary>
    /// Service for displaying notifications.
    /// Uses Windows 10+ Toast notifications when available,
    /// falls back to BalloonTip for older systems.
    /// 
    /// Strategy is chosen once at initialization:
    /// - If Toast API is available and initializes successfully, _toastBackend is non-null.
    /// - BalloonTip is always initialized as a fallback.
    /// - On per-call Toast errors, BalloonTip is used for that call only (no permanent degradation).
    /// </summary>
    public sealed class NotificationService : IDisposable
    {
        private bool _disposed;

        // Outlook Application reference — used for STA-thread operations
        // (category lookup, deleted notification cleanup).
        // NOT used from toast callbacks (MTA) — those use GetActiveObject.
        private readonly Microsoft.Office.Interop.Outlook.Application _outlookApp;

        // Toast backend — null if Toast is not supported or failed to initialize.
        // Once set in constructor, this reference never changes (no mutable strategy flag).
        private readonly ToastBackend _toastBackend;

        // BalloonTip fallback — always available
        private readonly BalloonBackend _balloonBackend;

        // Icon URIs resolved once at startup — null if file not found
        private Uri _iconUriAnswer;
        private Uri _iconUriSetMark;
        private Uri _iconUriMarkAsRead;
        private Uri _iconUriAppLogo;

        // Icon folder name
        private const string _IconFolderName = "Images";

        // Icon file names in the Images folder
        private static readonly string[] IconFileNames =
            { "Answer.png", "SetMark.png", "MarkAsRead.png", "AppLogo.png" };

        // Track shown notifications for removal when mail is read (Toast only)
        private readonly Dictionary<string, (string Tag, string StoreId)> _notificationTags = new Dictionary<string, (string, string)>();
        private readonly object _tagsLock = new object();

        /// <summary>
        /// Maximum length for a Toast notification tag.
        /// Pre-Creators Update (builds 14393–15062) limits tags to 16 characters.
        /// Since we support Anniversary Update and above, use the safe minimum.
        /// </summary>
        private const int MaxToastTagLength = 16;

        public NotificationService(Microsoft.Office.Interop.Outlook.Application outlookApp)
        {
            if (outlookApp == null)
                throw new ArgumentNullException("outlookApp");

            _outlookApp = outlookApp;

            SetupIconUris();

            // Try to create Toast backend (Windows 10 1607+)
            if (IsToastSupported())
            {
                try
                {
                    _toastBackend = new ToastBackend();
                    _toastBackend.Activated += OnToastActivated;
                    Debug.WriteLine("[NotificationService] Toast backend created");
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[NotificationService] Toast init failed, using BalloonTip only: " + ex.Message);
                    _toastBackend = null;
                }
            }

            // Always create BalloonTip backend
            _balloonBackend = new BalloonBackend();

            Debug.WriteLine("[NotificationService] Initialized. Primary mode: " +
                (_toastBackend != null ? "Toast" : "BalloonTip"));
        }

        #region Icon Setup

        /// <summary>
        /// Converts a local file path to a file:// URI.
        /// Returns null if the path is empty or the file does not exist.
        /// </summary>
        private static Uri FilePathToUri(string path)
        {
            if (string.IsNullOrEmpty(path))
                return null;

            if (!File.Exists(path))
                return null;

            UriBuilder uriBuilder = new UriBuilder();
            uriBuilder.Scheme = "file";
            uriBuilder.Path = path;
            return uriBuilder.Uri;
        }

        /// <summary>
        /// Sets an icon URI on a Toast button if the URI is available.
        /// </summary>
        private static void SetButtonIcon(
            Microsoft.Toolkit.Uwp.Notifications.ToastButton button, Uri iconUri)
        {
            if (iconUri != null)
            {
                button.SetImageUri(iconUri);
            }
        }

        private void SetupIconUris()
        {
            try
            {
                var imagesFolder = FindImagesFolder();

                if (imagesFolder != null)
                {
                    var iconUris = IconFileNames
                        .Select(name => FilePathToUri(Path.Combine(imagesFolder, name)))
                        .ToArray();

                    _iconUriAnswer     = iconUris[0];
                    _iconUriSetMark    = iconUris[1];
                    _iconUriMarkAsRead = iconUris[2];
                    _iconUriAppLogo    = iconUris[3];

                    Debug.WriteLine("[NotificationService] Icons folder: " + imagesFolder);
                    Debug.WriteLine("[NotificationService] Answer icon: " + (_iconUriAnswer != null));
                    Debug.WriteLine("[NotificationService] Set Mark icon: " + (_iconUriSetMark != null));
                    Debug.WriteLine("[NotificationService] Mark As Read icon: " + (_iconUriMarkAsRead != null));
                    Debug.WriteLine("[NotificationService] AppLogo: " + (_iconUriAppLogo != null));
                }
                else
                {
                    Debug.WriteLine("[NotificationService] Icons folder not found - buttons will have no icons");
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[NotificationService] Error setting up icon URIs: " + ex.Message);
            }
        }

        /// <summary>
        /// Searches for the Images folder across possible locations.
        /// Returns the full path to the folder, or null if not found.
        /// </summary>
        private static string FindImagesFolder()
        {
            var assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var assemblyFolder = Path.GetDirectoryName(assemblyPath);

            var possiblePaths = new[]
            {
                // Installed: Images subfolder next to DLL
                Path.Combine(assemblyFolder, _IconFolderName),
            
                // Installed: Images in same folder as DLL
                assemblyFolder,
            
                // Debug: bin\Debug\Images (VSTO copies DLL to temp, but Images stays in bin)
                Path.Combine(assemblyFolder, "..", "..", "..", "bin", "Debug", _IconFolderName),
                Path.Combine(assemblyFolder, "..", "..", "..", "bin", "Release", _IconFolderName),
            
            };

            foreach (var path in possiblePaths)
            {
                try
                {
                    var fullPath = Path.GetFullPath(path);
                    if (Directory.Exists(fullPath) && File.Exists(Path.Combine(fullPath, IconFileNames[0])))
                    {
                        return fullPath;
                    }
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[NotificationService] Path probe failed: " + ex.Message);
                }
            }

            // Fallback: try CodeBase
            try
            {
                var codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                if (!string.IsNullOrEmpty(codeBase))
                {
                    var codeBaseFolder = Path.GetDirectoryName(new Uri(codeBase).LocalPath);
                    var imagesPath = Path.Combine(codeBaseFolder, _IconFolderName);
                    if (Directory.Exists(imagesPath))
                    {
                        return imagesPath;
                    }
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[NotificationService] CodeBase probe failed: " + ex.Message);
            }

            return null;
        }

        #endregion

        #region OS Detection

        /// <summary>
        /// Checks if Windows Toast API is actually available.
        /// More reliable than checking OS version.
        /// </summary>
        private static bool IsToastSupported()
        {
            try
            {
                // Check Windows version via registry (more reliable)
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    if (key != null)
                    {
                        // Windows 10+ has CurrentMajorVersionNumber
                        var majorVersion = key.GetValue("CurrentMajorVersionNumber");
                        if (majorVersion is int)
                        {
                            int major = (int)majorVersion;
                            if (major >= 10)
                            {
                                // Also check build number for 1607+ (build 14393+)
                                var buildStr = key.GetValue("CurrentBuildNumber") as string;
                                if (buildStr != null && int.TryParse(buildStr, out int build))
                                {
                                    return build >= 14393;
                                }
                                return true;
                            }
                            return false;
                        }

                        // Fallback: check CurrentVersion string (for older method)
                        var versionStr = key.GetValue("CurrentVersion") as string;
                        if (versionStr != null)
                        {
                            // Windows 7 = 6.1, Windows 8 = 6.2, Windows 8.1 = 6.3
                            // These don't support Toast
                            if (versionStr.StartsWith("6."))
                            {
                                return false;
                            }
                        }
                    }
                }

                // Additional check: try to load WinRT type
                var type = Type.GetType("Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime");
                return type != null;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Tag Generation

        /// <summary>
        /// Generates a stable, collision-resistant tag for a Toast notification.
        /// Uses a truncated hex hash of EntryID to fit within the 16-char limit.
        /// Falls back to a truncated GUID if no EntryID is provided.
        /// </summary>
        private static string GenerateNotificationTag(string entryId)
        {
            if (string.IsNullOrEmpty(entryId))
                return Guid.NewGuid().ToString("N").Substring(0, MaxToastTagLength);

            // EntryID is unique per item but too long for a tag.
            // Use a stable hash that fits in 16 chars.
            // SHA256 is overkill here — we just need low collision probability
            // across a small set (active notifications, typically < 50).
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(entryId));
                // Take first 8 bytes → 16 hex chars
                return BitConverter.ToString(bytes, 0, 8).Replace("-", "");
            }
        }

        #endregion

        #region Toast Notifications (Windows 10+)

        private void ShowToastNotification(
            string mailboxName,
            string senderName,
            string subject,
            string bodyPreview,
            string entryId,
            string storeId,
            DateTime receivedTime,
            string senderPhotoPath)
        {
            var builder = new Microsoft.Toolkit.Uwp.Notifications.ToastContentBuilder();

            // App Logo Override — sender photo if available, otherwise app logo
            var logoUri = FilePathToUri(senderPhotoPath) ?? _iconUriAppLogo;
            if (logoUri != null)
            {
                var cropStyle = SettingsService.RoundAppLogo
                    ? Microsoft.Toolkit.Uwp.Notifications.ToastGenericAppLogoCrop.Circle
                    : Microsoft.Toolkit.Uwp.Notifications.ToastGenericAppLogoCrop.Default;

                builder.AddAppLogoOverride(logoUri, cropStyle);
            }

            // Line 1: Subject (title)
            builder.AddText(subject);

            // Line 2: From
            builder.AddText(string.Format(Strings.NotificationFrom, senderName));

            // Lines 3-4: Body preview (if available)
            if (!string.IsNullOrWhiteSpace(bodyPreview))
            {
                // Toast renders each AddText as a single line — flatten explicitly
                builder.AddText(TextUtils.TrimSingleLine(bodyPreview, 200));
            }

            // Attribution text (small font at bottom) - mailbox name + time
            var attributionText = string.Format("📬{0} • 🕰️{1}",
                mailboxName,
                receivedTime.ToString("dd.MM.yyyy HH:mm"));
            builder.AddAttributionText(attributionText);

            // Arguments for action handling
            builder.AddArgument("action", "open");
            builder.AddArgument("entryId", entryId ?? "");
            builder.AddArgument("storeId", storeId ?? "");

            // Category selection dropdown — populated from user's real Outlook categories.
            // Uses the existing Application reference (STA thread) instead of GetActiveObject.
            NameSpace _ns = null;
            List<CategoryService.CategoryInfo> userCategories = new List<CategoryService.CategoryInfo>();
            try
            {
                _ns = _outlookApp.GetNamespace("MAPI");
                userCategories = CategoryService.GetSelectedCategories(_ns);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[NotificationService] Failed to get Outlook categories, because " + ex.Message);
            }
            finally
            {
                ComHelper.SafeComRelease(_ns);
            }

            bool hasCategories = (userCategories.Count > 0);

            // Buttons with icons
            // Add Category button — only shown if user has categories
            if (hasCategories)
            {
                var categoryTuples = new ValueTuple<string, string>[userCategories.Count];
                for (int ci = 0; ci < userCategories.Count; ci++)
                {
                    var cat = userCategories[ci];
                    // key = real category name, display = emoji + name
                    categoryTuples[ci] = new ValueTuple<string, string>(
                        cat.Name,
                        cat.ColorEmoji + " " + cat.Name);
                }

                builder.AddComboBox(
                    "categoryCombo",
                    Strings.CategoryLabel,
                    (string)null,
                    categoryTuples);

                var btnSetMark = new Microsoft.Toolkit.Uwp.Notifications.ToastButton()
                    .SetContent(Strings.ButtonAddCategory)
                    .AddArgument("action", "addCategory")
                    .AddArgument("entryId", entryId ?? "")
                    .AddArgument("storeId", storeId ?? "");
                SetButtonIcon(btnSetMark, _iconUriSetMark);
                builder.AddButton(btnSetMark);
            }

            // Reply button
            var btnReply = new Microsoft.Toolkit.Uwp.Notifications.ToastButton()
                .SetContent(Strings.ButtonReply)
                .AddArgument("action", "reply")
                .AddArgument("entryId", entryId ?? "")
                .AddArgument("storeId", storeId ?? "");
            SetButtonIcon(btnReply, _iconUriAnswer);
            builder.AddButton(btnReply);

            // Mark as read button
            var btnMarkRead = new Microsoft.Toolkit.Uwp.Notifications.ToastButton()
                .SetContent(Strings.ButtonMarkRead)
                .AddArgument("action", "markRead")
                .AddArgument("entryId", entryId ?? "")
                .AddArgument("storeId", storeId ?? "");
            SetButtonIcon(btnMarkRead, _iconUriMarkAsRead);
            builder.AddButton(btnMarkRead);

            // Audio
            if (SettingsService.EnableSound)
            {
                builder.AddAudio(new Uri("ms-winsoundevent:Notification.Mail"));
            }
            else
            {
                builder.AddAudio(null, silent: true);
            }

            // Generate collision-resistant tag
            var tag = GenerateNotificationTag(entryId);

            // Show
            builder.Show(toast =>
            {
                toast.Tag = tag;
                toast.Group = "SharedMailbox";
            });

            // Store tag for later removal when mail is read
            if (!string.IsNullOrEmpty(entryId))
            {
                lock (_tagsLock)
                {
                    _notificationTags[entryId] = (tag, storeId);
                }
            }

            Debug.WriteLine("[NotificationService] Toast shown for: " + subject);
        }

        private void OnToastActivated(Microsoft.Toolkit.Uwp.Notifications.ToastNotificationActivatedEventArgsCompat e)
        {
            try
            {
                var args = Microsoft.Toolkit.Uwp.Notifications.ToastArguments.Parse(e.Argument);

                string action = args.Contains("action") ? args["action"] : "open";
                string entryId = args.Contains("entryId") ? args["entryId"] : null;
                string storeId = args.Contains("storeId") ? args["storeId"] : null;

                string selectedCategory = null;
                if (e.UserInput.ContainsKey("categoryCombo"))
                {
                    selectedCategory = e.UserInput["categoryCombo"] as string;
                }

                Debug.WriteLine(string.Format("[NotificationService] Toast activated: action={0}", action));

                ToastActionHandler.HandleAction(action, entryId, storeId, selectedCategory);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[NotificationService] Toast activation error: " + ex.Message);
            }
        }

        #endregion

        #region BalloonTip Fallback (Windows 7/8)

        private void ShowBalloonNotification(
            string mailboxName,
            string senderName,
            string subject,
            string bodyPreview,
            string entryId,
            string storeId,
            DateTime receivedTime)
        {
            _balloonBackend.Show(mailboxName, senderName, subject, bodyPreview, entryId, storeId, receivedTime);
        }

        #endregion

        #region Public API

        /// <summary>
        /// Shows a notification for a new email.
        /// </summary>
        public void ShowNewMailNotification(
            string mailboxName,
            string senderName,
            string subject,
            string bodyPreview,
            string entryId,
            string storeId)
        {
            ShowNewMailNotification(mailboxName, senderName, subject, bodyPreview, entryId, storeId, DateTime.Now, null);
        }

        /// <summary>
        /// Shows a notification for a new email with specific received time.
        /// </summary>
        public void ShowNewMailNotification(
            string mailboxName,
            string senderName,
            string subject,
            string bodyPreview,
            string entryId,
            string storeId,
            DateTime receivedTime,
            string senderPhotoPath)
        {
            if (_disposed)
                return;

            // Sanitize single-line fields (removes control chars, collapses spaces, truncates)
            mailboxName = TextUtils.Sanitize(mailboxName, 50);
            senderName = TextUtils.Sanitize(senderName, 100);
            subject = string.IsNullOrWhiteSpace(subject) ? Strings.NoSubject : TextUtils.Sanitize(subject, 150);

            // Try Toast first, fall back to Balloon per-call (no permanent degradation)
            if (_toastBackend != null)
            {
                try
                {
                ShowToastNotification(mailboxName, senderName, subject, bodyPreview, entryId, storeId, receivedTime, senderPhotoPath);
                    return;
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[NotificationService] Toast failed for this call, falling back to Balloon: " + ex.Message);
                    LogService.Error(Strings.LogFallbackToBalloon, ex);
                }
            }

            ShowBalloonNotification(mailboxName, senderName, subject, bodyPreview, entryId, storeId, receivedTime);
        }

        /// <summary>
        /// Simplified overload for backward compatibility.
        /// </summary>
        public void ShowNewMailNotification(string mailboxName, string senderName, string subject)
        {
            ShowNewMailNotification(mailboxName, senderName, subject, null, null, null, DateTime.Now, null);
        }

        /// <summary>
        /// Removes a notification for a mail item that has been read.
        /// Only works when Toast backend is available.
        /// </summary>
        /// <param name="entryId">The EntryID of the mail item</param>
        public void RemoveNotification(string entryId)
        {
            if (string.IsNullOrEmpty(entryId) || _toastBackend == null)
                return;

            try
            {
                string tag;
                lock (_tagsLock)
                {
                    if (!_notificationTags.TryGetValue(entryId, out var entry))
                        return;

                    tag = entry.Tag;
                    _notificationTags.Remove(entryId);
                }

                Microsoft.Toolkit.Uwp.Notifications.ToastNotificationManagerCompat.History.Remove(tag, "SharedMailbox");
                Debug.WriteLine("[NotificationService] Removed notification for read mail, tag: " + tag);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[NotificationService] Error removing notification: " + ex.Message);
            }
        }

        /// <summary>
        /// Cleans up notifications for deleted mail items.
        /// Called from MailboxWatcher.OnItemRemove on Outlook's STA thread,
        /// so we can safely use the existing _outlookApp reference.
        /// </summary>
        public void CleanupDeletedNotifications(string storeId)
        {
            if (_toastBackend == null)
                return;

            try
            {
                // Collect tracked entry IDs for this store
                List<string> trackedEntryIds;
                lock (_tagsLock)
                {
                    if (_notificationTags.Count == 0)
                        return;

                    trackedEntryIds = _notificationTags
                        .Where(kvp => string.Equals(kvp.Value.StoreId, storeId, StringComparison.OrdinalIgnoreCase))
                        .Select(kvp => kvp.Key)
                        .ToList();
                }

                if (trackedEntryIds.Count == 0)
                    return;

                // Use GetItemFromID instead of Items.Find — EntryID is not
                // a supported filter property for Find/Restrict.
                Microsoft.Office.Interop.Outlook.NameSpace ns = null;
                try
                {
                    ns = _outlookApp.GetNamespace("MAPI");

                    List<string> entryIdsToRemove = new List<string>();
                    foreach (var entryId in trackedEntryIds)
                    {
                        try
                        {
                            var item = ns.GetItemFromID(entryId, storeId);
                            ComHelper.SafeComRelease(item);
                            // Item still exists — keep the notification
                        }
                        catch
                        {
                            // GetItemFromID throws if item doesn't exist — mark for removal
                            entryIdsToRemove.Add(entryId);
                        }
                    }

                    foreach (var entryId in entryIdsToRemove)
                    {
                        RemoveNotification(entryId);
                        Debug.WriteLine("[NotificationService] Cleaned up notification for deleted mail");
                    }
                }
                finally
                {
                    ComHelper.SafeComRelease(ns);
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[NotificationService] Error cleaning up deleted notifications: " + ex.Message);
            }
        }

        #endregion

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;

            try
            {
                if (_toastBackend != null)
                {
                    _toastBackend.Dispose();
                }

                if (_balloonBackend != null)
                {
                    _balloonBackend.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[NotificationService] Dispose error: " + ex.Message);
            }
        }

        #region Inner Classes

        /// <summary>
        /// Encapsulates Toast notification initialization and cleanup.
        /// Existence of this object means Toast is available.
        /// </summary>
        private sealed class ToastBackend : IDisposable
        {
            public event Action<Microsoft.Toolkit.Uwp.Notifications.ToastNotificationActivatedEventArgsCompat> Activated;

            public ToastBackend()
            {
                // This will throw if Toast API is not properly available
                Microsoft.Toolkit.Uwp.Notifications.ToastNotificationManagerCompat.OnActivated += OnActivated;
                Debug.WriteLine("[ToastBackend] Initialized");
            }

            private void OnActivated(Microsoft.Toolkit.Uwp.Notifications.ToastNotificationActivatedEventArgsCompat e)
            {
                Activated?.Invoke(e);
            }

            public void Dispose()
            {
                try
                {
                    Microsoft.Toolkit.Uwp.Notifications.ToastNotificationManagerCompat.OnActivated -= OnActivated;
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[ToastBackend] Dispose error: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Encapsulates BalloonTip notification (NotifyIcon).
        /// Always available as a fallback.
        /// </summary>
        private sealed class BalloonBackend : IDisposable
        {
            private NotifyIcon _notifyIcon;

            // Store last notification data for click handling.
            // BalloonTip API does not support per-notification click data,
            // so this is inherently racy with rapid-fire notifications.
            private string _lastEntryId;
            private string _lastStoreId;

            public BalloonBackend()
            {
                Initialize();
            }

            private void Initialize()
            {
                try
                {
                    _notifyIcon = new NotifyIcon
                    {
                        Text = Strings.AppAttribution,
                        Visible = false
                    };

                    // Get icon from our own Outlook process (VSTO add-in runs in-process)
                    try
                    {
                        var currentProcess = Process.GetCurrentProcess();
                        _notifyIcon.Icon = Icon.ExtractAssociatedIcon(currentProcess.MainModule.FileName);
                        Debug.WriteLine("[BalloonBackend] Icon from current process: " + currentProcess.MainModule.FileName);
                    }
                    catch (System.Exception ex)
                    {
                        Debug.WriteLine("[BalloonBackend] Could not get process icon: " + ex.Message);
                        _notifyIcon.Icon = SystemIcons.Application;
                    }

                    _notifyIcon.BalloonTipClicked += OnBalloonClicked;
                    _notifyIcon.DoubleClick += OnBalloonClicked;

                    Debug.WriteLine("[BalloonBackend] Initialized");
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[BalloonBackend] Init error: " + ex.Message);
                }
            }

            public void Show(
                string mailboxName,
                string senderName,
                string subject,
                string bodyPreview,
                string entryId,
                string storeId,
                DateTime receivedTime)
            {
                if (_notifyIcon == null)
                    return;

                const int BalloonTextMaxLength = 250;
                const int BalloonTitleMaxLength = 63;

                try
                {
                    _lastEntryId = entryId;
                    _lastStoreId = storeId;

                    var title = TextUtils.TrimSingleLine(subject, BalloonTitleMaxLength);
                    var text = TextUtils.BuildMailBalloonText(
                        mailboxName: mailboxName,
                        senderName: senderName,
                        bodyPreview: bodyPreview,
                        receivedTime: receivedTime,
                        maxLength: BalloonTextMaxLength
                    );

                    _notifyIcon.Visible = true;
                    _notifyIcon.BalloonTipTitle = title;
                    _notifyIcon.BalloonTipText = text;
                    _notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
                    _notifyIcon.ShowBalloonTip(10000);

                    Debug.WriteLine("[BalloonBackend] Shown for: " + subject);
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[BalloonBackend] Show error: " + ex.Message);
                }
            }

            private void OnBalloonClicked(object sender, EventArgs e)
            {
                try
                {
                ToastActionHandler.HandleAction("open", _lastEntryId, _lastStoreId, null);
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[BalloonBackend] Click error: " + ex.Message);
                }
                finally
                {
                    if (_notifyIcon != null)
                        _notifyIcon.Visible = false;
                }
            }

            public void Dispose()
            {
                if (_notifyIcon != null)
                {
                    try
                    {
                    _notifyIcon.BalloonTipClicked -= OnBalloonClicked;
                        _notifyIcon.DoubleClick -= OnBalloonClicked;
                    _notifyIcon.Visible = false;
                    _notifyIcon.Dispose();
                    }
                    catch (System.Exception ex)
                    {
                        Debug.WriteLine("[BalloonBackend] Dispose error: " + ex.Message);
                    }
                    _notifyIcon = null;
                }
            }
        }

        #endregion
    }
}
