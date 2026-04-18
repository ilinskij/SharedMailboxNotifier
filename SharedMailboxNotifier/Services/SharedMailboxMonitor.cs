using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SharedMailboxNotifier.Resources;

namespace SharedMailboxNotifier.Services
{
    /// <summary>
    /// Main service that monitors shared mailboxes for incoming mail.
    /// Behavior depends on settings: can monitor only shared or all mailboxes.
    /// Dynamically reacts to stores being added or removed at runtime.
    /// </summary>
    /// <remarks>
    /// Threading model:
    ///
    /// This class operates primarily on Outlook's main STA thread. All Outlook COM
    /// events (StoreAdd, BeforeStoreRemove, ItemAdd, etc.) are marshaled by COM to
    /// the STA thread that created the objects — i.e. Outlook's UI thread. This means
    /// Initialize(), OnStoreAdd(), OnBeforeStoreRemove(), and Dispose() are never
    /// called concurrently by Outlook itself.
    ///
    /// The _lock exists as a defensive measure for two reasons:
    /// 1. MonitoredMailboxes property may be read from a non-STA context (e.g. UI
    ///    binding, logging, or diagnostics from a background thread).
    /// 2. Future-proofing: if any consumer ever calls Dispose() from a thread other
    ///    than the Outlook STA (e.g. a finalizer or AppDomain unload handler),
    ///    the lock prevents a torn read of _watchers during enumeration.
    ///
    /// The lock is NOT here because Outlook fires COM events from multiple threads —
    /// it doesn't. If you're reviewing this code and thinking "this needs
    /// ConcurrentDictionary / ReaderWriterLockSlim / full thread safety", please
    /// re-read the above before refactoring.
    /// </remarks>
    public sealed class SharedMailboxMonitor : IDisposable
    {
        private readonly Application _outlookApp;
        private readonly NotificationService _notificationService;
        private readonly List<MailboxWatcher> _watchers;
        private NameSpace _namespace;
        private Stores _stores; // IMPORTANT: hold reference to prevent GC (same pattern as Items)
        private volatile bool _disposed;
        private bool _initialized;

        // Protects _watchers and _watchedStoreIds. See <remarks> on the class
        // for why this lock exists and why it's intentionally lightweight.
        private readonly object _lock = new object();

        // Store IDs of user's own accounts to identify them
        private HashSet<string> _ownAccountStoreIds;

        // Track which StoreIDs we've already set up watchers for (to avoid duplicates)
        private readonly HashSet<string> _watchedStoreIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public IReadOnlyList<string> MonitoredMailboxes
        {
            get
            {
                lock (_lock)
                {
                    return _watchers
                        .Where(w => w.IsActive)
                        .Select(w => w.MailboxName)
                        .ToList();
                }
            }
        }

        public event EventHandler<MonitorStatusEventArgs> StatusChanged;

        public SharedMailboxMonitor(Application outlookApp)
        {
            if (outlookApp == null)
                throw new ArgumentNullException("outlookApp");

            _outlookApp = outlookApp;
            _notificationService = new NotificationService(outlookApp);
            _watchers = new List<MailboxWatcher>();
            _ownAccountStoreIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        public void Initialize()
        {
            if (_initialized)
            {
                Debug.WriteLine("[SharedMailboxMonitor] Already initialized");
                return;
            }

            try
            {
                _namespace = _outlookApp.GetNamespace("MAPI");

                // Collect all user's own account Store IDs
                CollectOwnAccountStoreIds();

                Debug.WriteLine("[SharedMailboxMonitor] Monitor mode: " + SettingsService.MonitorMode);
                Debug.WriteLine("[SharedMailboxMonitor] Own accounts: " + _ownAccountStoreIds.Count);

                // Hold reference to Stores collection to prevent GC
                // (same pattern as MailboxWatcher holds Items reference)
                _stores = _namespace.Stores;

                Debug.WriteLine("[SharedMailboxMonitor] Total stores: " + _stores.Count);

                int monitoredCount = 0;

                // Iterate through all stores (mailboxes)
                for (int i = 1; i <= _stores.Count; i++)
                {
                    Store store = null;
                    try
                    {
                        store = _stores[i];
                        if (TryWatchStore(store))
                        {
                            monitoredCount++;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Debug.WriteLine("[SharedMailboxMonitor] Error processing store: " + ex.Message);
                    }
                    finally
                    {
                        ComHelper.SafeComRelease(store);
                    }
                }

                // Subscribe to store lifecycle events
                _stores.StoreAdd += OnStoreAdd;
                _stores.BeforeStoreRemove += OnBeforeStoreRemove;

                _initialized = true;

                bool sharedOnly = (SettingsService.MonitorMode == SettingsService.MonitorModeEnum.SharedMailboxesOnly);
                string modeDescription = sharedOnly ? "shared" : "all";
                OnStatusChanged(new MonitorStatusEventArgs
                {
                    IsRunning = true,
                    SharedMailboxCount = monitoredCount,
                    Message = string.Format("Monitoring {0} mailbox(es) ({1} mode)", monitoredCount, modeDescription)
                });

                Debug.WriteLine(string.Format("[SharedMailboxMonitor] Initialization complete. Watching {0} mailboxes.", monitoredCount));
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SharedMailboxMonitor] Initialization failed: " + ex.Message);

                OnStatusChanged(new MonitorStatusEventArgs
                {
                    IsRunning = false,
                    SharedMailboxCount = 0,
                    Message = "Failed to initialize: " + ex.Message
                });

                throw;
            }
        }

        /// <summary>
        /// Evaluates a single store and creates a watcher for it if appropriate.
        /// Returns true if a watcher was created.
        /// The caller is responsible for releasing the store COM object.
        /// </summary>
        private bool TryWatchStore(Store store)
        {
            Folder rootFolder = null;
            try
            {
                var displayName = store.DisplayName;
                var storeId = store.StoreID;

                lock (_lock)
                {
                    // Already watching this store?
                    if (_watchedStoreIds.Contains(storeId))
                    {
                        Debug.WriteLine("[SharedMailboxMonitor] Already watching: " + displayName);
                        return false;
                    }
                }

                bool sharedOnly = (SettingsService.MonitorMode == SettingsService.MonitorModeEnum.SharedMailboxesOnly);
                bool isOwnAccount = IsOwnAccount(store);

                // In "Shared Only" mode, skip own accounts
                if (sharedOnly && isOwnAccount)
                {
                    Debug.WriteLine("[SharedMailboxMonitor] Skipping own account (shared-only mode): " + displayName);
                    return false;
                }

                // Always skip special stores (public folders, etc.)
                if (IsSpecialStore(store))
                {
                    Debug.WriteLine("[SharedMailboxMonitor] Skipping special store: " + displayName);
                    return false;
                }

                // Find Inbox folder in this store
                rootFolder = store.GetRootFolder() as Folder;
                if (rootFolder == null)
                {
                    Debug.WriteLine("[SharedMailboxMonitor] Cannot get root folder for: " + displayName);
                    return false;
                }

                var inbox = FolderNameResolver.FindInboxFolder(rootFolder);
                if (inbox == null)
                {
                    Debug.WriteLine("[SharedMailboxMonitor] No Inbox found in: " + displayName);
                    return false;
                }

                // MailboxWatcher takes ownership of inbox — do NOT release it here
                var watcher = new MailboxWatcher(inbox, displayName, _notificationService);

                lock (_lock)
                {
                    _watchers.Add(watcher);
                    _watchedStoreIds.Add(storeId);
                }

                string mailboxType = isOwnAccount ? "own" : "shared";
                Debug.WriteLine(string.Format("[SharedMailboxMonitor] Now watching {0} mailbox: {1}", mailboxType, displayName));
                return true;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SharedMailboxMonitor] Error in TryWatchStore: " + ex.Message);
                return false;
            }
            finally
            {
                ComHelper.SafeComRelease(rootFolder);
            }
        }

        /// <summary>
        /// Called when a new store is added to the Outlook profile at runtime.
        /// Examples: user connects a new shared mailbox, admin pushes one via GPO.
        /// </summary>
        private void OnStoreAdd(Store store)
        {
            if (_disposed)
                return;

            try
            {
                var displayName = store.DisplayName;
                Debug.WriteLine("[SharedMailboxMonitor] StoreAdd event: " + displayName);

                // Re-check own accounts (the new store might be a new personal account)
                CollectOwnAccountStoreIds();

                if (TryWatchStore(store))
                {
                    int activeCount;
                    lock (_lock)
                    {
                        activeCount = _watchers.Count(w => w.IsActive);
                    }
                    OnStatusChanged(new MonitorStatusEventArgs
                    {
                        IsRunning = true,
                        SharedMailboxCount = activeCount,
                        Message = string.Format("Store added: {0}. Now monitoring {1} mailbox(es).", displayName, activeCount)
                    });

                    LogService.Info(string.Format(Strings.LogStoreAdded, displayName));
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SharedMailboxMonitor] Error handling StoreAdd: " + ex.Message);
            }
            // Note: do NOT release the store COM object here — Outlook owns this reference
        }

        /// <summary>
        /// Called before a store is removed from the Outlook profile at runtime.
        /// We must dispose the watcher before the store becomes invalid.
        /// </summary>
        private void OnBeforeStoreRemove(Store store, ref bool Cancel)
        {
            if (_disposed)
                return;

            try
            {
                var storeId = store.StoreID;
                var displayName = store.DisplayName;
                Debug.WriteLine("[SharedMailboxMonitor] BeforeStoreRemove event: " + displayName);

                // Find and dispose the watcher for this store
                MailboxWatcher watcherToRemove;
                lock (_lock)
                {
                    watcherToRemove = _watchers.FirstOrDefault(w => 
                        w.StoreId != null && w.StoreId.Equals(storeId, StringComparison.OrdinalIgnoreCase));
                }

                if (watcherToRemove != null)
                {
                    try
                    {
                        watcherToRemove.Dispose();
                    }
                    catch (System.Exception ex)
                    {
                        Debug.WriteLine("[SharedMailboxMonitor] Error disposing watcher for removed store: " + ex.Message);
                    }

                    int activeCount;
                    lock (_lock)
                    {
                        _watchers.Remove(watcherToRemove);
                        _watchedStoreIds.Remove(storeId);
                        activeCount = _watchers.Count(w => w.IsActive);
                    }
                    OnStatusChanged(new MonitorStatusEventArgs
                    {
                        IsRunning = true,
                        SharedMailboxCount = activeCount,
                        Message = string.Format("Store removed: {0}. Now monitoring {1} mailbox(es).", displayName, activeCount)
                    });

                    LogService.Info(string.Format(Strings.LogStoreRemoved, displayName));
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SharedMailboxMonitor] Error handling BeforeStoreRemove: " + ex.Message);
            }
            // Note: do NOT release the store COM object here — Outlook owns this reference
        }

        /// <summary>
        /// Collects Store IDs for all accounts that belong to the current user.
        /// These are "own" mailboxes that Outlook handles notifications for.
        /// Rebuilds the set from scratch each time to avoid retaining stale IDs
        /// from accounts that were removed during the session.
        /// </summary>
        private void CollectOwnAccountStoreIds()
        {
            var newSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Method 1: All registered Accounts
                Accounts accounts = _namespace.Accounts;
                try
                {
                    for (int i = 1; i <= accounts.Count; i++)
                    {
                        Account account = null;
                        Store store = null;
                        try
                        {
                            account = accounts[i];
                            store = account.DeliveryStore;
                            if (store != null)
                            {
                                newSet.Add(store.StoreID);
                                Debug.WriteLine("[SharedMailboxMonitor] Own account found: " + account.DisplayName);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            Debug.WriteLine("[SharedMailboxMonitor] [CollectOwnAccountStoreIds] Suppressed: " + ex.Message);
                        }
                        finally
                        {
                            ComHelper.SafeComRelease(store);
                            ComHelper.SafeComRelease(account);
                        }
                    }
                }
                finally
                {
                    ComHelper.SafeComRelease(accounts);
                }

                // Method 2: Default Store is always "own"
                Store defaultStore = null;
                try
                {
                    defaultStore = _namespace.DefaultStore;
                    if (defaultStore != null)
                    {
                        newSet.Add(defaultStore.StoreID);
                    }
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[SharedMailboxMonitor] [CollectOwnAccountStoreIds] Suppressed: " + ex.Message);
                }
                finally
                {
                    ComHelper.SafeComRelease(defaultStore);
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SharedMailboxMonitor] Error collecting own accounts, because " + ex.Message);
            }

            // Atomic reference swap — safe for concurrent readers on other threads
            // (e.g. MonitoredMailboxes property). Old set becomes garbage.
            _ownAccountStoreIds = newSet;
        }

        /// <summary>
        /// Determines if a store belongs to the current user's own accounts.
        /// </summary>
        private bool IsOwnAccount(Store store)
        {
            try
            {
                // Check 1: Is this store in our list of own account stores?
                if (_ownAccountStoreIds.Contains(store.StoreID))
                {
                    return true;
                }

                // Check 2: For Exchange, check store type
                try
                {
                    var exchangeStoreType = store.ExchangeStoreType;

                    // Primary mailbox is user's own
                    if (exchangeStoreType == OlExchangeStoreType.olPrimaryExchangeMailbox)
                    {
                        return true;
                    }

                    // Delegate stores might need notification (user acts on behalf)
                    // Additional mailboxes are shared - don't skip
                    // olAdditionalExchangeMailbox = shared mailbox
                    // olDelegateExchangeMailbox = delegate access
                }
                catch
                {
                    // Not an Exchange store, or property not available
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Determines if a store is a special/system store that should be skipped.
        /// </summary>
        private bool IsSpecialStore(Store store)
        {
            try
            {
                // Check Exchange store type for special stores
                try
                {
                    var exchangeStoreType = store.ExchangeStoreType;

                    // Skip public folders
                    if (exchangeStoreType == OlExchangeStoreType.olExchangePublicFolder)
                    {
                        return true;
                    }
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[SharedMailboxMonitor] [IsSpecialStore] Suppressed: " + ex.Message);
                }

                // Check display name for known special folders
                var displayName = store.DisplayName;
                if (string.IsNullOrWhiteSpace(displayName))
                    return true;

                var specialNames = new[]
                {
                    "Public Folders",
                    "Общие папки",
                    "Публічні папки",
                    "Öffentliche Ordner",
                    "Dossiers publics",
                    "Carpetas públicas",
                    "Internet Calendars",
                    "Интернет-календари",
                    "Інтернет-календарі",
                    "SharePoint Lists",
                    "Списки SharePoint"
                };

                if (specialNames.Any(name => displayName.Equals(name, StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }

                // Skip Online Archive mailboxes (Exchange In-Place Archive).
                // DisplayName format: "Online Archive - <user>" (classic Outlook)
                // or "In-Place Archive - <user>" (New Outlook / OWA).
                // Outlook localizes the display name, so we need known translations.
                // These are read-only archive stores — no new mail arrives there.
                var archivePrefixes = new[]
                {
                    "Online Archive",       // English (classic Outlook)
                    "In-Place Archive",     // English (New Outlook / OWA)
                    "Сетевой архив",        // Russian
                    "Onlinearchiv",         // German
                    "Archive en ligne",     // French
                    "Archivo en línea",     // Spanish
                };

                if (archivePrefixes.Any(prefix => displayName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                {
                    Debug.WriteLine("[SharedMailboxMonitor] Skipping archive store: " + displayName);
                    return true;
                }

                // Skip PST archives (local files, not shared mailboxes)
                // Note: OST files are Exchange cache and may belong to shared mailboxes — don't skip them
                if (store.IsDataFileStore)
                {
                    var filePath = store.FilePath;
                    if (!string.IsNullOrEmpty(filePath) &&
                        filePath.EndsWith(".pst", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private void OnStatusChanged(MonitorStatusEventArgs args)
        {
            var handler = StatusChanged;
            if (handler != null)
            {
                handler(this, args);
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;

            // Unsubscribe from store events first
            if (_stores != null)
            {
                try
                {
                    _stores.StoreAdd -= OnStoreAdd;
                    _stores.BeforeStoreRemove -= OnBeforeStoreRemove;
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[SharedMailboxMonitor] [Dispose] Suppressed: " + ex.Message);
                }
            }

            // Snapshot watchers under lock, then dispose outside lock
            // (Dispose may trigger COM callbacks — don't hold lock during COM calls)
            List<MailboxWatcher> watchersSnapshot;
            lock (_lock)
            {
                watchersSnapshot = new List<MailboxWatcher>(_watchers);
                _watchers.Clear();
                _watchedStoreIds.Clear();
            }

            foreach (var watcher in watchersSnapshot)
            {
                try
                {
                    watcher.Dispose();
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[SharedMailboxMonitor] Error disposing watcher: " + ex.Message);
                }
            }

            if (_notificationService != null)
            {
                _notificationService.Dispose();
            }

            // Release COM objects
            ComHelper.SafeComRelease(_stores);
            _stores = null;

            ComHelper.SafeComRelease(_namespace);
            _namespace = null;

            _ownAccountStoreIds.Clear();

            // Clean up cached sender photos from temp folder
            MailboxWatcher.CleanupTempPhotos();

            OnStatusChanged(new MonitorStatusEventArgs
            {
                IsRunning = false,
                SharedMailboxCount = 0,
                Message = "Monitor stopped"
            });

            Debug.WriteLine("[SharedMailboxMonitor] Disposed");
        }
    }
}
