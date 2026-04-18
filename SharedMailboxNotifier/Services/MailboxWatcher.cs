using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SharedMailboxNotifier.Resources;

namespace SharedMailboxNotifier.Services
{
    /// <summary>
    /// Watches a single mailbox's Inbox folder for new items.
    /// </summary>
    public sealed class MailboxWatcher : IDisposable
    {
        private Folder _inboxFolder; // Owned - released in Dispose()
        private Items _inboxItems;
        private readonly string _mailboxDisplayName;
        private readonly string _storeId;
        private readonly NotificationService _notificationService;
        private bool _disposed;

        public event EventHandler<NewMailEventArgs> NewMailReceived;

        public string MailboxName { get { return _mailboxDisplayName; } }

        public string StoreId { get { return _storeId; } }

        public bool IsActive { get { return _inboxItems != null && !_disposed; } }

        /// <summary>
        /// Creates a watcher for the given Inbox folder.
        /// Takes ownership of inboxFolder - caller must NOT release it.
        /// </summary>
        public MailboxWatcher(Folder inboxFolder, string mailboxDisplayName, NotificationService notificationService)
        {
            if (inboxFolder == null)
                throw new ArgumentNullException("inboxFolder");
            if (mailboxDisplayName == null)
                throw new ArgumentNullException("mailboxDisplayName");
            if (notificationService == null)
                throw new ArgumentNullException("notificationService");

            _inboxFolder = inboxFolder;
            _mailboxDisplayName = mailboxDisplayName;
            _notificationService = notificationService;

            // Get StoreID for later use in action handling
            try
            {
                _storeId = inboxFolder.StoreID;
            }
            catch
            {
                _storeId = null;
            }

            Initialize();
        }

        private void Initialize()
        {
            try
            {
                // IMPORTANT: Hold reference to Items to prevent GC
                _inboxItems = _inboxFolder.Items;
                _inboxItems.ItemAdd += OnItemAdd;
                _inboxItems.ItemChange += OnItemChange;
                _inboxItems.ItemRemove += OnItemRemove;

                Debug.WriteLine("[MailboxWatcher] Started watching: " + _mailboxDisplayName);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[MailboxWatcher] Failed to initialize for " + _mailboxDisplayName + ": " + ex.Message);
                throw;
            }
        }

        private void OnItemChange(object item)
        {
            if (_disposed)
                return;

            MailItem mailItem = null;

            try
            {
                mailItem = item as MailItem;
                if (mailItem != null)
                {
                    // Check if mail was marked as read
                    if (!mailItem.UnRead)
                    {
                        // Mail is now read - remove notification
                        var entryId = mailItem.EntryID;
                        if (!string.IsNullOrEmpty(entryId))
                        {
                            _notificationService.RemoveNotification(entryId);
                        }
                    }
                }
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x8004010F))
            {
                // MAPI_E_NOT_FOUND - item is being moved/deleted, COM object
                // is no longer valid. Safe to ignore: if the item moved to
                // another folder, that folder's watcher will handle it.
                Debug.WriteLine("[MailboxWatcher] Item moved or deleted during change event (MAPI_E_NOT_FOUND)");
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[MailboxWatcher] Error processing item change: " + ex.Message);
            }
            finally
            {
                ComHelper.SafeComRelease(mailItem);
            }
        }

        private void OnItemRemove()
        {
            if (_disposed)
                return;

            try
            {
                // Item was removed - clean up notifications for deleted items
                _notificationService.CleanupDeletedNotifications(_storeId);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[MailboxWatcher] Error processing item remove: " + ex.Message);
            }
        }

        private void OnItemAdd(object item)
        {
            if (_disposed)
                return;

            MailItem mailItem = null;

            try
            {
                mailItem = item as MailItem;
                if (mailItem != null && mailItem.UnRead)
                {
                    ProcessNewMail(mailItem);
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[MailboxWatcher] Error processing item in " + _mailboxDisplayName + ": " + ex.Message);
            }
            finally
            {
                ComHelper.SafeComRelease(mailItem);
            }
        }
        private string GetSenderPhotoFromMAPI(MailItem mailItem)
        {
            AddressEntry sender = null;
            try
            {
                sender = mailItem.Sender;
                if (sender == null)
                    return null;

                // MAPI property tags for thumbnail photo
                var proptags = new[]
                {
                    "http://schemas.microsoft.com/mapi/proptag/0x8C9E0102",  // PR_EMS_AB_THUMBNAIL_PHOTO
                    "http://schemas.microsoft.com/mapi/proptag/0x8C980102",  // PR_THUMBNAIL_PHOTO (alternate)
                };

                foreach (var proptag in proptags)
                {
                    try
                    {
                        var photoBytes = sender.PropertyAccessor.GetProperty(proptag) as byte[];

                        if (photoBytes != null && photoBytes.Length > 0)
                        {
                            Debug.WriteLine("[MailboxWatcher] Found MAPI photo, size: " + photoBytes.Length);

                            // Save to temp file
                            var tempFolder = Path.Combine(Path.GetTempPath(), "SharedMailboxNotifier");
                            if (!Directory.Exists(tempFolder))
                            {
                                Directory.CreateDirectory(tempFolder);
                            }

                            // Use stable sanitized file name instead of GetHashCode() to avoid collisions
                            var safeFileName = TextUtils.SanitizeForFileName(mailItem.SenderEmailAddress ?? "unknown");
                            var photoPath = Path.Combine(tempFolder, "mapi_" + safeFileName + ".jpg");

                            if (!File.Exists(photoPath))
                            {
                                File.WriteAllBytes(photoPath, photoBytes);
                                Debug.WriteLine("[MailboxWatcher] Saved MAPI photo: " + photoPath);
                            }

                            return photoPath;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Debug.WriteLine("[MailboxWatcher] MAPI property " + proptag + " failed: " + ex.Message);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[MailboxWatcher] GetSenderPhotoFromMAPI error: " + ex.Message);
            }
            finally
            {
                ComHelper.SafeComRelease(sender);
            }

            return null;
        }

        /// <summary>
        /// Tries to get contact photo from sender and save to temp file.
        /// </summary>
        private string GetSenderPhoto(MailItem mailItem)
        {
            // Try MAPI first (AD photo via Exchange)
            try
            {
                var photo = GetSenderPhotoFromMAPI(mailItem);
                if (photo != null)
                    return photo;
            }
            catch (System.Exception ex) { Debug.WriteLine("[MailboxWatcher] Failed to get photo from AD, because " + ex.Message); }

            // Local Contacts search only if user opted in (can be slow)
            if (!SettingsService.SearchContactPhotos)
                return null;

            NameSpace ns = null;
            Folder contactsFolder = null;
            Items contactItems = null;
            ContactItem contact = null;
            try
            {
                var senderEmail = mailItem.SenderEmailAddress;
                var senderName = mailItem.SenderName;

                if (string.IsNullOrEmpty(senderEmail) && string.IsNullOrEmpty(senderName))
                    return null;

                // Get MAPI namespace
                ns = mailItem.Application.GetNamespace("MAPI");

                try
                {
                    // Try to find contact in default Contacts folder
                    contactsFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderContacts) as Folder;
                    if (contactsFolder == null)
                        return null;

                    contactItems = contactsFolder.Items;

                    // Search by email first
                    if (!string.IsNullOrEmpty(senderEmail))
                    {
                        // Try Email1, Email2, Email3
                        var filter = string.Format(
                            "[Email1Address] = '{0}' OR [Email2Address] = '{0}' OR [Email3Address] = '{0}'",
                            senderEmail.Replace("'", "''"));

                        try
                        {
                            contact = contactItems.Find(filter) as ContactItem;
                        }
                        catch
                        {
                            // Filter might fail for some email formats, try simpler search
                        }
                    }

                    // If not found by email, try by name
                    if (contact == null && !string.IsNullOrEmpty(senderName))
                    {
                        try
                        {
                            var filter = string.Format("[FullName] = '{0}'", senderName.Replace("'", "''"));
                            contact = contactItems.Find(filter) as ContactItem;
                        }
                        catch (System.Exception ex)
                        {
                            Debug.WriteLine("[MailboxWatcher] [unknown] Suppressed: " + ex.Message);
                        }
                    }

                    if (contact != null && contact.HasPicture)
                    {
                        // Save photo to temp file
                        var tempFolder = Path.Combine(Path.GetTempPath(), "SharedMailboxNotifier");
                        if (!Directory.Exists(tempFolder))
                        {
                            Directory.CreateDirectory(tempFolder);
                        }

                        // Use stable sanitized file name instead of GetHashCode()
                        var safeFileName = TextUtils.SanitizeForFileName(senderEmail ?? senderName ?? "unknown");
                        var photoPath = Path.Combine(tempFolder, "contact_" + safeFileName + ".jpg");

                        // Only extract if not already cached
                        if (!File.Exists(photoPath))
                        {
                            Attachments attachments = contact.Attachments;
                            try
                            {
                                for (int i = 1; i <= attachments.Count; i++)
                                {
                                    Attachment att = null;
                                    try
                                    {
                                        att = attachments[i];
                                        if (att.FileName == "ContactPicture.jpg" ||
                                            att.FileName.StartsWith("ContactPhoto") ||
                                            att.FileName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase))
                                        {
                                            att.SaveAsFile(photoPath);
                                            Debug.WriteLine("[MailboxWatcher] Saved contact photo: " + photoPath);
                                            break;
                                        }
                                    }
                                    finally
                                    {
                                        ComHelper.SafeComRelease(att);
                                    }
                                }
                            }
                            finally
                            {
                                ComHelper.SafeComRelease(attachments);
                            }
                        }

                        if (File.Exists(photoPath))
                        {
                            return photoPath;
                        }
                    }
                }
                finally
                {
                    ComHelper.SafeComRelease(contact);
                    ComHelper.SafeComRelease(contactItems);
                    ComHelper.SafeComRelease(contactsFolder);
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[MailboxWatcher] Error getting sender photo: " + ex.Message);
            }
            finally
            {
                ComHelper.SafeComRelease(ns);
            }

            return null;
        }

        private void ProcessNewMail(MailItem mailItem)
        {
            try
            {
                var senderName = GetSenderName(mailItem);
                var subject = mailItem.Subject ?? string.Empty;
                var entryId = mailItem.EntryID;
                var bodyPreview = GetBodyPreview(mailItem, 500);
                var receivedTime = mailItem.ReceivedTime;
                var senderPhoto = GetSenderPhoto(mailItem);

                // Show toast notification with all data
                _notificationService.ShowNewMailNotification(
                    _mailboxDisplayName,
                    senderName,
                    subject,
                    bodyPreview,
                    entryId,
                    _storeId,
                    receivedTime,
                    senderPhoto);

                // Raise event for additional handling if needed
                var handler = NewMailReceived;
                if (handler != null)
                {
                    handler(this, new NewMailEventArgs
                    {
                        MailboxName = _mailboxDisplayName,
                        SenderName = senderName,
                        Subject = subject,
                        BodyPreview = bodyPreview,
                        ReceivedTime = mailItem.ReceivedTime,
                        EntryId = entryId,
                        StoreId = _storeId
                    });
                }

                Debug.WriteLine(string.Format("[MailboxWatcher] New mail in {0} from {1}: {2}",
                    _mailboxDisplayName, senderName, subject));
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[MailboxWatcher] Error showing notification: " + ex.Message);
                LogService.Error(Strings.LogErrorNewMail + _mailboxDisplayName, ex);
            }
        }

        /// <summary>
        /// Extracts a normalized body preview from a mail item.
        /// Preserves paragraph structure (single newlines), removes noise
        /// (tabs, triple+ newlines, leading/trailing whitespace).
        /// Does NOT flatten to single line — each backend decides how to render.
        /// </summary>
        private static string GetBodyPreview(MailItem mailItem, int maxLength)
        {
            try
            {
                string body = null;
                try
                {
                    body = mailItem.Body;
                }
                catch
                {
                    // Body not available (not downloaded in Cached Mode)
                    return null;
                }

                if (string.IsNullOrWhiteSpace(body))
                    return null;

                // Normalize whitespace but keep paragraph breaks
                body = TextUtils.NormalizeMultilineText(body);

                // Coarse trim — fine trimming is done per-backend
                if (body.Length > maxLength)
                {
                    body = TextUtils.TrimMultilinePreservingWords(body, maxLength);
                }

                return body;
            }
            catch
            {
                return null;
            }
        }

        private static string GetSenderName(MailItem mailItem)
        {
            AddressEntry sender = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(mailItem.SenderName))
                    return mailItem.SenderName;

                if (!string.IsNullOrWhiteSpace(mailItem.SenderEmailAddress))
                    return mailItem.SenderEmailAddress;

                sender = mailItem.Sender;
                if (sender != null)
                {
                    if (!string.IsNullOrWhiteSpace(sender.Name))
                        return sender.Name;
                    if (!string.IsNullOrWhiteSpace(sender.Address))
                        return sender.Address;
                }

                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
            finally
            {
                ComHelper.SafeComRelease(sender);
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;

            if (_inboxItems != null)
            {
                try
                {
                    _inboxItems.ItemAdd -= OnItemAdd;
                    _inboxItems.ItemChange -= OnItemChange;
                    _inboxItems.ItemRemove -= OnItemRemove;
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[MailboxWatcher] [Dispose] Suppressed: " + ex.Message);
                }

                ComHelper.SafeComRelease(_inboxItems);
                _inboxItems = null;
            }

            ComHelper.SafeComRelease(_inboxFolder);
            _inboxFolder = null;

            Debug.WriteLine("[MailboxWatcher] Stopped watching: " + _mailboxDisplayName);
        }

        /// <summary>
        /// Cleans up cached sender photo files from the temp folder.
        /// Call once during add-in shutdown.
        /// </summary>
        public static void CleanupTempPhotos()
        {
            try
            {
                var tempFolder = Path.Combine(Path.GetTempPath(), "SharedMailboxNotifier");
                if (Directory.Exists(tempFolder))
                {
                    var files = Directory.GetFiles(tempFolder, "*.jpg");
                    foreach (var file in files)
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch (System.Exception ex)
                        {
                            Debug.WriteLine("[MailboxWatcher] [CleanupTempPhotos] Suppressed: " + ex.Message);
                        }
                    }
                    Debug.WriteLine("[MailboxWatcher] Cleaned up " + files.Length + " cached photos");
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[MailboxWatcher] Error cleaning up temp photos: " + ex.Message);
            }
        }
    }
}
