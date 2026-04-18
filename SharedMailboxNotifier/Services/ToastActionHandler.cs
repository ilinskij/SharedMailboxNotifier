using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;
using SharedMailboxNotifier.Resources;

namespace SharedMailboxNotifier.Services
{
    /// <summary>
    /// Handles actions from toast notification button clicks.
    /// </summary>
    public static class ToastActionHandler
    {
        /// <summary>
        /// Processes an action from a toast notification.
        /// </summary>
        /// <param name="action">Action type: open, reply, markRead, addCategory</param>
        /// <param name="entryId">Outlook EntryID of the mail item</param>
        /// <param name="storeId">Outlook StoreID of the mailbox</param>
        /// <param name="category">Selected category (for addCategory action)</param>
        public static void HandleAction(string action, string entryId, string storeId, string category)
        {
            if (string.IsNullOrEmpty(action))
                return;

            try
            {
                Debug.WriteLine("[ToastActionHandler] Handling action: " + action);

                switch (action.ToLowerInvariant())
                {
                    case "open":
                        OpenMailItem(entryId, storeId);
                        break;

                    case "reply":
                        ReplyToMailItem(entryId, storeId);
                        break;

                    case "markread":
                        MarkAsRead(entryId, storeId);
                        break;

                    case "addcategory":
                        AddCategory(entryId, storeId, category);
                        break;

                    default:
                        Debug.WriteLine("[ToastActionHandler] Unknown action: " + action);
                        BringOutlookToForeground();
                        break;
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[ToastActionHandler] Error: " + ex.Message);
                LogService.Error(Strings.LogErrorToastAction + action, ex);
            }
        }

        /// <summary>
        /// Opens the mail item in Outlook.
        /// </summary>
        private static void OpenMailItem(string entryId, string storeId)
        {
            var mailItem = GetMailItem(entryId, storeId);
            if (mailItem != null)
            {
                try
                {
                    mailItem.Display(false);
                    BringOutlookToForeground();
                }
                finally
                {
                    ComHelper.SafeComRelease(mailItem);
                }
            }
            else
            {
                // Fallback: just bring Outlook to front
                BringOutlookToForeground();
            }
        }

        /// <summary>
        /// Opens a reply window for the mail item.
        /// </summary>
        private static void ReplyToMailItem(string entryId, string storeId)
        {
            var mailItem = GetMailItem(entryId, storeId);
            if (mailItem != null)
            {
                MailItem reply = null;
                try
                {
                    reply = mailItem.Reply();
                    reply.Display(false);
                    BringOutlookToForeground();
                }
                finally
                {
                    ComHelper.SafeComRelease(reply);
                    ComHelper.SafeComRelease(mailItem);
                }
            }
            else
            {
                BringOutlookToForeground();
            }
        }

        /// <summary>
        /// Marks the mail item as read.
        /// </summary>
        private static void MarkAsRead(string entryId, string storeId)
        {
            var mailItem = GetMailItem(entryId, storeId);
            if (mailItem != null)
            {
                try
                {
                    if (mailItem.UnRead)
                    {
                        mailItem.UnRead = false;
                        mailItem.Save();
                        Debug.WriteLine("[ToastActionHandler] Marked as read");
                    }
                }
                finally
                {
                    ComHelper.SafeComRelease(mailItem);   
                }
            }
        }

        /// <summary>
        /// Adds a category on the mail item.
        /// The categoryName comes directly from the user's Outlook categories
        /// (read by CategoryService), so it's always a valid, existing category name.
        /// </summary>
        private static void AddCategory(string entryId, string storeId, string categoryName)
        {
            if (string.IsNullOrEmpty(categoryName))
            {
                Debug.WriteLine("[ToastActionHandler] No category selected");
                return;
            }

            var mailItem = GetMailItem(entryId, storeId);
            if (mailItem != null)
            {
                try
                {
                    var existing = mailItem.Categories ?? "";

                    // Check for duplicates
                    var categories = existing.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var cat in categories)
                    {
                        if (cat.Equals(categoryName, StringComparison.OrdinalIgnoreCase))
                        {
                            Debug.WriteLine("[ToastActionHandler] Category already present: " + categoryName);
                            return;
                        }
                    }

                    mailItem.Categories = string.IsNullOrEmpty(existing) ? categoryName : existing + ", " + categoryName;
                    mailItem.Save();

                    Debug.WriteLine("[ToastActionHandler] Added category: " + categoryName);
                }
                finally
                {
                    ComHelper.SafeComRelease(mailItem);
                }
            }
        }

        /// <summary>
        /// Retrieves a MailItem from Outlook by EntryID.
        /// </summary>
        private static MailItem GetMailItem(string entryId, string storeId)
        {
            if (string.IsNullOrEmpty(entryId))
                return null;

            Application outlookApp = null;
            NameSpace ns = null;
            try
            {
                // Get Outlook application instance
                try
                {
                    outlookApp = (Application)Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    // Outlook not running
                    Debug.WriteLine("[ToastActionHandler] Outlook not running");
                    return null;
                }

                if (outlookApp == null)
                    return null;

                ns = outlookApp.GetNamespace("MAPI");

                object item;
                if (!string.IsNullOrEmpty(storeId))
                {
                    item = ns.GetItemFromID(entryId, storeId);
                }
                else
                {
                    item = ns.GetItemFromID(entryId);
                }

                var mailItem = item as MailItem;
                if (mailItem == null && item != null)
                {
                    // Item exists but is not a MailItem (meeting request, etc.) — release it
                    ComHelper.SafeComRelease(item);
                }

                return mailItem;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[ToastActionHandler] Error getting mail item: " + ex.Message);
                return null;
            }
            finally
            {
                ComHelper.SafeComRelease(ns);
                // Release the COM reference from GetActiveObject
                // (this decrements the ref count, doesn't close Outlook)
                ComHelper.SafeComRelease(outlookApp);
            }
        }

        /// <summary>
        /// Brings the Outlook window to the foreground.
        /// </summary>
        private static void BringOutlookToForeground()
        {
            try
            {
                var processes = Process.GetProcessesByName("OUTLOOK");
                if (processes.Length > 0)
                {
                    IntPtr handle = processes[0].MainWindowHandle;
                    if (handle != IntPtr.Zero)
                    {
                        SetForegroundWindow(handle);
                        // Also restore if minimized
                        ShowWindow(handle, SW_RESTORE);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[ToastActionHandler] Error bringing Outlook to foreground: " + ex.Message);
            }
        }

        #region Native Methods

        private const int SW_RESTORE = 9;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        #endregion
    }
}
