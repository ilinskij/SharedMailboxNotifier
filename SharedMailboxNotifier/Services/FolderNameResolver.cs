using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;

namespace SharedMailboxNotifier.Services
{
    /// <summary>
    /// Resolves the Inbox folder across different Outlook language installations.
    /// </summary>
    public static class FolderNameResolver
    {
        private static readonly HashSet<string> KnownInboxNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Inbox",
            "Входящие",
            "Вхідні",
            "Posteingang",
            "Boîte de réception",
            "Bandeja de entrada",
            "Posta in arrivo",
            "Caixa de Entrada",
            "Skrzynka odbiorcza",
            "Postvak IN",
            "Inkorgen",
            "Innboks",
            "Indbakke",
            "Saapuneet",
            "Doručená pošta",
            "Beérkezett üzenetek",
            "Gelen Kutusu",
            "受信トレイ",
            "收件箱",
            "收件匣",
            "받은 편지함"
        };

        /// <summary>
        /// Attempts to find the Inbox folder within a mailbox root folder.
        /// Tries API method first (reliable), then falls back to known folder names.
        /// </summary>
        public static Folder FindInboxFolder(Folder rootFolder)
        {
            if (rootFolder == null)
                return null;

            Store store = null;
            try
            {
                // Method 1 (preferred): Try via Store.GetDefaultFolder — language-independent
                store = rootFolder.Store;
                if (store != null)
                {
                    try
                    {
                        var inbox = store.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;
                        if (inbox != null)
                        {
                            return inbox;
                        }
                    }
                    catch
                    {
                        // GetDefaultFolder may fail for some shared mailboxes — fall through
                    }
                }

                // Method 2 (fallback): Try to find by known localized names
                Folders folders = null;
                try
                {
                    folders = rootFolder.Folders;
                    for (int i = 1; i <= folders.Count; i++)
                    {
                        Folder folder = null;
                        bool matched = false;
                        try
                        {
                            folder = (Folder)folders[i];
                            matched = KnownInboxNames.Contains(folder.Name);
                            if (matched)
                            {
                                return folder;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                        finally
                        {
                            if (folder != null && !matched)
                            {
                                ComHelper.SafeComRelease(folder);
                            }
                        }
                    }
                }
                finally
                {
                    ComHelper.SafeComRelease(folders);
                }
            }
            catch
            {
                // Root folder may be inaccessible
            }
            finally
            {
                ComHelper.SafeComRelease(store);
            }

            return null;
        }

        public static bool IsInboxFolder(string folderName)
        {
            return !string.IsNullOrEmpty(folderName) && KnownInboxNames.Contains(folderName);
        }

        public static void AddInboxName(string name)
        {
            if (!string.IsNullOrWhiteSpace(name))
            {
                KnownInboxNames.Add(name);
            }
        }
    }
}
