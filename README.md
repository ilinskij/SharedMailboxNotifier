# Shared Mailbox Notifier

**English** · [Русский](README.ru.md)

An Outlook VSTO add-in that shows Windows Toast notifications for new emails arriving in shared mailboxes.

## The Problem

Outlook does not show desktop notifications for shared (additional) mailboxes — only for the user's own account. If you monitor several shared mailboxes (e.g. `info@`, `support@`, `hr@`), you have to keep checking them manually.

## What This Add-in Does

When a new unread email arrives in any monitored shared mailbox's Inbox, the add-in shows a native Windows Toast notification with:

- **Subject, sender name, and body preview**
- **Sender photo** (from Active Directory or local Contacts)
- **Action buttons**: Reply, Mark as Read, Assign Category
- **Mailbox name and timestamp** in attribution

Notifications are automatically removed when the email is read or deleted in Outlook.

## Features

- **Two monitor modes**: shared mailboxes only, or all mailboxes (with automatic suppression of Outlook's built-in alerts to avoid duplicates)
- **Toast notifications** on Windows 10/11 with interactive buttons
- **Balloon fallback** on Windows 7/8
- **Category assignment** directly from the notification (with a configurable selection of up to 5 categories)
- **Sender photos** from Exchange/AD (via MAPI properties) and optionally from local Contacts
- **Dynamic store tracking**: automatically picks up mailboxes added or removed at runtime
- **Localization**: English, German, Russian, Ukrainian
- **Settings page** integrated into Outlook's Options dialog
- **Event Log** integration for diagnostics

## Requirements

- Windows 7 or later (Windows 10 1607+ for Toast notifications)
- Microsoft Outlook 2010, 2013, 2016, 2019, or Microsoft 365 (desktop)
- .NET Framework 4.8
- Visual Studio 2022 with the **Office/SharePoint development** workload (for building from source)

## Building

1. Open `SharedMailboxNotifier.csproj` in Visual Studio 2022.
2. Ensure the **Office/SharePoint development** workload is installed.
3. Build the solution. NuGet packages will be restored automatically.
4. Press F5 to debug — Outlook will launch with the add-in loaded.

## Installation

For end-user deployment, publish via ClickOnce from Visual Studio or create an MSI installer. The add-in requires the Visual Studio 2010 Tools for Office Runtime (VSTO Runtime), which is included with Office 2010+.

## Configuration

After installation, open Outlook → **File → Options → Add-ins tab → Shared Mailbox Notifier**.

| Setting | Description |
|---|---|
| **Monitor mode** | *Shared only* — monitors additional mailboxes, Outlook handles personal notifications. *All mailboxes* — monitors everything, Outlook's built-in alerts are disabled. |
| **Sound notification** | Enable or disable the notification sound. |
| **Round icon** | Crop the sender photo / app logo to a circle in notifications. |
| **Search contact photos** | Look up sender photos in local Contacts (may slow down notifications for large address books). |
| **Configure categories** | Select up to 5 Outlook categories to show in the notification's category dropdown. |

## Project Structure

```
SharedMailboxNotifier/
├── ThisAddIn.cs                    # VSTO entry point
├── Services/
│   ├── SharedMailboxMonitor.cs     # Store discovery and watcher lifecycle
│   ├── MailboxWatcher.cs           # Per-mailbox Inbox monitoring (ItemAdd/Change/Remove)
│   ├── NotificationService.cs      # Toast and Balloon notification backends
│   ├── ToastActionHandler.cs       # Reply, Mark Read, Assign Category from notifications
│   ├── CategoryService.cs          # Outlook category reading and selection
│   ├── SettingsService.cs          # Registry-based settings + Outlook notification control
│   ├── LogService.cs               # Windows Event Log integration
│   ├── FolderNameResolver.cs       # Localized Inbox folder detection
│   ├── TextUtils.cs                # String sanitization and truncation
│   ├── ComUtils.cs                 # COM object release helper
│   └── EventArgs.cs                # Custom event argument types
├── UI/
│   ├── SettingsPage.cs             # Outlook Options integration (PropertyPage)
│   ├── CategorySelectorForm.cs     # Category picker dialog
│   └── CategorySelectorForm.Designer.cs
├── Resources/
│   ├── Strings.resx                # English (default)
│   ├── Strings.de.resx             # German
│   ├── Strings.ru.resx             # Russian
│   └── Strings.uk.resx             # Ukrainian
└── Images/                         # Notification button icons and app logo
```

## Technical Notes

- New mail detection relies on the `Items.ItemAdd` event on the Inbox folder. The Outlook object model does not provide a dedicated "new mail delivered to this store" event for shared mailboxes, so `ItemAdd` is the only reliable mechanism. A side effect is that any unread item appearing in the Inbox — whether delivered normally, moved from Junk Email, or dragged from another mailbox — will trigger a notification. This is a conscious trade-off, not a bug.
- The add-in runs in-process on Outlook's STA thread. All COM event handlers (`ItemAdd`, `StoreAdd`, `BeforeStoreRemove`) are marshaled to this thread by COM.
- COM object references (`Items`, `Stores`, `Folder`) are explicitly held to prevent garbage collection, which would silently unsubscribe event handlers.
- Toast notification callbacks arrive on an MTA thread. The `ToastActionHandler` obtains a fresh Outlook `Application` reference via `Marshal.GetActiveObject` rather than using the STA-bound reference.
- Settings are stored in `HKCU\Software\SharedMailboxNotifier`. The add-in also manages Outlook's own notification registry keys (`NewmailDesktopAlerts`, `PlaySound`) to prevent duplicate alerts in "All mailboxes" mode.

## License

MIT License — see [LICENSE](LICENSE).

## Credits

© Ilinskij Aleksandr

Developed with [Claude](https://claude.ai) (Anthropic)
