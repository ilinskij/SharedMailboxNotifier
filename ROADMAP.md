# Roadmap

Current version: **2.0.0.0**

This document outlines known improvements and future plans for the project. Items are grouped by priority. Contributions and suggestions are welcome — feel free to open an Issue to discuss any of these.

## High Priority

### Cache categories and settings on the hot path

Every Toast notification triggers a full cycle: `GetNamespace("MAPI")` → COM iteration over all Outlook categories → registry read (`SelectedCategories`) → string parse → match. Additionally, `SettingsService.RoundAppLogo` and `EnableSound` each open and close a registry key.

**Improvement:** cache the resolved category list and settings in memory, invalidate on save (from `CategorySelectorForm` and `SettingsPage`).

### Replace SHA256 per-notification with a lighter approach

`GenerateNotificationTag` creates and disposes a `SHA256` instance for every notification. For a set of <50 active notifications, this is overkill.

**Options:** reuse a static `SHA256` instance (safe on STA thread), or use a simpler hash.

## Medium Priority

### Event Log source registration with a manifest

Currently, `LogService` probes a hardcoded list of existing Event Log sources (`Microsoft Office 16 Alerts`, `Application`, etc.) and borrows whatever works. This means log entries may display with a "the description for Event ID ... cannot be found" warning, and the source name in Event Viewer doesn't match the add-in.

**Goal:** register a proper Event Log source (`Outlook Shared Mailbox Notifier`) with a message resource DLL / manifest during installation (MSI). This would allow the add-in to write cleanly to the `OAlerts` log or its own log with correctly formatted messages.

### Startup performance: EventLog.SourceExists probe chain

`LogService.Initialize` iterates up to 10 source names, calling `EventLog.SourceExists` for each. Each call scans `HKLM\SYSTEM\CurrentControlSet\Services\EventLog\*`, which can take 20–50 ms. In the worst case (no source found), this adds up to ~500 ms on the startup path — enough for Outlook to flag the add-in as slow.

**Improvement:** try the primary source first and fall back to `Application` immediately if it fails, skipping the intermediate Office sources. Alternatively, cache the result in registry after first successful probe.

### Clean up dead code

Several public methods are declared but never called anywhere in the project:

- `TextUtils.CollapseWhitespace()`
- `LogService.LogNewMail()`
- `FolderNameResolver.IsInboxFolder()` and `AddInboxName()`
- `NotificationService.ShowNewMailNotification(string, string, string)` (3-parameter overload)
- `MailboxWatcher.NewMailReceived` event (raised but never subscribed to)

These are artifacts of earlier development sessions. Removing them reduces surface area and confusion for contributors.

## Low Priority

### Unify code style across files

The codebase was developed across multiple sessions, which left some stylistic inconsistencies:

- `[unknown]` placeholder in some debug message method tags
- Varying debug message formats: `"Error X: "`, `"Failed to X, because "`, `"[MethodName] Suppressed: "`

### Add XML documentation to undocumented classes

`ComHelper`, `EventArgs` classes, and `FolderNameResolver` have no XML documentation. Other files (written in earlier sessions) have thorough documentation. Leveling this up would help contributors.

### Temp photo cleanup on startup

Cached sender photos in `%TEMP%\SharedMailboxNotifier\` are only cleaned up on add-in shutdown. If Outlook crashes or the add-in is disabled, stale files accumulate. Adding cleanup of old files on startup (e.g. older than 7 days) would be more robust.

### Finer notification filter for moved items

The add-in uses `Items.ItemAdd` on the Inbox folder, which fires for any item appearing in the folder — including items moved from Junk Email or dragged from another mailbox. A heuristic filter (e.g. comparing `ReceivedTime` to current time, skipping items older than N minutes) could reduce false notifications, but would need careful tuning to avoid missing legitimate delayed deliveries.

## Ideas / Future

- **Per-mailbox notification settings** (sound, categories, enabled/disabled)
- **Notification grouping** when multiple emails arrive simultaneously
- **"Do Not Disturb" schedule** to suppress notifications during focus time
- **Support for monitoring subfolders** beyond Inbox
