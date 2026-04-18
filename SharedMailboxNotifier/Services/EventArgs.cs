using System;

namespace SharedMailboxNotifier.Services
{
    public class MonitorStatusEventArgs : EventArgs
    {
        public bool IsRunning { get; set; }
        public int SharedMailboxCount { get; set; }
        public string Message { get; set; }
    }

    public class NewMailEventArgs : EventArgs
    {
        public string MailboxName { get; set; }
        public string SenderName { get; set; }
        public string Subject { get; set; }
        public string BodyPreview { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string EntryId { get; set; }
        public string StoreId { get; set; }
    }
}
