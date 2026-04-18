using System;
using System.Text;
using System.Text.RegularExpressions;

namespace SharedMailboxNotifier.Services
{
    /// <summary>
    /// Pure string utility methods used across the add-in.
    /// Extracted for testability — no COM or Outlook dependencies.
    /// </summary>
    public static class TextUtils
    {
        private static readonly Regex MultiSpaceRegex =
            new Regex(@" {2,}", RegexOptions.Compiled);
        private static readonly Regex TripleNewlineRegex =
            new Regex(@"\n{3,}", RegexOptions.Compiled);
        private static readonly Regex WhitespaceRegex =
            new Regex(@"\s+", RegexOptions.Compiled);

        /// <summary>
        /// Removes control characters, collapses multiple spaces, and truncates to maxLength.
        /// Used to sanitize user-visible text in notifications.
        /// </summary>
        public static string Sanitize(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            var sb = new StringBuilder(text.Length);
            foreach (char c in text)
            {
                if (!char.IsControl(c))
                    sb.Append(c);
            }

            text = MultiSpaceRegex.Replace(sb.ToString(), " ").Trim();

            if (text.Length > maxLength)
            {
                if (maxLength <= 3)
                    return text.Substring(0, maxLength);

                text = text.Substring(0, maxLength - 3) + "...";
            }

            return text;
        }

        /// <summary>
        /// Collapses a string to a single line, removes extra whitespace,
        /// and truncates with "…" if longer than maxLength.
        /// </summary>
        public static string TrimSingleLine(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text)) { return string.Empty; }

            string result = text.Replace("\r", " ").Replace("\n", " ").Trim();

            result = MultiSpaceRegex.Replace(result, " ");

            if (result.Length <= maxLength) { return result; }
            if (maxLength <= 1) { return "…"; }

            return result.Substring(0, maxLength - 1).TrimEnd() + "…";
        }

        /// <summary>
        /// Truncates multiline text to maxLength, trying to break at word boundaries.
        /// Falls back to hard cut if no good break point is found within first 10 chars.
        /// </summary>
        public static string TrimMultilinePreservingWords(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || maxLength <= 0) { return string.Empty; }
            if (text.Length <= maxLength) { return text; }
            if (maxLength == 1) { return "…"; }

            int cut = maxLength - 1;

            while (cut > 0 && !char.IsWhiteSpace(text[cut - 1]) && text[cut - 1] != '-' && text[cut - 1] != ',' && text[cut - 1] != '.')
            {
                cut--;
            }

            if (cut < 10)
            {
                cut = maxLength - 1;
            }

            return text.Substring(0, cut).TrimEnd() + "…";
        }

        /// <summary>
        /// Normalizes multiline text: unifies line endings, collapses whitespace,
        /// limits consecutive blank lines to one.
        /// </summary>
        public static string NormalizeMultilineText(string text)
        {
            if (string.IsNullOrEmpty(text)) { return string.Empty; }

            string result = text.Replace("\r\n", "\n").Replace('\r', '\n');
            result = result.Replace('\t', ' ');

            result = MultiSpaceRegex.Replace(result, " ");
            result = TripleNewlineRegex.Replace(result, "\n\n");

            return result.Replace("\n", Environment.NewLine).Trim();
        }

        /// <summary>
        /// Builds the text body for a BalloonTip notification.
        /// </summary>
        public static string BuildMailBalloonText(string mailboxName, string senderName, string bodyPreview, DateTime receivedTime, int maxLength)
        {
            const string Divider = "•—•—•—•";

            string fromToLine = TrimSingleLine(senderName, 80) + " •—> " + TrimSingleLine(mailboxName, 80);
            string receivedDateRepresented = receivedTime.ToString("dddd, d MMMM yyyy, HH:mm");
            string dateLine = char.ToUpper(receivedDateRepresented[0]) + receivedDateRepresented.Substring(1);

            string prefix =
                fromToLine + Environment.NewLine +
                Divider + Environment.NewLine;

            string suffix =
                Environment.NewLine + Divider + Environment.NewLine +
                dateLine;

            int availableForBodyPreview = maxLength - prefix.Length - suffix.Length;

            if (availableForBodyPreview < 0) { availableForBodyPreview = 0; }

            string finalBodyPreview = TrimMultilinePreservingWords(bodyPreview, availableForBodyPreview);

            return prefix + finalBodyPreview + suffix;
        }

        /// <summary>
        /// Collapses all whitespace (newlines, tabs, multiple spaces) into single spaces.
        /// Used for email body preview generation.
        /// </summary>
        public static string CollapseWhitespace(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            return WhitespaceRegex.Replace(text, " ").Trim();
        }

        /// <summary>
        /// Converts an email address or name into a safe file name component.
        /// Replaces invalid chars with underscores, truncates to 80 chars.
        /// </summary>
        public static string SanitizeForFileName(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "unknown";

            var chars = input.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (!char.IsLetterOrDigit(chars[i]) && chars[i] != '.' && chars[i] != '-' && chars[i] != '_')
                {
                    chars[i] = '_';
                }
            }

            var result = new string(chars);
            if (result.Length > 80)
                result = result.Substring(0, 80);

            return result;
        }
    }
}
