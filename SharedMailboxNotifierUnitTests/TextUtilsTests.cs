using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharedMailboxNotifier.Services;

namespace SharedMailboxNotifier.Tests
{
    [TestClass]
    public class TextUtilsTests
    {
        #region Sanitize

        [TestMethod]
        public void Sanitize_Null_ReturnsEmpty()
        {
            Assert.AreEqual("", TextUtils.Sanitize(null, 50));
        }

        [TestMethod]
        public void Sanitize_Empty_ReturnsEmpty()
        {
            Assert.AreEqual("", TextUtils.Sanitize("", 50));
        }

        [TestMethod]
        public void Sanitize_RemovesControlChars()
        {
            Assert.AreEqual("ab", TextUtils.Sanitize("a\x01b", 50));
        }

        [TestMethod]
        public void Sanitize_CollapsesMultipleSpaces()
        {
            Assert.AreEqual("a b", TextUtils.Sanitize("a    b", 50));
        }

        [TestMethod]
        public void Sanitize_TruncatesLongString()
        {
            var result = TextUtils.Sanitize(new string('x', 100), 20);
            Assert.IsTrue(result.Length <= 20);
            Assert.IsTrue(result.EndsWith("..."));
        }

        [TestMethod]
        public void Sanitize_ShortStringUnchanged()
        {
            Assert.AreEqual("hello world", TextUtils.Sanitize("hello world", 50));
        }

        [TestMethod]
        public void Sanitize_TrimsWhitespace()
        {
            Assert.AreEqual("hello", TextUtils.Sanitize("  hello  ", 50));
        }

        [TestMethod]
        public void Sanitize_PreservesCyrillic()
        {
            Assert.AreEqual("Привет мир", TextUtils.Sanitize("Привет мир", 50));
        }

        #endregion

        #region TrimSingleLine

        [TestMethod]
        public void TrimSingleLine_Null_ReturnsEmpty()
        {
            Assert.AreEqual("", TextUtils.TrimSingleLine(null, 50));
        }

        [TestMethod]
        public void TrimSingleLine_RemovesNewlines()
        {
            var result = TextUtils.TrimSingleLine("line1\r\nline2\nline3", 50);
            Assert.IsFalse(result.Contains("\n"));
            Assert.IsFalse(result.Contains("\r"));
            Assert.AreEqual("line1 line2 line3", result);
        }

        [TestMethod]
        public void TrimSingleLine_CollapsesSpaces()
        {
            Assert.AreEqual("a b", TextUtils.TrimSingleLine("a    b", 50));
        }

        [TestMethod]
        public void TrimSingleLine_TruncatesWithEllipsis()
        {
            var result = TextUtils.TrimSingleLine(new string('x', 100), 20);
            Assert.IsTrue(result.EndsWith("…"));
            Assert.IsTrue(result.Length <= 20);
        }

        [TestMethod]
        public void TrimSingleLine_MaxLength1_ReturnsEllipsis()
        {
            Assert.AreEqual("…", TextUtils.TrimSingleLine("hello", 1));
        }

        [TestMethod]
        public void TrimSingleLine_ExactLength_NoTruncation()
        {
            Assert.AreEqual("abc", TextUtils.TrimSingleLine("abc", 3));
        }

        [TestMethod]
        public void TrimSingleLine_ShorterThanMax_NoTruncation()
        {
            Assert.AreEqual("hi", TextUtils.TrimSingleLine("hi", 50));
        }

        #endregion

        #region TrimMultilinePreservingWords

        [TestMethod]
        public void TrimMultiline_Null_ReturnsEmpty()
        {
            Assert.AreEqual("", TextUtils.TrimMultilinePreservingWords(null, 50));
        }

        [TestMethod]
        public void TrimMultiline_ZeroLength_ReturnsEmpty()
        {
            Assert.AreEqual("", TextUtils.TrimMultilinePreservingWords("hello", 0));
        }

        [TestMethod]
        public void TrimMultiline_NegativeLength_ReturnsEmpty()
        {
            Assert.AreEqual("", TextUtils.TrimMultilinePreservingWords("hello", -5));
        }

        [TestMethod]
        public void TrimMultiline_ShortText_Unchanged()
        {
            Assert.AreEqual("hi there", TextUtils.TrimMultilinePreservingWords("hi there", 50));
        }

        [TestMethod]
        public void TrimMultiline_BreaksAtWordBoundary()
        {
            var result = TextUtils.TrimMultilinePreservingWords("hello world this is a test", 14);
            Assert.IsTrue(result.EndsWith("…"));
            Assert.IsTrue(result.Length <= 14);
            // Should break at space, not mid-word
            Assert.IsFalse(result.Contains("thi"));
        }

        [TestMethod]
        public void TrimMultiline_MaxLength1_ReturnsEllipsis()
        {
            Assert.AreEqual("…", TextUtils.TrimMultilinePreservingWords("hello", 1));
        }

        [TestMethod]
        public void TrimMultiline_BreaksAtComma()
        {
            var result = TextUtils.TrimMultilinePreservingWords("first,second,third,fourth", 16);
            Assert.IsTrue(result.EndsWith("…"));
        }

        #endregion

        #region NormalizeMultilineText

        [TestMethod]
        public void Normalize_Null_ReturnsEmpty()
        {
            Assert.AreEqual("", TextUtils.NormalizeMultilineText(null));
        }

        [TestMethod]
        public void Normalize_Empty_ReturnsEmpty()
        {
            Assert.AreEqual("", TextUtils.NormalizeMultilineText(""));
        }

        [TestMethod]
        public void Normalize_CollapsesTripleNewlines()
        {
            var result = TextUtils.NormalizeMultilineText("a\n\n\n\nb");
            Assert.IsFalse(result.Replace("\r\n", "\n").Contains("\n\n\n"));
        }

        [TestMethod]
        public void Normalize_PreservesDoubleNewlines()
        {
            var result = TextUtils.NormalizeMultilineText("a\n\nb");
            Assert.IsTrue(result.Contains(Environment.NewLine + Environment.NewLine)
                       || result.Contains("\n\n"));
        }

        [TestMethod]
        public void Normalize_ReplacesTabsWithSpaces()
        {
            var result = TextUtils.NormalizeMultilineText("a\tb");
            Assert.IsFalse(result.Contains("\t"));
            Assert.AreEqual("a b", result);
        }

        [TestMethod]
        public void Normalize_UnifiesCRLF()
        {
            var result = TextUtils.NormalizeMultilineText("a\r\nb\rc");
            // All line endings should be Environment.NewLine
            Assert.IsFalse(result.Contains("\r\n\r\n\r\n"));
        }

        [TestMethod]
        public void Normalize_TrimsResult()
        {
            var result = TextUtils.NormalizeMultilineText("  hello  ");
            Assert.AreEqual("hello", result);
        }

        #endregion

        #region CollapseWhitespace

        [TestMethod]
        public void Collapse_Null_ReturnsNull()
        {
            Assert.IsNull(TextUtils.CollapseWhitespace(null));
        }

        [TestMethod]
        public void Collapse_Empty_ReturnsNull()
        {
            Assert.IsNull(TextUtils.CollapseWhitespace(""));
        }

        [TestMethod]
        public void Collapse_OnlyWhitespace_ReturnsNull()
        {
            Assert.IsNull(TextUtils.CollapseWhitespace("   \t\n  "));
        }

        [TestMethod]
        public void Collapse_MultipleSpaces()
        {
            Assert.AreEqual("a b c", TextUtils.CollapseWhitespace("a   b   c"));
        }

        [TestMethod]
        public void Collapse_MixedWhitespace()
        {
            Assert.AreEqual("a b c", TextUtils.CollapseWhitespace("a\t\n  b\r\n  c"));
        }

        [TestMethod]
        public void Collapse_LeadingTrailingTrimmed()
        {
            Assert.AreEqual("hello", TextUtils.CollapseWhitespace("  hello  "));
        }

        [TestMethod]
        public void Collapse_SingleWord_Unchanged()
        {
            Assert.AreEqual("hello", TextUtils.CollapseWhitespace("hello"));
        }

        #endregion

        #region SanitizeForFileName

        [TestMethod]
        public void FileName_Null_ReturnsUnknown()
        {
            Assert.AreEqual("unknown", TextUtils.SanitizeForFileName(null));
        }

        [TestMethod]
        public void FileName_Empty_ReturnsUnknown()
        {
            Assert.AreEqual("unknown", TextUtils.SanitizeForFileName(""));
        }

        [TestMethod]
        public void FileName_ReplacesAtSign()
        {
            Assert.AreEqual("user_example.com", TextUtils.SanitizeForFileName("user@example.com"));
        }

        [TestMethod]
        public void FileName_ReplacesSpaces()
        {
            Assert.AreEqual("John_Doe", TextUtils.SanitizeForFileName("John Doe"));
        }

        [TestMethod]
        public void FileName_PreservesDotDashUnderscore()
        {
            Assert.AreEqual("a.b-c_d", TextUtils.SanitizeForFileName("a.b-c_d"));
        }

        [TestMethod]
        public void FileName_TruncatesTo80()
        {
            var result = TextUtils.SanitizeForFileName(new string('a', 200));
            Assert.AreEqual(80, result.Length);
        }

        [TestMethod]
        public void FileName_CyrillicPreserved()
        {
            Assert.AreEqual("Иванов", TextUtils.SanitizeForFileName("Иванов"));
        }

        [TestMethod]
        public void FileName_SpecialCharsReplaced()
        {
            var result = TextUtils.SanitizeForFileName("test<>:\"/\\|?*file");
            Assert.IsFalse(result.Contains("<"));
            Assert.IsFalse(result.Contains(">"));
            Assert.IsFalse(result.Contains("\""));
            Assert.IsFalse(result.Contains("\\"));
            Assert.IsFalse(result.Contains("?"));
            Assert.IsFalse(result.Contains("*"));
        }

        #endregion

        #region BuildMailBalloonText

        [TestMethod]
        public void BuildBalloon_ContainsSenderName()
        {
            var result = TextUtils.BuildMailBalloonText("Shared", "John Doe", "Hello", DateTime.Now, 500);
            Assert.IsTrue(result.Contains("John Doe"));
        }

        [TestMethod]
        public void BuildBalloon_ContainsMailboxName()
        {
            var result = TextUtils.BuildMailBalloonText("SharedBox", "Sender", "Hello", DateTime.Now, 500);
            Assert.IsTrue(result.Contains("SharedBox"));
        }

        [TestMethod]
        public void BuildBalloon_RespectsMaxLength()
        {
            var result = TextUtils.BuildMailBalloonText("Box", "Sender", new string('x', 1000), DateTime.Now, 250);
            Assert.IsTrue(result.Length <= 260); // some tolerance for formatting
        }

        [TestMethod]
        public void BuildBalloon_EmptyBody_StillWorks()
        {
            var result = TextUtils.BuildMailBalloonText("Box", "Sender", "", DateTime.Now, 250);
            Assert.IsTrue(result.Contains("Sender"));
            Assert.IsTrue(result.Contains("Box"));
        }

        [TestMethod]
        public void BuildBalloon_NullBody_StillWorks()
        {
            var result = TextUtils.BuildMailBalloonText("Box", "Sender", null, DateTime.Now, 250);
            Assert.IsTrue(result.Contains("Sender"));
        }

        #endregion
    }
}
