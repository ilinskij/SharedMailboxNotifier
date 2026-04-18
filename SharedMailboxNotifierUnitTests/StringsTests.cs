using System;
using System.Globalization;
using System.Reflection;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharedMailboxNotifier.Resources;

namespace SharedMailboxNotifier.Tests
{
    /// <summary>
    /// Tests that all string resources are present and non-empty for all supported languages,
    /// and that format strings contain expected placeholders.
    /// </summary>
    [TestClass]
    public class StringsTests
    {
        private static readonly string[] SupportedLanguages = { "en", "ru", "uk", "de" };

        private static readonly string[] AllStringProperties =
        {
            "NotificationTitle", "NotificationFrom", "NoSubject", "InboxFolderName",
            "AppAttribution", "CategoryLabel",
            "ButtonAddCategory", "ButtonReply", "ButtonMarkRead",
            "SettingsMonitorModeGroup", "SettingsAllMailboxes", "SettingsSharedOnly",
            "SettingsEnableSound", "SettingsRoundAppLogo", "SettingsSearchContactPhotos",
            "SettingsGroupGeneral", "SettingsGroupAppearance", "SettingsGroupAbout",
            "SettingsWarningAllMailboxes", "SettingsWarningSharedOnly",
            "SettingsRestartMessage", "SettingsAboutDescription",
            "SettingsAboutPublisher", "SettingsAboutAuthor",
            "CategorySelectorTitle", "CategoryAvailable", "CategorySelected",
            "CategoryOk", "CategoryCancel", "CategoryLimitReached", "CategoryCount",
            "ErrorSavingSettings",
            "LogStartupSuccess", "LogMonitoringCount", "LogShutdown",
            "LogNewMail", "LogMailbox", "LogFrom", "LogSubject",
            "LogStoreAdded", "LogStoreRemoved",
            "LogFailedToStart", "LogErrorShutdown",
            "LogErrorToastAction", "LogErrorNewMail", "LogFallbackToBalloon",
            "LogOutlookAlertsSaved", "LogOutlookAlertsDisabled",
            "LogOutlookAlertsRestored", "LogOutlookSoundChanged"
        };

        /// <summary>
        /// Keys that should contain {0} placeholder.
        /// </summary>
        private static readonly string[] FormatStringKeys =
        {
            "NotificationFrom", "LogMonitoringCount",
            "LogMailbox", "LogFrom", "LogSubject",
            "LogStoreAdded", "LogStoreRemoved",
            "CategoryLimitReached", "CategoryCount",
            "LogOutlookAlertsSaved", "LogOutlookAlertsRestored",
            "LogOutlookSoundChanged"
        };

        [TestMethod]
        public void AllStrings_NonEmpty_ForAllLanguages()
        {
            var originalCulture = Thread.CurrentThread.CurrentUICulture;

            try
            {
                foreach (var lang in SupportedLanguages)
                {
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo(lang);

                    foreach (var propName in AllStringProperties)
                    {
                        var prop = typeof(Strings).GetProperty(propName, BindingFlags.Public | BindingFlags.Static);
                        Assert.IsNotNull(prop, string.Format("Property '{0}' not found in Strings class", propName));

                        var value = (string)prop.GetValue(null);
                        Assert.IsFalse(
                            string.IsNullOrWhiteSpace(value),
                            string.Format("String '{0}' is empty for language '{1}'", propName, lang));
                    }
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentUICulture = originalCulture;
            }
        }

        [TestMethod]
        public void FormatStrings_ContainPlaceholder_ForAllLanguages()
        {
            var originalCulture = Thread.CurrentThread.CurrentUICulture;

            try
            {
                foreach (var lang in SupportedLanguages)
                {
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo(lang);

                    foreach (var propName in FormatStringKeys)
                    {
                        var prop = typeof(Strings).GetProperty(propName, BindingFlags.Public | BindingFlags.Static);
                        var value = (string)prop.GetValue(null);

                        Assert.IsTrue(
                            value.Contains("{0}"),
                            string.Format("Format string '{0}' is missing {{0}} placeholder for language '{1}'. Value: '{2}'",
                                propName, lang, value));
                    }
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentUICulture = originalCulture;
            }
        }

        [TestMethod]
        public void UnknownLanguage_FallsBackToEnglish()
        {
            var originalCulture = Thread.CurrentThread.CurrentUICulture;

            try
            {
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("ja"); // Japanese — not supported

                // Should return English fallback, not empty or key name
                var title = Strings.NotificationTitle;
                Assert.AreEqual("New message", title);
            }
            finally
            {
                Thread.CurrentThread.CurrentUICulture = originalCulture;
            }
        }
    }
}
