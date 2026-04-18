using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SharedMailboxNotifier.Resources;
using SharedMailboxNotifier.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SharedMailboxNotifier.UI
{
    /// <summary>
    /// Settings page for the add-in, displayed in Outlook's Add-in Options dialog.
    /// Implements Outlook.PropertyPage interface for integration with Outlook.
    /// Layout: three GroupBoxes (General, Appearance, About) without TabControl,
    /// since Outlook already wraps the page in its own tab.
    /// </summary>
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567891")]
    public partial class SettingsPage : UserControl, Outlook.PropertyPage
    {
        // DispId for Caption property - required by Outlook to get page title
        private const int CaptionDispId = -518;

        private bool _isDirty;
        private bool _isLoading;
        private Outlook.PropertyPageSite _propertyPageSite;

        // Track original values to detect changes
        private SettingsService.MonitorModeEnum _originalMode;
        private bool _originalSound;
        private bool _originalRoundAppLogo;
        private bool _originalSearchContactPhotos;

        // Controls — General group
        private GroupBox _groupGeneral;
        private GroupBox _groupMonitorMode;
        private RadioButton _radioSharedOnly;
        private RadioButton _radioAllMailboxes;
        private Label _labelWarning;

        // Controls — Appearance group
        private GroupBox _groupAppearance;
        private CheckBox _checkEnableSound;
        private CheckBox _checkRoundAppLogo;
        private CheckBox _checkSearchContactPhotos;
        private Button _btnConfigureCategories;

        // Outlook Application reference — passed from ThisAddIn
        private readonly Outlook.Application _outlookApp;

        // Controls — About group
        private GroupBox _groupAbout;
        private Label _labelDescription;
        private Label _labelPublisher;
        private Label _labelAuthor;
        private Label _labelVersion;

        public SettingsPage(Outlook.Application outlookApp)
        {
            if (outlookApp == null)
                throw new ArgumentNullException("outlookApp");

            _outlookApp = outlookApp;

            InitializeControls();
            LoadSettings();
        }

        #region UI Construction

        private void InitializeControls()
        {
            this.SuspendLayout();

            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.BackColor = SystemColors.Window;
            this.Size = new Size(449, 410);

            const int groupLeft = 8;
            const int groupWidth = 398;
            const int innerLeft = 6;
            const int innerWidth = 386;

            // ═══════════════════════════════════════════════
            // Group: General
            // ═══════════════════════════════════════════════
            _groupGeneral = new GroupBox
            {
                Text = Strings.SettingsGroupGeneral,
                Location = new Point(groupLeft, 6),
                Size = new Size(groupWidth, 130),
                TabIndex = 0,
                FlatStyle = FlatStyle.System
            };

            // Nested group: Monitor Mode
            _groupMonitorMode = new GroupBox
        {
                Text = Strings.SettingsMonitorModeGroup,
                Location = new Point(innerLeft, 19),
                Size = new Size(innerWidth, 44),
                TabIndex = 0,
                FlatStyle = FlatStyle.System
            };

            _radioSharedOnly = new RadioButton
            {
                Text = Strings.SettingsSharedOnly,
                AutoSize = true,
                Location = new Point(6, 19),
                TabIndex = 0,
                FlatStyle = FlatStyle.System
            };
            _radioSharedOnly.CheckedChanged += OnSettingChanged;

            _radioAllMailboxes = new RadioButton
            {
                Text = Strings.SettingsAllMailboxes,
                AutoSize = true,
                Location = new Point(193, 19),
                TabIndex = 1,
                FlatStyle = FlatStyle.System
            };
            _radioAllMailboxes.CheckedChanged += OnSettingChanged;

            _groupMonitorMode.Controls.Add(_radioSharedOnly);
            _groupMonitorMode.Controls.Add(_radioAllMailboxes);

            // Warning label — hidden by default, shown when mode changes.
            // Height allows for two lines of localized text with emoji.
            _labelWarning = new Label
            {
                Text = "",
                Location = new Point(innerLeft + 1, 70),
                Size = new Size(innerWidth - 2, 56),
                ForeColor = Color.DarkOrange,
                Visible = false,
                TabIndex = 1
            };

            _groupGeneral.Controls.Add(_groupMonitorMode);
            _groupGeneral.Controls.Add(_labelWarning);

            // ═══════════════════════════════════════════════
            // Group: Appearance
            // ═══════════════════════════════════════════════
            _groupAppearance = new GroupBox
            {
                Text = Strings.SettingsGroupAppearance,
                Location = new Point(groupLeft, 142),
                Size = new Size(groupWidth, 110),
                TabIndex = 1,
                FlatStyle = FlatStyle.System
            };

            _checkEnableSound = new CheckBox
            {
                Text = Strings.SettingsEnableSound,
                AutoSize = true,
                Location = new Point(innerLeft, 19),
                TabIndex = 0,
                FlatStyle = FlatStyle.System
            };
            _checkEnableSound.CheckedChanged += OnSettingChanged;

            _checkRoundAppLogo = new CheckBox
            {
                Text = Strings.SettingsRoundAppLogo,
                AutoSize = true,
                Location = new Point(innerLeft, 42),
                TabIndex = 1,
                FlatStyle = FlatStyle.System
            };
            _checkRoundAppLogo.CheckedChanged += OnSettingChanged;

            _checkSearchContactPhotos = new CheckBox
            {
                Text = Strings.SettingsSearchContactPhotos,
                Location = new Point(innerLeft, 65),
                Size = new Size(innerWidth, 36),
                TabIndex = 2,
                FlatStyle = FlatStyle.System
            };
            _checkSearchContactPhotos.CheckedChanged += OnSettingChanged;

            _btnConfigureCategories = new Button
            {
                Text = Strings.SettingsConfigureCategories,
                Location = new Point(220, 19),
                Size = new Size(160, 28),
                TabIndex = 3,
                FlatStyle = FlatStyle.System
            };
            _btnConfigureCategories.Click += OnConfigureCategoriesClick;

            _groupAppearance.Controls.Add(_checkEnableSound);
            _groupAppearance.Controls.Add(_checkRoundAppLogo);
            _groupAppearance.Controls.Add(_checkSearchContactPhotos);
            _groupAppearance.Controls.Add(_btnConfigureCategories);

            // ═══════════════════════════════════════════════
            // Group: About
            // ═══════════════════════════════════════════════
            _groupAbout = new GroupBox
            {
                Text = Strings.SettingsGroupAbout,
                Location = new Point(groupLeft, 258),
                Size = new Size(groupWidth, 94),
                TabIndex = 2,
                FlatStyle = FlatStyle.System
            };

            _labelDescription = new Label
            {
                Text = Strings.SettingsAboutDescription,
                AutoSize = true,
                Location = new Point(innerLeft, 16),
                ForeColor = SystemColors.ControlText,
                TabIndex = 0
            };

            _labelPublisher = new Label
            {
                Text = Strings.SettingsAboutPublisher,
                AutoSize = true,
                Location = new Point(innerLeft, 34),
                ForeColor = SystemColors.GrayText,
                TabIndex = 1
            };

            _labelAuthor = new Label
            {
                Text = Strings.SettingsAboutAuthor,
                AutoSize = true,
                Location = new Point(innerLeft, 52),
                ForeColor = SystemColors.GrayText,
                TabIndex = 2
            };

            _labelVersion = new Label
            {
                Text = GetVersionString(),
                AutoSize = true,
                Location = new Point(innerLeft, 70),
                ForeColor = SystemColors.GrayText,
                TabIndex = 3
            };

            _groupAbout.Controls.Add(_labelDescription);
            _groupAbout.Controls.Add(_labelPublisher);
            _groupAbout.Controls.Add(_labelAuthor);
            _groupAbout.Controls.Add(_labelVersion);

            // ═══════════════════════════════════════════════
            // Add all groups to form
            // ═══════════════════════════════════════════════
            this.Controls.Add(_groupGeneral);
            this.Controls.Add(_groupAppearance);
            this.Controls.Add(_groupAbout);

            this.ResumeLayout(false);
        }

        #endregion

        #region Settings Load / Save

        private void LoadSettings()
        {
            _isLoading = true;

            try
            {
                // Monitor mode
                _originalMode = SettingsService.MonitorMode;
                _radioAllMailboxes.Checked = (_originalMode == SettingsService.MonitorModeEnum.AllMailboxes);
                _radioSharedOnly.Checked = (_originalMode == SettingsService.MonitorModeEnum.SharedMailboxesOnly);

                // Sound
                _originalSound = SettingsService.EnableSound;
                _checkEnableSound.Checked = _originalSound;

                // Round App Logo
                _originalRoundAppLogo = SettingsService.RoundAppLogo;
                _checkRoundAppLogo.Checked = _originalRoundAppLogo;

                // Search contact photos
                _originalSearchContactPhotos = SettingsService.SearchContactPhotos;
                _checkSearchContactPhotos.Checked = _originalSearchContactPhotos;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsPage] Error loading settings: " + ex.Message);
            }

            _isLoading = false;
            _isDirty = false;
            UpdateWarningLabel();
        }

        private bool SaveSettings()
        {
            try
            {
                // Monitor mode
                SettingsService.MonitorMode = _radioAllMailboxes.Checked
                    ? SettingsService.MonitorModeEnum.AllMailboxes
                    : SettingsService.MonitorModeEnum.SharedMailboxesOnly;

                // Sound
                SettingsService.EnableSound = _checkEnableSound.Checked;

                // Round App Logo
                SettingsService.RoundAppLogo = _checkRoundAppLogo.Checked;

                // Search contact photos
                SettingsService.SearchContactPhotos = _checkSearchContactPhotos.Checked;

                // Apply Outlook notification settings
                bool requiresRestart = SettingsService.ApplyOutlookNotificationSettings();

                Debug.WriteLine("[SettingsPage] Settings saved. Restart required: " + requiresRestart);

                // Show message if restart required
                if (requiresRestart)
                {
                    MessageBox.Show(
                        Strings.SettingsRestartMessage,
                        "Shared Mailbox Notifier",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                // Update original values
                _originalMode = SettingsService.MonitorMode;
                _originalSound = SettingsService.EnableSound;
                _originalRoundAppLogo = SettingsService.RoundAppLogo;
                _originalSearchContactPhotos = SettingsService.SearchContactPhotos;

                return true;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsPage] Error saving settings: " + ex.Message);
                MessageBox.Show(
                    Strings.ErrorSavingSettings + ex.Message,
                    "Shared Mailbox Notifier",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            return false;
        }

        #endregion

        #region Event Handlers

        private void OnSettingChanged(object sender, EventArgs e)
        {
            if (_isLoading)
                return;

            _isDirty = true;

            UpdateWarningLabel();

            // Notify Outlook that settings have changed (enables Apply button)
            NotifyPropertyPageSite();
        }

        private void OnConfigureCategoriesClick(object sender, EventArgs e)
        {
            try
            {
                using (var form = new CategorySelectorForm(_outlookApp))
                {
                    form.ShowDialog(this);
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsPage] Error opening category selector: " + ex.Message);
            }
        }

        private void UpdateWarningLabel()
        {
            var currentMode = _radioAllMailboxes.Checked
                ? SettingsService.MonitorModeEnum.AllMailboxes
                : SettingsService.MonitorModeEnum.SharedMailboxesOnly;

            bool modeChanged = (currentMode != _originalMode);

            if (modeChanged)
            {
                if (currentMode == SettingsService.MonitorModeEnum.AllMailboxes)
                {
                    // Switching to "All mailboxes" - will disable Outlook alerts
                    _labelWarning.Text = Strings.SettingsWarningAllMailboxes;
                    _labelWarning.ForeColor = Color.DarkOrange;
                }
                else
                {
                    // Switching to "Shared only" - will enable Outlook alerts
                    _labelWarning.Text = Strings.SettingsWarningSharedOnly;
                    _labelWarning.ForeColor = Color.Green;
                }
                _labelWarning.Visible = true;
            }
            else
            {
                _labelWarning.Visible = false;
            }
        }

        #endregion

        #region PropertyPageSite

        private void NotifyPropertyPageSite()
        {
            if (_propertyPageSite == null)
            {
                _propertyPageSite = GetPropertyPageSite();
            }

            if (_propertyPageSite != null)
            {
                try
                {
                    _propertyPageSite.OnStatusChange();
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("[SettingsPage] Error notifying PropertyPageSite: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Gets the PropertyPageSite using reflection (required hack for .NET).
        /// </summary>
        private Outlook.PropertyPageSite GetPropertyPageSite()
        {
            try
            {
                // Get System.Windows.Forms assembly
                var codeBaseUri = new Uri(typeof(object).Assembly.CodeBase);
                string formsAssemblyPath = codeBaseUri.LocalPath
                    .Replace("mscorlib.dll", "System.Windows.Forms.dll");

                string assemblyName = AssemblyName.GetAssemblyName(formsAssemblyPath).FullName;

                // Get UnsafeNativeMethods type
                Type unsafeMethodsType = Type.GetType(
                    Assembly.CreateQualifiedName(assemblyName, "System.Windows.Forms.UnsafeNativeMethods"));

                if (unsafeMethodsType == null)
                    return null;

                // Get IOleObject nested type
                Type oleObjectType = unsafeMethodsType.GetNestedType("IOleObject");
                if (oleObjectType == null)
                    return null;

                // Get GetClientSite method
                MethodInfo getClientSiteMethod = oleObjectType.GetMethod("GetClientSite");

                if (getClientSiteMethod == null)
                    return null;

                // Invoke GetClientSite
                object site = getClientSiteMethod.Invoke(this, null);

                return site as Outlook.PropertyPageSite;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[SettingsPage] Error getting PropertyPageSite: " + ex.Message);
                return null;
            }
        }

        #endregion

        #region Outlook.PropertyPage Implementation

        /// <summary>
        /// Caption for the property page tab.
        /// </summary>
        [DispId(CaptionDispId)]
        public string Caption
        {
            get { return "Shared Mailbox Notifier"; }
        }

        /// <summary>
        /// Returns true if settings have been modified.
        /// </summary>
        bool Outlook.PropertyPage.Dirty
        {
            get { return _isDirty; }
        }

        /// <summary>
        /// Called when user clicks OK or Apply.
        /// </summary>
        void Outlook.PropertyPage.Apply()
        {
            if (SaveSettings())
            {
                _isDirty = false;
                UpdateWarningLabel();
            }
        }

        /// <summary>
        /// Gets help file info (not used).
        /// </summary>
        void Outlook.PropertyPage.GetPageInfo(ref string HelpFile, ref int HelpContext)
        {
            HelpFile = string.Empty;
            HelpContext = 0;
        }

        #endregion

        #region Helpers

        private static string GetVersionString()
        {
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                return string.Format("v{0}.{1}.{2}", version.Major, version.Minor, version.Build);
            }
            catch
            {
                return "";
            }
        }

        #endregion
    }
}
