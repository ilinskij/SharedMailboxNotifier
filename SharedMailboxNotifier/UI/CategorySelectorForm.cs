using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SharedMailboxNotifier.Resources;
using SharedMailboxNotifier.Services;

namespace SharedMailboxNotifier.UI
{
    /// <summary>
    /// Dialog for selecting which Outlook categories appear in toast notifications.
    /// Two-list layout: Available (left) ↔ Selected (right), max 5 selected.
    /// ListBoxes use owner-draw to show colored circles next to category names.
    /// </summary>
    public partial class CategorySelectorForm : Form
    {
        private readonly List<CategoryService.CategoryInfo> _allCategories;

        public CategorySelectorForm(Microsoft.Office.Interop.Outlook.Application outlookApp)
        {
            if (outlookApp == null)
                throw new ArgumentNullException("outlookApp");

            InitializeComponent();

            Microsoft.Office.Interop.Outlook.NameSpace ns = null;
            try
            {
                ns = outlookApp.GetNamespace("MAPI");
                _allCategories = CategoryService.GetOutlookCategories(ns);
            }
            finally
            {
                ComHelper.SafeComRelease(ns);
            }

            // Apply localized strings
            this.Text = Strings.CategorySelectorTitle;
            _labelAvailable.Text = Strings.CategoryAvailable;
            _labelSelected.Text = Strings.CategorySelected;
            _btnOk.Text = Strings.CategoryOk;
            _btnCancel.Text = Strings.CategoryCancel;

            LoadCurrentSelection();
            UpdateButtonStates();
        }

        #region Data Loading

        private void LoadCurrentSelection()
        {
            var selectedNames = CategoryService.LoadSelectedNames();

            _listSelected.Items.Clear();
            _listAvailable.Items.Clear();

            // Fill Selected with saved names (in order), only if they still exist
            foreach (var name in selectedNames)
            {
                var cat = _allCategories.Find(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (cat != null)
                {
                    _listSelected.Items.Add(cat);
                }
            }

            // Fill Available with everything NOT in Selected
            foreach (var cat in _allCategories)
            {
                bool isSelected = selectedNames.Any(n => n.Equals(cat.Name, StringComparison.OrdinalIgnoreCase));
                if (!isSelected)
                {
                    _listAvailable.Items.Add(cat);
                }
            }

            UpdateLimitLabel();
        }

        #endregion

        #region Owner Draw

        private void OnDrawItem(object sender, DrawItemEventArgs e)
        {
            var listBox = (ListBox)sender;
            if (e.Index < 0 || e.Index >= listBox.Items.Count)
                return;

            e.DrawBackground();

            var cat = listBox.Items[e.Index] as CategoryService.CategoryInfo;
            if (cat == null)
                return;

            const int circleSize = 12;
            const int circleMarginLeft = 4;
            const int textMarginLeft = 6;

            // Circle position — vertically centered
            int circleY = e.Bounds.Top + (e.Bounds.Height - circleSize) / 2;
            var circleRect = new Rectangle(
                e.Bounds.Left + circleMarginLeft,
                circleY,
                circleSize,
                circleSize);

            // Draw filled circle
            using (var brush = new SolidBrush(cat.Color))
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.FillEllipse(brush, circleRect);
            }

            // Draw text
            var textColor = (e.State & DrawItemState.Selected) != 0
                ? SystemColors.HighlightText
                : SystemColors.WindowText;
            var textX = circleRect.Right + textMarginLeft;
            var textRect = new Rectangle(textX, e.Bounds.Top, e.Bounds.Width - textX, e.Bounds.Height);

            TextRenderer.DrawText(
                e.Graphics,
                cat.Name,
                listBox.Font,
                textRect,
                textColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);

            e.DrawFocusRectangle();
        }

        #endregion

        #region Button State

        private void UpdateButtonStates()
        {
            _btnAdd.Enabled = _listAvailable.SelectedIndex >= 0
                              && _listSelected.Items.Count < CategoryService.MaxCategories;
            _btnRemove.Enabled = _listSelected.SelectedIndex >= 0;
            _btnUp.Enabled = _listSelected.SelectedIndex > 0;
            _btnDown.Enabled = _listSelected.SelectedIndex >= 0
                               && _listSelected.SelectedIndex < _listSelected.Items.Count - 1;
        }

        private void UpdateLimitLabel()
        {
            _labelLimit.Text = string.Format(Strings.CategoryCount,
                _listSelected.Items.Count, CategoryService.MaxCategories);
        }

        #endregion

        #region Event Handlers

        private void OnAddClick(object sender, EventArgs e)
        {
            if (_listAvailable.SelectedIndex < 0)
                return;

            if (_listSelected.Items.Count >= CategoryService.MaxCategories)
                return;

            var item = _listAvailable.SelectedItem;
            _listAvailable.Items.Remove(item);
            _listSelected.Items.Add(item);
            UpdateLimitLabel();
            UpdateButtonStates();
        }

        private void OnRemoveClick(object sender, EventArgs e)
        {
            if (_listSelected.SelectedIndex < 0)
                return;

            var cat = _listSelected.SelectedItem as CategoryService.CategoryInfo;
            _listSelected.Items.Remove(cat);

            // Insert back in original Outlook order
            int insertIndex = 0;
            foreach (var c in _allCategories)
            {
                if (c == cat)
                    break;

                if (_listAvailable.Items.Contains(c))
                    insertIndex++;
            }
            _listAvailable.Items.Insert(Math.Min(insertIndex, _listAvailable.Items.Count), cat);
            UpdateLimitLabel();
            UpdateButtonStates();
        }

        private void OnMoveUpClick(object sender, EventArgs e)
        {
            int index = _listSelected.SelectedIndex;
            if (index <= 0)
                return;

            var item = _listSelected.Items[index];
            _listSelected.Items.RemoveAt(index);
            _listSelected.Items.Insert(index - 1, item);
            _listSelected.SelectedIndex = index - 1;
            UpdateButtonStates();
        }

        private void OnMoveDownClick(object sender, EventArgs e)
        {
            int index = _listSelected.SelectedIndex;
            if (index < 0 || index >= _listSelected.Items.Count - 1)
                return;

            var item = _listSelected.Items[index];
            _listSelected.Items.RemoveAt(index);
            _listSelected.Items.Insert(index + 1, item);
            _listSelected.SelectedIndex = index + 1;
            UpdateButtonStates();
        }

        private void OnOkClick(object sender, EventArgs e)
        {
            var names = new List<string>();
            foreach (var item in _listSelected.Items)
            {
                var cat = item as CategoryService.CategoryInfo;
                if (cat != null)
                    names.Add(cat.Name);
            }
            CategoryService.SaveSelectedNames(names);

            Debug.WriteLine("[CategorySelectorForm] Saved " + names.Count + " categories");
        }

        private void OnSelectionChanged(object sender, EventArgs e)
        {
            UpdateButtonStates();
        }

        #endregion
    }
}
