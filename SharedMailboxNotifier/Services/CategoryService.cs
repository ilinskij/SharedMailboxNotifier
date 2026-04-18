using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;

namespace SharedMailboxNotifier.Services
{
    /// <summary>
    /// Provides access to the user's Outlook categories.
    /// Reads real category names and colors — never creates or modifies categories.
    /// </summary>
    public static class CategoryService
    {
        /// <summary>
        /// Represents a single Outlook category with its display name, color emoji, and color value.
        /// </summary>
        public class CategoryInfo
        {
            public string Name { get; set; }
            public string ColorEmoji { get; set; }
            public System.Drawing.Color Color { get; set; }
        }

        /// <summary>
        /// Maximum number of categories to show in the toast ComboBox.
        /// Toast API limitation: 5 items max in a selection box.
        /// </summary>
        public const int MaxCategories = 5;

        /// <summary>
        /// Reads the user's categories from Outlook.
        /// Returns an empty list if no categories exist or on error.
        /// </summary>
        public static List<CategoryInfo> GetOutlookCategories(NameSpace ns)
        {
            var result = new List<CategoryInfo>();

            if (ns == null)
                return result;

            Categories categories = null;
            try
            {
                categories = ns.Categories;
                if (categories == null || categories.Count == 0)
                    return result;

                for (int i = 1; i <= categories.Count; i++)
                {
                    Category cat = null;
                    try
                    {
                        cat = categories[i];
                        result.Add(new CategoryInfo
                        {
                            Name = cat.Name,
                            ColorEmoji = GetColorEmoji(cat.Color),
                            Color = GetDrawingColor(cat.Color)
                        });
                    }
                    catch (System.Exception ex)
                    {
                        Debug.WriteLine("[CategoryService] Error reading category: " + ex.Message);
                    }
                    finally
                    {
                        ComHelper.SafeComRelease(cat);
                    }
                }

                Debug.WriteLine(string.Format("[CategoryService] Loaded {0} categories", result.Count));
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[CategoryService] Error reading categories: " + ex.Message);
            }
            finally
            {
                ComHelper.SafeComRelease(categories);
            }

            return result;
        }

        /// <summary>
        /// Ensures that a category selection exists in settings.
        /// On first run, auto-selects the first MaxCategories from Outlook.
        /// On subsequent runs, does nothing.
        /// </summary>
        public static void InitializeSelectedCategories(Application outlookApp)
        {
            if (SettingsService.HasSelectedCategories)
                return;

            NameSpace ns = null;
            try
            {
                ns = outlookApp.GetNamespace("MAPI");
                // ResolveSelectedCategories handles first-run auto-population
                GetSelectedCategories(ns);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[CategoryService] Error initializing categories, because " + ex.Message);
            }
            finally
            {
                ComHelper.SafeComRelease(ns);
            }
        }

        /// <summary>
        /// Returns the categories to show in toast notifications.
        /// Uses saved selection from settings, synced against current Outlook categories.
        /// On first run (no saved selection), auto-selects the first MaxCategories.
        /// </summary>
        public static List<CategoryInfo> GetSelectedCategories(NameSpace ns)
        {
            var allCategories = GetOutlookCategories(ns);
            return ResolveSelectedCategories(allCategories);
        }

        /// <summary>
        /// Core logic: resolves saved category names against the current Outlook
        /// category list. On first run, auto-selects the first MaxCategories and
        /// persists the selection. On subsequent runs, drops categories that no
        /// longer exist in Outlook and persists the cleaned-up list.
        /// </summary>
        private static List<CategoryInfo> ResolveSelectedCategories(List<CategoryInfo> allCategories)
        {
            if (allCategories.Count == 0)
                return new List<CategoryInfo>();

            // First run — no selection saved yet
            if (!SettingsService.HasSelectedCategories)
            {
                var initial = allCategories.GetRange(0, Math.Min(allCategories.Count, MaxCategories));
                SaveSelectedNames(initial.ConvertAll(c => c.Name));
                Debug.WriteLine("[CategoryService] First run: saved " + initial.Count + " categories to registry");
                return initial;
            }

            // Saved selection exists — sync with current Outlook categories
            var savedNames = LoadSelectedNames();
            var result = new List<CategoryInfo>();

            foreach (var name in savedNames)
            {
                var match = allCategories.Find(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                {
                    result.Add(match);
                }
                // else: category was deleted from Outlook — silently drop it
            }

            // If some saved categories were deleted, persist the cleaned-up list
            if (result.Count < savedNames.Count)
            {
                SaveSelectedNames(result.ConvertAll(c => c.Name));
                Debug.WriteLine("[CategoryService] Cleaned up orphaned categories. Remaining: " + result.Count);
            }

            return result;
        }

        /// <summary>
        /// Parses the saved category names from settings.
        /// </summary>
        public static List<string> LoadSelectedNames()
        {
            var raw = SettingsService.SelectedCategories;
            if (string.IsNullOrEmpty(raw))
                return new List<string>();

            var names = new List<string>();
            foreach (var part in raw.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = part.Trim();
                if (trimmed.Length > 0)
                    names.Add(trimmed);
            }
            return names;
        }

        /// <summary>
        /// Saves the selected category names to settings.
        /// </summary>
        public static void SaveSelectedNames(List<string> names)
        {
            if (names == null || names.Count == 0)
            {
                SettingsService.SelectedCategories = "";
                return;
            }
            SettingsService.SelectedCategories = string.Join(", ", names);
        }

        /// <summary>
        /// Maps OlCategoryColor enum to a display emoji.
        /// </summary>
        private static string GetColorEmoji(OlCategoryColor color)
        {
            switch (color)
            {
                case OlCategoryColor.olCategoryColorRed:
                case OlCategoryColor.olCategoryColorDarkRed:
                    return "🔴";
                case OlCategoryColor.olCategoryColorOrange:
                case OlCategoryColor.olCategoryColorPeach:
                    return "🟠";
                case OlCategoryColor.olCategoryColorYellow:
                case OlCategoryColor.olCategoryColorDarkYellow:
                    return "🟡";
                case OlCategoryColor.olCategoryColorGreen:
                case OlCategoryColor.olCategoryColorDarkGreen:
                    return "🟢";
                case OlCategoryColor.olCategoryColorBlue:
                case OlCategoryColor.olCategoryColorDarkBlue:
                    return "🔵";
                case OlCategoryColor.olCategoryColorPurple:
                case OlCategoryColor.olCategoryColorDarkPurple:
                case OlCategoryColor.olCategoryColorMaroon:
                    return "🟣";
                case OlCategoryColor.olCategoryColorTeal:
                case OlCategoryColor.olCategoryColorDarkTeal:
                    return "🟢";
                case OlCategoryColor.olCategoryColorOlive:
                case OlCategoryColor.olCategoryColorDarkOlive:
                    return "🟤";
                case OlCategoryColor.olCategoryColorSteel:
                case OlCategoryColor.olCategoryColorDarkSteel:
                case OlCategoryColor.olCategoryColorGray:
                case OlCategoryColor.olCategoryColorDarkGray:
                    return "⚫";
                default:
                    return "⬜";
            }
        }

        /// <summary>
        /// Maps OlCategoryColor enum to a System.Drawing.Color for owner-drawn controls.
        /// </summary>
        private static System.Drawing.Color GetDrawingColor(OlCategoryColor color)
        {
            switch (color)
            {
                case OlCategoryColor.olCategoryColorRed:
                    return System.Drawing.Color.FromArgb(227, 58, 58);
                case OlCategoryColor.olCategoryColorDarkRed:
                    return System.Drawing.Color.FromArgb(178, 34, 34);
                case OlCategoryColor.olCategoryColorOrange:
                    return System.Drawing.Color.FromArgb(243, 154, 29);
                case OlCategoryColor.olCategoryColorPeach:
                    return System.Drawing.Color.FromArgb(255, 181, 108);
                case OlCategoryColor.olCategoryColorYellow:
                    return System.Drawing.Color.FromArgb(253, 220, 55);
                case OlCategoryColor.olCategoryColorDarkYellow:
                    return System.Drawing.Color.FromArgb(194, 176, 54);
                case OlCategoryColor.olCategoryColorGreen:
                    return System.Drawing.Color.FromArgb(78, 166, 70);
                case OlCategoryColor.olCategoryColorDarkGreen:
                    return System.Drawing.Color.FromArgb(34, 120, 51);
                case OlCategoryColor.olCategoryColorTeal:
                    return System.Drawing.Color.FromArgb(0, 171, 169);
                case OlCategoryColor.olCategoryColorDarkTeal:
                    return System.Drawing.Color.FromArgb(0, 130, 114);
                case OlCategoryColor.olCategoryColorOlive:
                    return System.Drawing.Color.FromArgb(143, 157, 74);
                case OlCategoryColor.olCategoryColorDarkOlive:
                    return System.Drawing.Color.FromArgb(107, 114, 59);
                case OlCategoryColor.olCategoryColorBlue:
                    return System.Drawing.Color.FromArgb(55, 118, 199);
                case OlCategoryColor.olCategoryColorDarkBlue:
                    return System.Drawing.Color.FromArgb(35, 72, 139);
                case OlCategoryColor.olCategoryColorPurple:
                    return System.Drawing.Color.FromArgb(131, 92, 175);
                case OlCategoryColor.olCategoryColorDarkPurple:
                    return System.Drawing.Color.FromArgb(95, 60, 140);
                case OlCategoryColor.olCategoryColorMaroon:
                    return System.Drawing.Color.FromArgb(140, 50, 80);
                case OlCategoryColor.olCategoryColorSteel:
                    return System.Drawing.Color.FromArgb(132, 144, 168);
                case OlCategoryColor.olCategoryColorDarkSteel:
                    return System.Drawing.Color.FromArgb(93, 105, 126);
                case OlCategoryColor.olCategoryColorGray:
                    return System.Drawing.Color.FromArgb(160, 160, 160);
                case OlCategoryColor.olCategoryColorDarkGray:
                    return System.Drawing.Color.FromArgb(110, 110, 110);
                default:
                    return System.Drawing.Color.FromArgb(200, 200, 200);
            }
        }
    }
}
