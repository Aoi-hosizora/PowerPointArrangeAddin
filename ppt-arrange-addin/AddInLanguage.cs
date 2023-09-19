using System;
using System.Globalization;
using System.Threading;

namespace ppt_arrange_addin {

    public enum AddInLanguage {
        Default,
        English,
        SimplifiedChinese,
        TraditionalChinese,
        Japanese
    }

    public static class AddInLanguageChanger {

        private static int _defaultLanguageId;
        private static Action _uiInvalidator;

        public static void RegisterAddIn(int defaultLanguageId, Action uiInvalidator) {
            _defaultLanguageId = defaultLanguageId;
            _uiInvalidator = uiInvalidator;
        }

        public static void ChangeLanguage(AddInLanguage language) {
            var cultureInfo = language == AddInLanguage.Default
                ? new CultureInfo(_defaultLanguageId)
                : new CultureInfo(language.ToLanguageString());
            Thread.CurrentThread.CurrentCulture = cultureInfo;
            Thread.CurrentThread.CurrentUICulture = cultureInfo;
            Properties.Resources.Culture = cultureInfo;
            Ribbon.ArrangeRibbonResources.Culture = cultureInfo;
            _uiInvalidator?.Invoke();
        }

        #region Extensions

        public static AddInLanguage ToAddInLanguage(this string language) {
            return language switch {
                "default" => AddInLanguage.Default,
                "en" => AddInLanguage.English,
                "zh-hans" => AddInLanguage.SimplifiedChinese,
                "zh-hant" => AddInLanguage.TraditionalChinese,
                "ja" => AddInLanguage.Japanese,
                _ => AddInLanguage.Default
            };
        }

        public static string ToLanguageString(this AddInLanguage language) {
            return language switch {
                AddInLanguage.Default => "default",
                AddInLanguage.English => "en",
                AddInLanguage.SimplifiedChinese => "zh-hans",
                AddInLanguage.TraditionalChinese => "zh-hant",
                AddInLanguage.Japanese => "ja",
                _ => "default"
            };
        }

        public static AddInLanguage ToAddInLanguage(this int language) {
            return language switch {
                0 => AddInLanguage.Default,
                1 => AddInLanguage.English,
                2 => AddInLanguage.SimplifiedChinese,
                3 => AddInLanguage.TraditionalChinese,
                4 => AddInLanguage.Japanese,
                _ => AddInLanguage.Default
            };
        }

        public static int ToLanguageIndex(this AddInLanguage language) {
            return language switch {
                AddInLanguage.Default => 0,
                AddInLanguage.English => 1,
                AddInLanguage.SimplifiedChinese => 2,
                AddInLanguage.TraditionalChinese => 3,
                AddInLanguage.Japanese => 4,
                _ => 0
            };
        }

        #endregion

    }

}
