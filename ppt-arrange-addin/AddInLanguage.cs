using System;
using System.Collections.Generic;
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
        private static Action _uiInvalidater;

        public static void RegisterAddIn(int defaultLanguageId, Action uiInvalidater) {
            _defaultLanguageId = defaultLanguageId;
            _uiInvalidater = uiInvalidater;
        }

        public static void ChangeLanguage(AddInLanguage language) {
            var cultureInfo = language switch {
                AddInLanguage.Default => new CultureInfo(_defaultLanguageId),
                AddInLanguage.English => new CultureInfo("en"),
                AddInLanguage.SimplifiedChinese => new CultureInfo("zh-hans"),
                AddInLanguage.TraditionalChinese => new CultureInfo("zh-hant"),
                AddInLanguage.Japanese => new CultureInfo("ja"),
                _ => new CultureInfo(_defaultLanguageId)
            };
            Thread.CurrentThread.CurrentCulture = cultureInfo;
            Thread.CurrentThread.CurrentUICulture = cultureInfo;
            Properties.Resources.Culture = cultureInfo;
            ArrangeRibbonResources.Culture = cultureInfo;
            _uiInvalidater?.Invoke();
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
