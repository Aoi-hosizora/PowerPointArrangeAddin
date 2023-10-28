using System;
using System.Globalization;
using System.Threading;

#nullable enable

namespace PowerPointArrangeAddin.Misc {

    public enum AddInLanguage {
        Default,
        English,
        SimplifiedChinese,
        TraditionalChinese,
        Japanese
    }

    public static class AddInLanguageChanger {

        private static AddInLanguage? _defaultLanguage; // will never be `AddInLanguage.Default`
        private static Action? _uiInvalidator; // used to invalidate ui when language is changed

        public static void RegisterAddIn(int defaultLanguageId, Action uiInvalidator) {
            _defaultLanguage = new CultureInfo(defaultLanguageId).ToAddInLanguage();
            _uiInvalidator = uiInvalidator; // TODO use instance instead
        }

        public static void ChangeLanguage(AddInLanguage language) {
            if (language == AddInLanguage.Default) {
                language = _defaultLanguage ?? AddInLanguage.English;
            }
            var cultureInfo = new CultureInfo(language.ToLanguageString());
            Thread.CurrentThread.CurrentCulture = cultureInfo;
            Thread.CurrentThread.CurrentUICulture = cultureInfo;
            Properties.Resources.Culture = cultureInfo;
            Ribbon.ArrangeRibbonResources.Culture = cultureInfo;
            MiscResources.Culture = cultureInfo;
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

        private static AddInLanguage ToAddInLanguage(this CultureInfo culture) {
            return culture.Name.ToLower() switch {
                "zh" or "zh-hans" or "zh-chs" or "zh-cn" or "zh-sg" => AddInLanguage.SimplifiedChinese,
                "zh-hant" or "zh-cht" or "zh-tw" or "zh-hk" or "zh-mo" => AddInLanguage.TraditionalChinese,
                "ja" or "ja-jp" => AddInLanguage.Japanese,
                _ => AddInLanguage.English
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
