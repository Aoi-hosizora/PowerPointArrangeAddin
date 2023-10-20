using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;

#nullable enable

namespace PowerPointArrangeAddin.Misc {

    public class AddInDescription {

        private AddInDescription() {
            _culture = Thread.CurrentThread.CurrentCulture;
            UpdateFields(_culture);
        }

        private static AddInDescription? _instance;

        public static AddInDescription Instance {
            get {
                _instance ??= new AddInDescription();
                return _instance;
            }
        }

        private CultureInfo _culture;

        public CultureInfo Culture {
            get => _culture;
            set {
                _culture = value;
                UpdateFields(value);
            }
        }

        public string Title { get; private set; } = "";
        public string VersionKey { get; private set; } = "";
        public string Version { get; private set; } = "";
        public string AuthorKey { get; private set; } = "";
        public string Author { get; private set; } = "";
        public string HomepageKey { get; private set; } = "";
        public string Homepage { get; private set; } = "";
        public string Copyright { get; private set; } = "";

        private static readonly Dictionary<string, string> TitleMap = new() {
            { AddInLanguage.English.ToLanguageString(), "\"PowerPoint Arrangement Assistant Add-in\"" },
            { AddInLanguage.SimplifiedChinese.ToLanguageString(), "【PowerPoint 排列辅助加载项】" },
            { AddInLanguage.TraditionalChinese.ToLanguageString(), "【PowerPoint 排列輔助增益集】" },
            { AddInLanguage.Japanese.ToLanguageString(), "【PowerPoint 配置補助アドイン】" }
        };

        private static readonly Dictionary<string, string> VersionKeyMap = new() {
            { AddInLanguage.English.ToLanguageString(), "Version" },
            { AddInLanguage.SimplifiedChinese.ToLanguageString(), "版本" },
            { AddInLanguage.TraditionalChinese.ToLanguageString(), "版本" },
            { AddInLanguage.Japanese.ToLanguageString(), "バージョン" }
        };

        private static readonly Dictionary<string, string> AuthorKeyMap = new() {
            { AddInLanguage.English.ToLanguageString(), "Author" },
            { AddInLanguage.SimplifiedChinese.ToLanguageString(), "作者" },
            { AddInLanguage.TraditionalChinese.ToLanguageString(), "作者" },
            { AddInLanguage.Japanese.ToLanguageString(), "作者" }
        };

        private static readonly Dictionary<string, string> HomepageKeyMap = new() {
            { AddInLanguage.English.ToLanguageString(), "Homepage" },
            { AddInLanguage.SimplifiedChinese.ToLanguageString(), "主页" },
            { AddInLanguage.TraditionalChinese.ToLanguageString(), "主頁" },
            { AddInLanguage.Japanese.ToLanguageString(), "ホームページ" }
        };

        private void UpdateFields(CultureInfo culture) {
            var name = culture.Name.ToLower();
            Title = TitleMap[name];
            VersionKey = VersionKeyMap[name];
            var ver = Assembly.GetExecutingAssembly().GetName().Version;
            Version = $"{ver.Major}.{ver.Minor}.{ver.Build}";
            AuthorKey = AuthorKeyMap[name];
            Author = "AoiHosizora (https://github.com/Aoi-hosizora)";
            HomepageKey = HomepageKeyMap[name];
            Homepage = "https://github.com/Aoi-hosizora/PowerPointArrangeAddin";
            var att = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
            Copyright = (att.FirstOrDefault() as AssemblyCopyrightAttribute)?.Copyright ?? "";
        }

        public override string ToString() {
            var title = Title;
            var version = $"{VersionKey}: {Version}";
            var author = $"{AuthorKey}: {Author}";
            var homepage = $"{HomepageKey}: {Homepage}";
            var copyright = Copyright;
            return string.Join("\r\n\r\n", title, version, author, homepage, copyright);
        }

    }

}
