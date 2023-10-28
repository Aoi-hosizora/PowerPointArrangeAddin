using System.Linq;
using System.Reflection;

#nullable enable

namespace PowerPointArrangeAddin.Misc {

    public class AddInDescription {

        private AddInDescription() { }

        private static AddInDescription? _instance;

        public static AddInDescription Instance {
            get {
                _instance ??= new AddInDescription();
                return _instance;
            }
        }

        public string Title { get; private set; } = "";
        private string TitleWrapper { get; set; } = "";
        private string VersionKey { get; set; } = "";
        public string Version { get; private set; } = "";
        private string AuthorKey { get; set; } = "";
        public string Author { get; private set; } = "";
        private string HomepageKey { get; set; } = "";
        public string Homepage { get; private set; } = "";
        public string Copyright { get; private set; } = "";

        public void UpdateFields() {
            Title = MiscResources.Desc_Title;
            TitleWrapper = MiscResources.Desc_TitleWrapper;
            VersionKey = MiscResources.Desc_VersionKey;
            Version = AddInVersion.Instance.GetAssemblyVersionInString();
            AuthorKey = MiscResources.Desc_AuthorKey;
            Author = MiscResources.Desc_Author;
            HomepageKey = MiscResources.Desc_HomepageKey;
            Homepage = MiscResources.Desc_Homepage;
            var att = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
            Copyright = (att.FirstOrDefault() as AssemblyCopyrightAttribute)?.Copyright ?? "";
        }

        public override string ToString() {
            var title = $"{TitleWrapper[0]}{Title}{TitleWrapper[1]}";
            var version = $"{VersionKey}: v{Version}";
            var author = $"{AuthorKey}: {Author}";
            var homepage = $"{HomepageKey}: {Homepage}";
            var copyright = Copyright;
            return string.Join("\r\n\r\n", title, version, author, homepage, copyright);
        }

    }

}
