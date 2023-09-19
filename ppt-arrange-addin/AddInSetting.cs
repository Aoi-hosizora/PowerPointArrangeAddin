#nullable enable

namespace ppt_arrange_addin {

    public class AddInSetting {

        private AddInSetting() { }

        private static AddInSetting? _instance;

        public static AddInSetting Instance {
            get {
                _instance ??= new AddInSetting();
                return _instance;
            }
        }

        public bool ShowWordArtGroup { get; set; }
        public bool ShowShapeTextboxGroup { get; set; }
        public bool ShowShapeSizeAndPositionGroup { get; set; }
        public bool ShowReplacePictureGroup { get; set; }
        public bool ShowPictureSizeAndPositionGroup { get; set; }
        public AddInLanguage Language { get; set; }

        public void Load() {
            ShowWordArtGroup = Properties.Settings.Default.showWordArtGroup;
            ShowShapeTextboxGroup = Properties.Settings.Default.showShapeTextboxGroup;
            ShowShapeSizeAndPositionGroup = Properties.Settings.Default.showShapeSizeAndPositionGroup;
            ShowReplacePictureGroup = Properties.Settings.Default.showReplacePictureGroup;
            ShowPictureSizeAndPositionGroup = Properties.Settings.Default.showPictureSizeAndPositionGroup;
            Language = Properties.Settings.Default.language.ToAddInLanguage();
        }

        public void Save() {
            Properties.Settings.Default.showWordArtGroup = ShowWordArtGroup;
            Properties.Settings.Default.showShapeTextboxGroup = ShowShapeTextboxGroup;
            Properties.Settings.Default.showShapeSizeAndPositionGroup = ShowShapeSizeAndPositionGroup;
            Properties.Settings.Default.showReplacePictureGroup = ShowReplacePictureGroup;
            Properties.Settings.Default.showPictureSizeAndPositionGroup = ShowPictureSizeAndPositionGroup;
            Properties.Settings.Default.language = Language.ToLanguageString();
            Properties.Settings.Default.Save();
        }

    }

}
