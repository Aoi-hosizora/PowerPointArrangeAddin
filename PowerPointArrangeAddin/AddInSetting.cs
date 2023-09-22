#nullable enable

namespace PowerPointArrangeAddin {

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
        public bool ShowVideoSizeAndPositionGroup { get; set; }
        public bool ShowAudioSizeAndPositionGroup { get; set; }
        public bool ShowTableSizeAndPositionGroup { get; set; }
        public bool ShowChartSizeAndPositionGroup { get; set; }
        public bool ShowSmartartSizeAndPositionGroup { get; set; }
        public AddInLanguage Language { get; set; }
        public bool LessButtonsForArrangementGroup { get; set; }

        public void Load() {
            ShowWordArtGroup = Properties.Settings.Default.showWordArtGroup;
            ShowShapeTextboxGroup = Properties.Settings.Default.showShapeTextboxGroup;
            ShowShapeSizeAndPositionGroup = Properties.Settings.Default.showShapeSizeAndPositionGroup;
            ShowReplacePictureGroup = Properties.Settings.Default.showReplacePictureGroup;
            ShowPictureSizeAndPositionGroup = Properties.Settings.Default.showPictureSizeAndPositionGroup;
            ShowVideoSizeAndPositionGroup = Properties.Settings.Default.showVideoSizeAndPositionGroup;
            ShowAudioSizeAndPositionGroup = Properties.Settings.Default.showAudioSizeAndPositionGroup;
            ShowTableSizeAndPositionGroup = Properties.Settings.Default.showTableSizeAndPositionGroup;
            ShowChartSizeAndPositionGroup = Properties.Settings.Default.showChartSizeAndPositionGroup;
            ShowSmartartSizeAndPositionGroup = Properties.Settings.Default.showSmartArtSizeAndPositionGroup;
            Language = Properties.Settings.Default.language.ToAddInLanguage();
            LessButtonsForArrangementGroup = Properties.Settings.Default.lessButtonsForArrangementGroup;
        }

        public void Save() {
            Properties.Settings.Default.showWordArtGroup = ShowWordArtGroup;
            Properties.Settings.Default.showShapeTextboxGroup = ShowShapeTextboxGroup;
            Properties.Settings.Default.showShapeSizeAndPositionGroup = ShowShapeSizeAndPositionGroup;
            Properties.Settings.Default.showReplacePictureGroup = ShowReplacePictureGroup;
            Properties.Settings.Default.showPictureSizeAndPositionGroup = ShowPictureSizeAndPositionGroup;
            Properties.Settings.Default.showVideoSizeAndPositionGroup = ShowVideoSizeAndPositionGroup;
            Properties.Settings.Default.showAudioSizeAndPositionGroup = ShowAudioSizeAndPositionGroup;
            Properties.Settings.Default.showTableSizeAndPositionGroup = ShowTableSizeAndPositionGroup;
            Properties.Settings.Default.showChartSizeAndPositionGroup = ShowChartSizeAndPositionGroup;
            Properties.Settings.Default.showSmartArtSizeAndPositionGroup = ShowSmartartSizeAndPositionGroup;
            Properties.Settings.Default.language = Language.ToLanguageString();
            Properties.Settings.Default.lessButtonsForArrangementGroup = LessButtonsForArrangementGroup;
            Properties.Settings.Default.Save();
        }

    }

}
