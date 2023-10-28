#nullable enable

namespace PowerPointArrangeAddin.Misc {

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
        public bool ShowArrangementGroup { get; set; }
        public bool ShowShapeTextboxGroup { get; set; }
        public bool ShowReplacePictureGroup { get; set; }
        public bool ShowSizeAndPositionGroup { get; set; }
        public bool ShowShapeSizeAndPositionGroup { get; set; }
        public bool ShowPictureSizeAndPositionGroup { get; set; }
        public bool ShowVideoSizeAndPositionGroup { get; set; }
        public bool ShowAudioSizeAndPositionGroup { get; set; }
        public bool ShowTableSizeAndPositionGroup { get; set; }
        public bool ShowChartSizeAndPositionGroup { get; set; }
        public bool ShowSmartartSizeAndPositionGroup { get; set; }
        public AddInLanguage Language { get; set; }
        public bool CheckUpdateWhenStartUp { get; set; }
        public string IgnoreUpdateRecord { get; set; } = "";
        public bool LessButtonsForArrangementGroup { get; set; }
        public bool HideMarginSettingForTextboxGroup { get; set; }

        public void Load() {
            ShowWordArtGroup = Properties.Settings.Default.showWordArtGroup;
            ShowArrangementGroup = Properties.Settings.Default.showArrangementGroup;
            ShowShapeTextboxGroup = Properties.Settings.Default.showShapeTextboxGroup;
            ShowReplacePictureGroup = Properties.Settings.Default.showReplacePictureGroup;
            ShowSizeAndPositionGroup = Properties.Settings.Default.showSizeAndPositionGroup;
            ShowShapeSizeAndPositionGroup = Properties.Settings.Default.showShapeSizeAndPositionGroup;
            ShowPictureSizeAndPositionGroup = Properties.Settings.Default.showPictureSizeAndPositionGroup;
            ShowVideoSizeAndPositionGroup = Properties.Settings.Default.showVideoSizeAndPositionGroup;
            ShowAudioSizeAndPositionGroup = Properties.Settings.Default.showAudioSizeAndPositionGroup;
            ShowTableSizeAndPositionGroup = Properties.Settings.Default.showTableSizeAndPositionGroup;
            ShowChartSizeAndPositionGroup = Properties.Settings.Default.showChartSizeAndPositionGroup;
            ShowSmartartSizeAndPositionGroup = Properties.Settings.Default.showSmartArtSizeAndPositionGroup;
            Language = Properties.Settings.Default.language.ToAddInLanguage();
            CheckUpdateWhenStartUp = Properties.Settings.Default.checkUpdateWhenStartUp;
            IgnoreUpdateRecord = Properties.Settings.Default.ignoreUpdateRecord;
            LessButtonsForArrangementGroup = Properties.Settings.Default.lessButtonsForArrangementGroup;
            HideMarginSettingForTextboxGroup = Properties.Settings.Default.hideMarginSettingForTextboxGroup;
        }

        public void Save() {
            Properties.Settings.Default.showWordArtGroup = ShowWordArtGroup;
            Properties.Settings.Default.showArrangementGroup = ShowArrangementGroup;
            Properties.Settings.Default.showShapeTextboxGroup = ShowShapeTextboxGroup;
            Properties.Settings.Default.showReplacePictureGroup = ShowReplacePictureGroup;
            Properties.Settings.Default.showSizeAndPositionGroup = ShowSizeAndPositionGroup;
            Properties.Settings.Default.showShapeSizeAndPositionGroup = ShowShapeSizeAndPositionGroup;
            Properties.Settings.Default.showPictureSizeAndPositionGroup = ShowPictureSizeAndPositionGroup;
            Properties.Settings.Default.showVideoSizeAndPositionGroup = ShowVideoSizeAndPositionGroup;
            Properties.Settings.Default.showAudioSizeAndPositionGroup = ShowAudioSizeAndPositionGroup;
            Properties.Settings.Default.showTableSizeAndPositionGroup = ShowTableSizeAndPositionGroup;
            Properties.Settings.Default.showChartSizeAndPositionGroup = ShowChartSizeAndPositionGroup;
            Properties.Settings.Default.showSmartArtSizeAndPositionGroup = ShowSmartartSizeAndPositionGroup;
            Properties.Settings.Default.language = Language.ToLanguageString();
            Properties.Settings.Default.checkUpdateWhenStartUp = CheckUpdateWhenStartUp;
            Properties.Settings.Default.ignoreUpdateRecord = IgnoreUpdateRecord;
            Properties.Settings.Default.lessButtonsForArrangementGroup = LessButtonsForArrangementGroup;
            Properties.Settings.Default.hideMarginSettingForTextboxGroup = HideMarginSettingForTextboxGroup;
            Properties.Settings.Default.Save();
        }

        public bool ShowShapeSizeAndPositionGroup2 => ShowSizeAndPositionGroup && ShowShapeSizeAndPositionGroup;
        public bool ShowPictureSizeAndPositionGroup2 => ShowSizeAndPositionGroup && ShowPictureSizeAndPositionGroup;
        public bool ShowVideoSizeAndPositionGroup2 => ShowSizeAndPositionGroup && ShowVideoSizeAndPositionGroup;
        public bool ShowAudioSizeAndPositionGroup2 => ShowSizeAndPositionGroup && ShowAudioSizeAndPositionGroup;
        public bool ShowTableSizeAndPositionGroup2 => ShowSizeAndPositionGroup && ShowTableSizeAndPositionGroup;
        public bool ShowChartSizeAndPositionGroup2 => ShowSizeAndPositionGroup && ShowChartSizeAndPositionGroup;
        public bool ShowSmartartSizeAndPositionGroup2 => ShowSizeAndPositionGroup && ShowSmartartSizeAndPositionGroup;

    }

}
