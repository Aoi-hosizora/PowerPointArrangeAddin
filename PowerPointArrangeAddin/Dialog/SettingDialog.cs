using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPointArrangeAddin.Misc;

namespace PowerPointArrangeAddin.Dialog {

    public sealed partial class SettingDialog : Form {

        public SettingDialog() {
            InitializeComponent();

            tlpMain.AutoSize = true;
            tlpMain.Dock = DockStyle.Fill;

            AutoScaleMode = AutoScaleMode.Dpi;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            AutoSize = true;
            Font = SystemFonts.MessageBoxFont;

            tbxDescription.Text = AddInDescription.Instance.ToString();
        }

        private void SettingDialog_Load(object sender, EventArgs e) {
            chkWordArt.Checked = AddInSetting.Instance.ShowWordArtGroup;
            chkArrangement.Checked = AddInSetting.Instance.ShowArrangementGroup;
            chkShapeTextbox.Checked = AddInSetting.Instance.ShowShapeTextboxGroup;
            chkReplacePicture.Checked = AddInSetting.Instance.ShowReplacePictureGroup;
            chkSizeAndPosition.Checked = AddInSetting.Instance.ShowSizeAndPositionGroup;
            chkShapeSizeAndPosition.Checked = AddInSetting.Instance.ShowShapeSizeAndPositionGroup;
            chkPictureSizeAndPosition.Checked = AddInSetting.Instance.ShowPictureSizeAndPositionGroup;
            chkVideoSizeAndPosition.Checked = AddInSetting.Instance.ShowVideoSizeAndPositionGroup;
            chkAudioSizeAndPosition.Checked = AddInSetting.Instance.ShowAudioSizeAndPositionGroup;
            chkTableSizeAndPosition.Checked = AddInSetting.Instance.ShowTableSizeAndPositionGroup;
            chkChartSizeAndPosition.Checked = AddInSetting.Instance.ShowChartSizeAndPositionGroup;
            chkSmartartSizeAndPosition.Checked = AddInSetting.Instance.ShowSmartartSizeAndPositionGroup;
            cboLanguage.SelectedIndex = AddInSetting.Instance.Language.ToLanguageIndex();
            chkCheckUpdateWhenStartUp.Checked = AddInSetting.Instance.CheckUpdateWhenStartUp;
            chkLessButtonsForArrangement.Checked = AddInSetting.Instance.LessButtonsForArrangementGroup;
            chkHideMarginSettingForTextbox.Checked = AddInSetting.Instance.HideMarginSettingForTextboxGroup;
            cboIconStyle.SelectedIndex = AddInSetting.Instance.IconStyle.ToIconStyleIndex();
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            AddInSetting.Instance.ShowWordArtGroup = chkWordArt.Checked;
            AddInSetting.Instance.ShowArrangementGroup = chkArrangement.Checked;
            AddInSetting.Instance.ShowShapeTextboxGroup = chkShapeTextbox.Checked;
            AddInSetting.Instance.ShowReplacePictureGroup = chkReplacePicture.Checked;
            AddInSetting.Instance.ShowSizeAndPositionGroup = chkSizeAndPosition.Checked;
            AddInSetting.Instance.ShowShapeSizeAndPositionGroup = chkShapeSizeAndPosition.Checked;
            AddInSetting.Instance.ShowPictureSizeAndPositionGroup = chkPictureSizeAndPosition.Checked;
            AddInSetting.Instance.ShowVideoSizeAndPositionGroup = chkVideoSizeAndPosition.Checked;
            AddInSetting.Instance.ShowAudioSizeAndPositionGroup = chkAudioSizeAndPosition.Checked;
            AddInSetting.Instance.ShowTableSizeAndPositionGroup = chkTableSizeAndPosition.Checked;
            AddInSetting.Instance.ShowChartSizeAndPositionGroup = chkChartSizeAndPosition.Checked;
            AddInSetting.Instance.ShowSmartartSizeAndPositionGroup = chkSmartartSizeAndPosition.Checked;
            AddInSetting.Instance.Language = cboLanguage.SelectedIndex.ToAddInLanguage();
            AddInSetting.Instance.CheckUpdateWhenStartUp = chkCheckUpdateWhenStartUp.Checked;
            AddInSetting.Instance.LessButtonsForArrangementGroup = chkLessButtonsForArrangement.Checked;
            AddInSetting.Instance.HideMarginSettingForTextboxGroup = chkHideMarginSettingForTextbox.Checked;
            AddInSetting.Instance.IconStyle = cboIconStyle.SelectedIndex.ToAddInIconStyle();
            AddInSetting.Instance.Save();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void ChkSizeAndPosition_CheckedChanged(object sender, EventArgs e) {
            var check = chkSizeAndPosition.Checked;
            chkShapeSizeAndPosition.Enabled = check;
            chkPictureSizeAndPosition.Enabled = check;
            chkVideoSizeAndPosition.Enabled = check;
            chkAudioSizeAndPosition.Enabled = check;
            chkTableSizeAndPosition.Enabled = check;
            chkChartSizeAndPosition.Enabled = check;
            chkSmartartSizeAndPosition.Enabled = check;
        }

    }

}
