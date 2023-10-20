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
            chkArrangement.Checked = true;
            chkShapeTextbox.Checked = AddInSetting.Instance.ShowShapeTextboxGroup;
            chkShapeSizeAndPosition.Checked = AddInSetting.Instance.ShowShapeSizeAndPositionGroup;
            chkReplacePicture.Checked = AddInSetting.Instance.ShowReplacePictureGroup;
            chkPictureSizeAndPosition.Checked = AddInSetting.Instance.ShowPictureSizeAndPositionGroup;
            chkVideoSizeAndPosition.Checked = AddInSetting.Instance.ShowVideoSizeAndPositionGroup;
            chkAudioSizeAndPosition.Checked = AddInSetting.Instance.ShowAudioSizeAndPositionGroup;
            chkTableSizeAndPosition.Checked = AddInSetting.Instance.ShowTableSizeAndPositionGroup;
            chkChartSizeAndPosition.Checked = AddInSetting.Instance.ShowChartSizeAndPositionGroup;
            chkSmartartSizeAndPosition.Checked = AddInSetting.Instance.ShowSmartartSizeAndPositionGroup;
            cboLanguage.SelectedIndex = AddInSetting.Instance.Language.ToLanguageIndex();
            chkLessButtonsForArrange.Checked = AddInSetting.Instance.LessButtonsForArrangementGroup;
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            AddInSetting.Instance.ShowWordArtGroup = chkWordArt.Checked;
            AddInSetting.Instance.ShowShapeTextboxGroup = chkShapeTextbox.Checked;
            AddInSetting.Instance.ShowShapeSizeAndPositionGroup = chkShapeSizeAndPosition.Checked;
            AddInSetting.Instance.ShowReplacePictureGroup = chkReplacePicture.Checked;
            AddInSetting.Instance.ShowPictureSizeAndPositionGroup = chkPictureSizeAndPosition.Checked;
            AddInSetting.Instance.ShowVideoSizeAndPositionGroup = chkVideoSizeAndPosition.Checked;
            AddInSetting.Instance.ShowAudioSizeAndPositionGroup = chkAudioSizeAndPosition.Checked;
            AddInSetting.Instance.ShowTableSizeAndPositionGroup = chkTableSizeAndPosition.Checked;
            AddInSetting.Instance.ShowChartSizeAndPositionGroup = chkChartSizeAndPosition.Checked;
            AddInSetting.Instance.ShowSmartartSizeAndPositionGroup = chkSmartartSizeAndPosition.Checked;
            AddInSetting.Instance.Language = cboLanguage.SelectedIndex.ToAddInLanguage();
            AddInSetting.Instance.LessButtonsForArrangementGroup = chkLessButtonsForArrange.Checked;
            AddInSetting.Instance.Save();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.Cancel;
            Close();
        }

    }

}
