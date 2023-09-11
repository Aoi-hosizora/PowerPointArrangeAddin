using System;
using System.Drawing;
using System.Windows.Forms;

namespace ppt_arrange_addin {

    public sealed partial class SettingDialog : Form {

        public SettingDialog() {
            InitializeComponent();
            Font = SystemFonts.MessageBoxFont;
            tlpMain.Dock = DockStyle.Fill;
        }

        private void SettingDialog_Load(object sender, EventArgs e) {
            cbxArtWord.Checked = Properties.Settings.Default.showWordArtGroup;
            cbxShapeTextbox.Checked = Properties.Settings.Default.showShapeTextboxGroup;
            cbxShapeSizeAndPosition.Checked = Properties.Settings.Default.showShapeSizeAndPositionGroup;
            cbxReplacePicture.Checked = Properties.Settings.Default.showReplacePictureGroup;
            cbxPictureSizeAndPosition.Checked = Properties.Settings.Default.showPictureSizeAndPositionGroup;
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            Properties.Settings.Default.showWordArtGroup = cbxArtWord.Checked;
            Properties.Settings.Default.showShapeTextboxGroup = cbxShapeTextbox.Checked;
            Properties.Settings.Default.showShapeSizeAndPositionGroup = cbxShapeSizeAndPosition.Checked;
            Properties.Settings.Default.showReplacePictureGroup = cbxReplacePicture.Checked;
            Properties.Settings.Default.showPictureSizeAndPositionGroup = cbxPictureSizeAndPosition.Checked;
            Properties.Settings.Default.Save();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void BtnHomepage_Click(object sender, EventArgs e) {
            const string url = "https://github.com/Aoi-hosizora/ppt-arrange-addin";
            try {
                System.Diagnostics.Process.Start(url);
            } catch (Exception) {
                // ignored
            }
        }

    }

}
