using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ppt_arrange_addin {

    public sealed partial class SettingDialog : Form {

        public SettingDialog() {
            InitializeComponent();

            Font = SystemFonts.MessageBoxFont;
            tlpMain.Dock = DockStyle.Fill;
            LoadDescription();
        }

        private void SettingDialog_Load(object sender, EventArgs e) {
            cbxArtWord.Checked = AddInSetting.Instance.ShowWordArtGroup;
            cbxShapeTextbox.Checked = AddInSetting.Instance.ShowShapeTextboxGroup;
            cbxShapeSizeAndPosition.Checked = AddInSetting.Instance.ShowShapeSizeAndPositionGroup;
            cbxReplacePicture.Checked = AddInSetting.Instance.ShowReplacePictureGroup;
            cbxPictureSizeAndPosition.Checked = AddInSetting.Instance.ShowPictureSizeAndPositionGroup;
            cboLanguage.SelectedIndex = AddInSetting.Instance.Language.ToLanguageIndex();
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            AddInSetting.Instance.ShowWordArtGroup = cbxArtWord.Checked;
            AddInSetting.Instance.ShowShapeTextboxGroup = cbxShapeTextbox.Checked;
            AddInSetting.Instance.ShowShapeSizeAndPositionGroup = cbxShapeSizeAndPosition.Checked;
            AddInSetting.Instance.ShowReplacePictureGroup = cbxReplacePicture.Checked;
            AddInSetting.Instance.ShowPictureSizeAndPositionGroup = cbxPictureSizeAndPosition.Checked;
            AddInSetting.Instance.Language = cboLanguage.SelectedIndex.ToAddInLanguage();
            AddInSetting.Instance.Save();
            AddInLanguageChanger.ChangeLanguage(AddInSetting.Instance.Language);
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private readonly string _addInTitle = "\"Arrangement Assistant Add-in\"";
        private readonly string _addInVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        private readonly string _addInAuthor = "AoiHosizora (https://github.com/Aoi-hosizora)";
        private readonly string _addInHomepage = "https://github.com/Aoi-hosizora/ppt-arrange-addin";

        private void LoadDescription() {
            var title = GetResourceString(key: "_title", defaultValue: _addInTitle);
            var version = $"{GetResourceString(key: "_version", defaultValue: "Version")}: {_addInVersion}";
            var author = $"{GetResourceString(key: "_author", defaultValue: "Author")}: {_addInAuthor}";
            var homepage = $"{GetResourceString(key: "_homepage", defaultValue: "Homepage")}: {_addInHomepage}";
            var copyright = GetAttributeFromAssembly<AssemblyCopyrightAttribute>()?.Copyright ?? "";
            var description = $"{title}\r\n\r\n{version}\r\n\r\n{author}\r\n\r\n{homepage}\r\n\r\n{copyright}";
            tbxDescription.Text = description;
        }

        private static T GetAttributeFromAssembly<T>(T defaultValue = default) {
            var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(T), false);
            return attributes.Length > 0 ? (T) attributes[0] : defaultValue;
        }

        private static string GetResourceString(string key, string defaultValue) {
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingDialog));
            return resources.GetString(key) ?? defaultValue;
        }

    }

}
