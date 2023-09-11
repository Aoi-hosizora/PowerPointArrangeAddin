
namespace ppt_arrange_addin {
    sealed partial class SettingDialog {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.cbxArtWord = new System.Windows.Forms.CheckBox();
            this.cbxShapeTextbox = new System.Windows.Forms.CheckBox();
            this.cbxShapeSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.cbxReplacePicture = new System.Windows.Forms.CheckBox();
            this.cbxPictureSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.tlpMain = new System.Windows.Forms.TableLayoutPanel();
            this.grpGroupVisibility = new System.Windows.Forms.GroupBox();
            this.tlpGroupVisibility = new System.Windows.Forms.TableLayoutPanel();
            this.grpOtherSetting = new System.Windows.Forms.GroupBox();
            this.tlpOtherSetting = new System.Windows.Forms.TableLayoutPanel();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.cboLanguage = new System.Windows.Forms.ComboBox();
            this.tbxDescription = new System.Windows.Forms.TextBox();
            this.tlpButton = new System.Windows.Forms.TableLayoutPanel();
            this.tlpMain.SuspendLayout();
            this.grpGroupVisibility.SuspendLayout();
            this.tlpGroupVisibility.SuspendLayout();
            this.grpOtherSetting.SuspendLayout();
            this.tlpOtherSetting.SuspendLayout();
            this.tlpButton.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnOK.Location = new System.Drawing.Point(5, 5);
            this.btnOK.Margin = new System.Windows.Forms.Padding(0, 0, 3, 0);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnCancel.Location = new System.Drawing.Point(86, 5);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // cbxArtWord
            // 
            this.cbxArtWord.AutoSize = true;
            this.cbxArtWord.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.cbxArtWord.Location = new System.Drawing.Point(0, 3);
            this.cbxArtWord.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.cbxArtWord.Name = "cbxArtWord";
            this.cbxArtWord.Size = new System.Drawing.Size(194, 17);
            this.cbxArtWord.TabIndex = 2;
            this.cbxArtWord.Text = "Show &art word group (home tab)";
            this.cbxArtWord.UseVisualStyleBackColor = true;
            // 
            // cbxShapeTextbox
            // 
            this.cbxShapeTextbox.AutoSize = true;
            this.cbxShapeTextbox.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.cbxShapeTextbox.Location = new System.Drawing.Point(0, 26);
            this.cbxShapeTextbox.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.cbxShapeTextbox.Name = "cbxShapeTextbox";
            this.cbxShapeTextbox.Size = new System.Drawing.Size(193, 17);
            this.cbxShapeTextbox.TabIndex = 2;
            this.cbxShapeTextbox.Text = "Show &textbox group (shape tab)";
            this.cbxShapeTextbox.UseVisualStyleBackColor = true;
            // 
            // cbxShapeSizeAndPosition
            // 
            this.cbxShapeSizeAndPosition.AutoSize = true;
            this.cbxShapeSizeAndPosition.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.cbxShapeSizeAndPosition.Location = new System.Drawing.Point(0, 49);
            this.cbxShapeSizeAndPosition.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.cbxShapeSizeAndPosition.Name = "cbxShapeSizeAndPosition";
            this.cbxShapeSizeAndPosition.Size = new System.Drawing.Size(241, 17);
            this.cbxShapeSizeAndPosition.TabIndex = 2;
            this.cbxShapeSizeAndPosition.Text = "Show &size and position group (shape tab)";
            this.cbxShapeSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // cbxReplacePicture
            // 
            this.cbxReplacePicture.AutoSize = true;
            this.cbxReplacePicture.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.cbxReplacePicture.Location = new System.Drawing.Point(0, 72);
            this.cbxReplacePicture.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.cbxReplacePicture.Name = "cbxReplacePicture";
            this.cbxReplacePicture.Size = new System.Drawing.Size(236, 17);
            this.cbxReplacePicture.TabIndex = 2;
            this.cbxReplacePicture.Text = "Show &replace picture group (picture tab)";
            this.cbxReplacePicture.UseVisualStyleBackColor = true;
            // 
            // cbxPictureSizeAndPosition
            // 
            this.cbxPictureSizeAndPosition.AutoSize = true;
            this.cbxPictureSizeAndPosition.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.cbxPictureSizeAndPosition.Location = new System.Drawing.Point(0, 95);
            this.cbxPictureSizeAndPosition.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.cbxPictureSizeAndPosition.Name = "cbxPictureSizeAndPosition";
            this.cbxPictureSizeAndPosition.Size = new System.Drawing.Size(246, 17);
            this.cbxPictureSizeAndPosition.TabIndex = 2;
            this.cbxPictureSizeAndPosition.Text = "Show size and &position group (picture tab)";
            this.cbxPictureSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // tlpMain
            // 
            this.tlpMain.AutoSize = true;
            this.tlpMain.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tlpMain.ColumnCount = 1;
            this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tlpMain.Controls.Add(this.grpGroupVisibility, 0, 0);
            this.tlpMain.Controls.Add(this.grpOtherSetting, 0, 1);
            this.tlpMain.Controls.Add(this.tbxDescription, 0, 2);
            this.tlpMain.Controls.Add(this.tlpButton, 0, 3);
            this.tlpMain.Location = new System.Drawing.Point(12, 12);
            this.tlpMain.Name = "tlpMain";
            this.tlpMain.RowCount = 4;
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.Size = new System.Drawing.Size(352, 364);
            this.tlpMain.TabIndex = 3;
            // 
            // grpGroupVisibility
            // 
            this.grpGroupVisibility.AutoSize = true;
            this.grpGroupVisibility.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.grpGroupVisibility.Controls.Add(this.tlpGroupVisibility);
            this.grpGroupVisibility.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpGroupVisibility.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.grpGroupVisibility.Location = new System.Drawing.Point(8, 8);
            this.grpGroupVisibility.Margin = new System.Windows.Forms.Padding(8, 8, 8, 0);
            this.grpGroupVisibility.Name = "grpGroupVisibility";
            this.grpGroupVisibility.Padding = new System.Windows.Forms.Padding(10, 2, 10, 5);
            this.grpGroupVisibility.Size = new System.Drawing.Size(336, 134);
            this.grpGroupVisibility.TabIndex = 6;
            this.grpGroupVisibility.TabStop = false;
            this.grpGroupVisibility.Text = "Group visibility";
            // 
            // tlpGroupVisibility
            // 
            this.tlpGroupVisibility.AutoSize = true;
            this.tlpGroupVisibility.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tlpGroupVisibility.BackColor = System.Drawing.SystemColors.Control;
            this.tlpGroupVisibility.ColumnCount = 1;
            this.tlpGroupVisibility.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tlpGroupVisibility.Controls.Add(this.cbxArtWord, 0, 0);
            this.tlpGroupVisibility.Controls.Add(this.cbxShapeTextbox, 0, 1);
            this.tlpGroupVisibility.Controls.Add(this.cbxPictureSizeAndPosition, 0, 4);
            this.tlpGroupVisibility.Controls.Add(this.cbxReplacePicture, 0, 3);
            this.tlpGroupVisibility.Controls.Add(this.cbxShapeSizeAndPosition, 0, 2);
            this.tlpGroupVisibility.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpGroupVisibility.Location = new System.Drawing.Point(10, 14);
            this.tlpGroupVisibility.Name = "tlpGroupVisibility";
            this.tlpGroupVisibility.RowCount = 5;
            this.tlpGroupVisibility.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpGroupVisibility.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpGroupVisibility.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpGroupVisibility.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpGroupVisibility.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpGroupVisibility.Size = new System.Drawing.Size(316, 115);
            this.tlpGroupVisibility.TabIndex = 4;
            // 
            // grpOtherSetting
            // 
            this.grpOtherSetting.AutoSize = true;
            this.grpOtherSetting.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.grpOtherSetting.Controls.Add(this.tlpOtherSetting);
            this.grpOtherSetting.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpOtherSetting.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.grpOtherSetting.Location = new System.Drawing.Point(8, 150);
            this.grpOtherSetting.Margin = new System.Windows.Forms.Padding(8, 8, 8, 0);
            this.grpOtherSetting.Name = "grpOtherSetting";
            this.grpOtherSetting.Padding = new System.Windows.Forms.Padding(10, 2, 10, 5);
            this.grpOtherSetting.Size = new System.Drawing.Size(336, 47);
            this.grpOtherSetting.TabIndex = 4;
            this.grpOtherSetting.TabStop = false;
            this.grpOtherSetting.Text = "Other setting";
            // 
            // tlpOtherSetting
            // 
            this.tlpOtherSetting.AutoSize = true;
            this.tlpOtherSetting.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tlpOtherSetting.ColumnCount = 2;
            this.tlpOtherSetting.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tlpOtherSetting.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpOtherSetting.Controls.Add(this.lblLanguage, 0, 0);
            this.tlpOtherSetting.Controls.Add(this.cboLanguage, 1, 0);
            this.tlpOtherSetting.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpOtherSetting.Location = new System.Drawing.Point(10, 14);
            this.tlpOtherSetting.Name = "tlpOtherSetting";
            this.tlpOtherSetting.RowCount = 1;
            this.tlpOtherSetting.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpOtherSetting.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tlpOtherSetting.Size = new System.Drawing.Size(316, 28);
            this.tlpOtherSetting.TabIndex = 4;
            // 
            // lblLanguage
            // 
            this.lblLanguage.AutoSize = true;
            this.lblLanguage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblLanguage.Location = new System.Drawing.Point(0, 4);
            this.lblLanguage.Margin = new System.Windows.Forms.Padding(0, 4, 0, 4);
            this.lblLanguage.Name = "lblLanguage";
            this.lblLanguage.Size = new System.Drawing.Size(59, 20);
            this.lblLanguage.TabIndex = 1;
            this.lblLanguage.Text = "Language: ";
            this.lblLanguage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cboLanguage
            // 
            this.cboLanguage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cboLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboLanguage.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.cboLanguage.FormattingEnabled = true;
            this.cboLanguage.Location = new System.Drawing.Point(59, 4);
            this.cboLanguage.Margin = new System.Windows.Forms.Padding(0, 4, 0, 4);
            this.cboLanguage.Name = "cboLanguage";
            this.cboLanguage.Size = new System.Drawing.Size(257, 20);
            this.cboLanguage.TabIndex = 0;
            // 
            // tbxDescription
            // 
            this.tbxDescription.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbxDescription.Location = new System.Drawing.Point(8, 205);
            this.tbxDescription.Margin = new System.Windows.Forms.Padding(8, 8, 8, 0);
            this.tbxDescription.Multiline = true;
            this.tbxDescription.Name = "tbxDescription";
            this.tbxDescription.ReadOnly = true;
            this.tbxDescription.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tbxDescription.Size = new System.Drawing.Size(336, 120);
            this.tbxDescription.TabIndex = 4;
            this.tbxDescription.Text = "【Arrangement Add-in for Microsoft Office PowerPoint】\r\n\r\nVersion: v1.0.0\r\n\r\nAuthor" +
    ": AoiHosizora (https://gist.github.com/Aoi-hosizora)\r\n\r\nHomepage: https://github" +
    ".com/Aoi-hosizora/ppt-arrange-addin";
            // 
            // tlpButton
            // 
            this.tlpButton.AutoSize = true;
            this.tlpButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tlpButton.ColumnCount = 2;
            this.tlpButton.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tlpButton.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tlpButton.Controls.Add(this.btnCancel, 1, 0);
            this.tlpButton.Controls.Add(this.btnOK, 0, 0);
            this.tlpButton.Dock = System.Windows.Forms.DockStyle.Right;
            this.tlpButton.Location = new System.Drawing.Point(183, 328);
            this.tlpButton.Name = "tlpButton";
            this.tlpButton.Padding = new System.Windows.Forms.Padding(5);
            this.tlpButton.RowCount = 1;
            this.tlpButton.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpButton.Size = new System.Drawing.Size(166, 33);
            this.tlpButton.TabIndex = 4;
            // 
            // SettingDialog
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(376, 418);
            this.Controls.Add(this.tlpMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Arrangement Add-in Setting";
            this.Load += new System.EventHandler(this.SettingDialog_Load);
            this.tlpMain.ResumeLayout(false);
            this.tlpMain.PerformLayout();
            this.grpGroupVisibility.ResumeLayout(false);
            this.grpGroupVisibility.PerformLayout();
            this.tlpGroupVisibility.ResumeLayout(false);
            this.tlpGroupVisibility.PerformLayout();
            this.grpOtherSetting.ResumeLayout(false);
            this.grpOtherSetting.PerformLayout();
            this.tlpOtherSetting.ResumeLayout(false);
            this.tlpOtherSetting.PerformLayout();
            this.tlpButton.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.CheckBox cbxArtWord;
        private System.Windows.Forms.CheckBox cbxShapeTextbox;
        private System.Windows.Forms.CheckBox cbxShapeSizeAndPosition;
        private System.Windows.Forms.CheckBox cbxReplacePicture;
        private System.Windows.Forms.CheckBox cbxPictureSizeAndPosition;
        private System.Windows.Forms.TableLayoutPanel tlpMain;
        private System.Windows.Forms.GroupBox grpGroupVisibility;
        private System.Windows.Forms.GroupBox grpOtherSetting;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.ComboBox cboLanguage;
        private System.Windows.Forms.TableLayoutPanel tlpGroupVisibility;
        private System.Windows.Forms.TableLayoutPanel tlpOtherSetting;
        private System.Windows.Forms.TextBox tbxDescription;
        private System.Windows.Forms.TableLayoutPanel tlpButton;
    }
}