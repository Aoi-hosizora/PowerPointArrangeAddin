﻿
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingDialog));
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
            resources.ApplyResources(this.btnOK, "btnOK");
            this.btnOK.Name = "btnOK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // cbxArtWord
            // 
            resources.ApplyResources(this.cbxArtWord, "cbxArtWord");
            this.cbxArtWord.Name = "cbxArtWord";
            this.cbxArtWord.UseVisualStyleBackColor = true;
            // 
            // cbxShapeTextbox
            // 
            resources.ApplyResources(this.cbxShapeTextbox, "cbxShapeTextbox");
            this.cbxShapeTextbox.Name = "cbxShapeTextbox";
            this.cbxShapeTextbox.UseVisualStyleBackColor = true;
            // 
            // cbxShapeSizeAndPosition
            // 
            resources.ApplyResources(this.cbxShapeSizeAndPosition, "cbxShapeSizeAndPosition");
            this.cbxShapeSizeAndPosition.Name = "cbxShapeSizeAndPosition";
            this.cbxShapeSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // cbxReplacePicture
            // 
            resources.ApplyResources(this.cbxReplacePicture, "cbxReplacePicture");
            this.cbxReplacePicture.Name = "cbxReplacePicture";
            this.cbxReplacePicture.UseVisualStyleBackColor = true;
            // 
            // cbxPictureSizeAndPosition
            // 
            resources.ApplyResources(this.cbxPictureSizeAndPosition, "cbxPictureSizeAndPosition");
            this.cbxPictureSizeAndPosition.Name = "cbxPictureSizeAndPosition";
            this.cbxPictureSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // tlpMain
            // 
            resources.ApplyResources(this.tlpMain, "tlpMain");
            this.tlpMain.Controls.Add(this.grpGroupVisibility, 0, 0);
            this.tlpMain.Controls.Add(this.grpOtherSetting, 0, 1);
            this.tlpMain.Controls.Add(this.tbxDescription, 0, 2);
            this.tlpMain.Controls.Add(this.tlpButton, 0, 3);
            this.tlpMain.Name = "tlpMain";
            // 
            // grpGroupVisibility
            // 
            resources.ApplyResources(this.grpGroupVisibility, "grpGroupVisibility");
            this.grpGroupVisibility.Controls.Add(this.tlpGroupVisibility);
            this.grpGroupVisibility.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.grpGroupVisibility.Name = "grpGroupVisibility";
            this.grpGroupVisibility.TabStop = false;
            // 
            // tlpGroupVisibility
            // 
            resources.ApplyResources(this.tlpGroupVisibility, "tlpGroupVisibility");
            this.tlpGroupVisibility.BackColor = System.Drawing.SystemColors.Control;
            this.tlpGroupVisibility.Controls.Add(this.cbxArtWord, 0, 0);
            this.tlpGroupVisibility.Controls.Add(this.cbxShapeTextbox, 0, 1);
            this.tlpGroupVisibility.Controls.Add(this.cbxPictureSizeAndPosition, 0, 4);
            this.tlpGroupVisibility.Controls.Add(this.cbxReplacePicture, 0, 3);
            this.tlpGroupVisibility.Controls.Add(this.cbxShapeSizeAndPosition, 0, 2);
            this.tlpGroupVisibility.Name = "tlpGroupVisibility";
            // 
            // grpOtherSetting
            // 
            resources.ApplyResources(this.grpOtherSetting, "grpOtherSetting");
            this.grpOtherSetting.Controls.Add(this.tlpOtherSetting);
            this.grpOtherSetting.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.grpOtherSetting.Name = "grpOtherSetting";
            this.grpOtherSetting.TabStop = false;
            // 
            // tlpOtherSetting
            // 
            resources.ApplyResources(this.tlpOtherSetting, "tlpOtherSetting");
            this.tlpOtherSetting.Controls.Add(this.lblLanguage, 0, 0);
            this.tlpOtherSetting.Controls.Add(this.cboLanguage, 1, 0);
            this.tlpOtherSetting.Name = "tlpOtherSetting";
            // 
            // lblLanguage
            // 
            resources.ApplyResources(this.lblLanguage, "lblLanguage");
            this.lblLanguage.Name = "lblLanguage";
            // 
            // cboLanguage
            // 
            resources.ApplyResources(this.cboLanguage, "cboLanguage");
            this.cboLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboLanguage.FormattingEnabled = true;
            this.cboLanguage.Items.AddRange(new object[] {
            resources.GetString("cboLanguage.Items"),
            resources.GetString("cboLanguage.Items1"),
            resources.GetString("cboLanguage.Items2"),
            resources.GetString("cboLanguage.Items3"),
            resources.GetString("cboLanguage.Items4")});
            this.cboLanguage.Name = "cboLanguage";
            // 
            // tbxDescription
            // 
            resources.ApplyResources(this.tbxDescription, "tbxDescription");
            this.tbxDescription.Name = "tbxDescription";
            this.tbxDescription.ReadOnly = true;
            // 
            // tlpButton
            // 
            resources.ApplyResources(this.tlpButton, "tlpButton");
            this.tlpButton.Controls.Add(this.btnCancel, 1, 0);
            this.tlpButton.Controls.Add(this.btnOK, 0, 0);
            this.tlpButton.Name = "tlpButton";
            // 
            // SettingDialog
            // 
            this.AcceptButton = this.btnOK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.tlpMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
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