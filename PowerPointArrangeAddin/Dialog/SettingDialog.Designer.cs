
namespace PowerPointArrangeAddin.Dialog {
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
            this.chkWordArt = new System.Windows.Forms.CheckBox();
            this.chkShapeTextbox = new System.Windows.Forms.CheckBox();
            this.chkShapeSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.chkReplacePicture = new System.Windows.Forms.CheckBox();
            this.chkPictureSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.tlpMain = new System.Windows.Forms.TableLayoutPanel();
            this.grpGroupVisibility = new System.Windows.Forms.GroupBox();
            this.tlpGroupVisibility = new System.Windows.Forms.TableLayoutPanel();
            this.chkArrangement = new System.Windows.Forms.CheckBox();
            this.chkSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.tlpSizeAndPosition = new System.Windows.Forms.TableLayoutPanel();
            this.chkChartSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.chkSmartartSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.chkTableSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.chkAudioSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.chkVideoSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.grpOtherSetting = new System.Windows.Forms.GroupBox();
            this.tlpOtherSetting = new System.Windows.Forms.TableLayoutPanel();
            this.cboIconStyle = new System.Windows.Forms.ComboBox();
            this.txtIconStyle = new System.Windows.Forms.Label();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.cboLanguage = new System.Windows.Forms.ComboBox();
            this.chkCheckUpdateWhenStartUp = new System.Windows.Forms.CheckBox();
            this.chkHideMarginSettingForTextbox = new System.Windows.Forms.CheckBox();
            this.chkLessButtonsForArrangement = new System.Windows.Forms.CheckBox();
            this.chkAllowDoublePressExtendButton = new System.Windows.Forms.CheckBox();
            this.tbxDescription = new System.Windows.Forms.TextBox();
            this.tlpButton = new System.Windows.Forms.TableLayoutPanel();
            this.tlpMain.SuspendLayout();
            this.grpGroupVisibility.SuspendLayout();
            this.tlpGroupVisibility.SuspendLayout();
            this.tlpSizeAndPosition.SuspendLayout();
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
            // chkWordArt
            // 
            resources.ApplyResources(this.chkWordArt, "chkWordArt");
            this.chkWordArt.Checked = true;
            this.chkWordArt.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkWordArt.Name = "chkWordArt";
            this.chkWordArt.UseVisualStyleBackColor = true;
            // 
            // chkShapeTextbox
            // 
            resources.ApplyResources(this.chkShapeTextbox, "chkShapeTextbox");
            this.chkShapeTextbox.Checked = true;
            this.chkShapeTextbox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkShapeTextbox.Name = "chkShapeTextbox";
            this.chkShapeTextbox.UseVisualStyleBackColor = true;
            // 
            // chkShapeSizeAndPosition
            // 
            resources.ApplyResources(this.chkShapeSizeAndPosition, "chkShapeSizeAndPosition");
            this.chkShapeSizeAndPosition.Checked = true;
            this.chkShapeSizeAndPosition.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkShapeSizeAndPosition.Name = "chkShapeSizeAndPosition";
            this.chkShapeSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkReplacePicture
            // 
            resources.ApplyResources(this.chkReplacePicture, "chkReplacePicture");
            this.chkReplacePicture.Checked = true;
            this.chkReplacePicture.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkReplacePicture.Name = "chkReplacePicture";
            this.chkReplacePicture.UseVisualStyleBackColor = true;
            // 
            // chkPictureSizeAndPosition
            // 
            resources.ApplyResources(this.chkPictureSizeAndPosition, "chkPictureSizeAndPosition");
            this.chkPictureSizeAndPosition.Checked = true;
            this.chkPictureSizeAndPosition.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPictureSizeAndPosition.Name = "chkPictureSizeAndPosition";
            this.chkPictureSizeAndPosition.UseVisualStyleBackColor = true;
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
            this.tlpGroupVisibility.Controls.Add(this.chkShapeTextbox, 0, 2);
            this.tlpGroupVisibility.Controls.Add(this.chkArrangement, 0, 1);
            this.tlpGroupVisibility.Controls.Add(this.chkWordArt, 0, 0);
            this.tlpGroupVisibility.Controls.Add(this.chkReplacePicture, 0, 3);
            this.tlpGroupVisibility.Controls.Add(this.chkSizeAndPosition, 0, 11);
            this.tlpGroupVisibility.Controls.Add(this.tlpSizeAndPosition, 0, 12);
            this.tlpGroupVisibility.Name = "tlpGroupVisibility";
            // 
            // chkArrangement
            // 
            resources.ApplyResources(this.chkArrangement, "chkArrangement");
            this.chkArrangement.Checked = true;
            this.chkArrangement.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkArrangement.Name = "chkArrangement";
            this.chkArrangement.UseVisualStyleBackColor = true;
            // 
            // chkSizeAndPosition
            // 
            resources.ApplyResources(this.chkSizeAndPosition, "chkSizeAndPosition");
            this.chkSizeAndPosition.Checked = true;
            this.chkSizeAndPosition.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSizeAndPosition.Name = "chkSizeAndPosition";
            this.chkSizeAndPosition.UseVisualStyleBackColor = true;
            this.chkSizeAndPosition.CheckedChanged += new System.EventHandler(this.ChkSizeAndPosition_CheckedChanged);
            // 
            // tlpSizeAndPosition
            // 
            resources.ApplyResources(this.tlpSizeAndPosition, "tlpSizeAndPosition");
            this.tlpSizeAndPosition.Controls.Add(this.chkShapeSizeAndPosition, 0, 0);
            this.tlpSizeAndPosition.Controls.Add(this.chkPictureSizeAndPosition, 1, 0);
            this.tlpSizeAndPosition.Controls.Add(this.chkChartSizeAndPosition, 1, 2);
            this.tlpSizeAndPosition.Controls.Add(this.chkSmartartSizeAndPosition, 0, 3);
            this.tlpSizeAndPosition.Controls.Add(this.chkTableSizeAndPosition, 0, 2);
            this.tlpSizeAndPosition.Controls.Add(this.chkAudioSizeAndPosition, 1, 1);
            this.tlpSizeAndPosition.Controls.Add(this.chkVideoSizeAndPosition, 0, 1);
            this.tlpSizeAndPosition.Name = "tlpSizeAndPosition";
            // 
            // chkChartSizeAndPosition
            // 
            resources.ApplyResources(this.chkChartSizeAndPosition, "chkChartSizeAndPosition");
            this.chkChartSizeAndPosition.Checked = true;
            this.chkChartSizeAndPosition.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkChartSizeAndPosition.Name = "chkChartSizeAndPosition";
            this.chkChartSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkSmartartSizeAndPosition
            // 
            resources.ApplyResources(this.chkSmartartSizeAndPosition, "chkSmartartSizeAndPosition");
            this.chkSmartartSizeAndPosition.Checked = true;
            this.chkSmartartSizeAndPosition.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSmartartSizeAndPosition.Name = "chkSmartartSizeAndPosition";
            this.chkSmartartSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkTableSizeAndPosition
            // 
            resources.ApplyResources(this.chkTableSizeAndPosition, "chkTableSizeAndPosition");
            this.chkTableSizeAndPosition.Checked = true;
            this.chkTableSizeAndPosition.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkTableSizeAndPosition.Name = "chkTableSizeAndPosition";
            this.chkTableSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkAudioSizeAndPosition
            // 
            resources.ApplyResources(this.chkAudioSizeAndPosition, "chkAudioSizeAndPosition");
            this.chkAudioSizeAndPosition.Checked = true;
            this.chkAudioSizeAndPosition.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAudioSizeAndPosition.Name = "chkAudioSizeAndPosition";
            this.chkAudioSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkVideoSizeAndPosition
            // 
            resources.ApplyResources(this.chkVideoSizeAndPosition, "chkVideoSizeAndPosition");
            this.chkVideoSizeAndPosition.Checked = true;
            this.chkVideoSizeAndPosition.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkVideoSizeAndPosition.Name = "chkVideoSizeAndPosition";
            this.chkVideoSizeAndPosition.UseVisualStyleBackColor = true;
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
            this.tlpOtherSetting.Controls.Add(this.cboIconStyle, 1, 5);
            this.tlpOtherSetting.Controls.Add(this.txtIconStyle, 0, 5);
            this.tlpOtherSetting.Controls.Add(this.lblLanguage, 0, 0);
            this.tlpOtherSetting.Controls.Add(this.cboLanguage, 1, 0);
            this.tlpOtherSetting.Controls.Add(this.chkCheckUpdateWhenStartUp, 0, 1);
            this.tlpOtherSetting.Controls.Add(this.chkHideMarginSettingForTextbox, 0, 3);
            this.tlpOtherSetting.Controls.Add(this.chkLessButtonsForArrangement, 0, 2);
            this.tlpOtherSetting.Controls.Add(this.chkAllowDoublePressExtendButton, 0, 4);
            this.tlpOtherSetting.Name = "tlpOtherSetting";
            // 
            // cboIconStyle
            // 
            resources.ApplyResources(this.cboIconStyle, "cboIconStyle");
            this.cboIconStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboIconStyle.FormattingEnabled = true;
            this.cboIconStyle.Items.AddRange(new object[] {
            resources.GetString("cboIconStyle.Items"),
            resources.GetString("cboIconStyle.Items1")});
            this.cboIconStyle.Name = "cboIconStyle";
            // 
            // txtIconStyle
            // 
            resources.ApplyResources(this.txtIconStyle, "txtIconStyle");
            this.txtIconStyle.Name = "txtIconStyle";
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
            // chkCheckUpdateWhenStartUp
            // 
            resources.ApplyResources(this.chkCheckUpdateWhenStartUp, "chkCheckUpdateWhenStartUp");
            this.tlpOtherSetting.SetColumnSpan(this.chkCheckUpdateWhenStartUp, 2);
            this.chkCheckUpdateWhenStartUp.Name = "chkCheckUpdateWhenStartUp";
            this.chkCheckUpdateWhenStartUp.UseVisualStyleBackColor = true;
            // 
            // chkHideMarginSettingForTextbox
            // 
            resources.ApplyResources(this.chkHideMarginSettingForTextbox, "chkHideMarginSettingForTextbox");
            this.tlpOtherSetting.SetColumnSpan(this.chkHideMarginSettingForTextbox, 2);
            this.chkHideMarginSettingForTextbox.Name = "chkHideMarginSettingForTextbox";
            this.chkHideMarginSettingForTextbox.UseVisualStyleBackColor = true;
            // 
            // chkLessButtonsForArrangement
            // 
            resources.ApplyResources(this.chkLessButtonsForArrangement, "chkLessButtonsForArrangement");
            this.tlpOtherSetting.SetColumnSpan(this.chkLessButtonsForArrangement, 2);
            this.chkLessButtonsForArrangement.Name = "chkLessButtonsForArrangement";
            this.chkLessButtonsForArrangement.UseVisualStyleBackColor = true;
            // 
            // chkAllowDoublePressExtendButton
            // 
            resources.ApplyResources(this.chkAllowDoublePressExtendButton, "chkAllowDoublePressExtendButton");
            this.tlpOtherSetting.SetColumnSpan(this.chkAllowDoublePressExtendButton, 2);
            this.chkAllowDoublePressExtendButton.Name = "chkAllowDoublePressExtendButton";
            this.chkAllowDoublePressExtendButton.UseVisualStyleBackColor = true;
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
            this.tlpButton.Controls.Add(this.btnOK, 0, 0);
            this.tlpButton.Controls.Add(this.btnCancel, 1, 0);
            this.tlpButton.Name = "tlpButton";
            // 
            // SettingDialog
            // 
            this.AcceptButton = this.btnOK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.tlpMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
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
            this.tlpSizeAndPosition.ResumeLayout(false);
            this.tlpSizeAndPosition.PerformLayout();
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
        private System.Windows.Forms.CheckBox chkWordArt;
        private System.Windows.Forms.CheckBox chkShapeTextbox;
        private System.Windows.Forms.CheckBox chkShapeSizeAndPosition;
        private System.Windows.Forms.CheckBox chkReplacePicture;
        private System.Windows.Forms.CheckBox chkPictureSizeAndPosition;
        private System.Windows.Forms.TableLayoutPanel tlpMain;
        private System.Windows.Forms.GroupBox grpGroupVisibility;
        private System.Windows.Forms.GroupBox grpOtherSetting;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.ComboBox cboLanguage;
        private System.Windows.Forms.TableLayoutPanel tlpGroupVisibility;
        private System.Windows.Forms.TableLayoutPanel tlpOtherSetting;
        private System.Windows.Forms.TextBox tbxDescription;
        private System.Windows.Forms.TableLayoutPanel tlpButton;
        private System.Windows.Forms.CheckBox chkArrangement;
        private System.Windows.Forms.CheckBox chkVideoSizeAndPosition;
        private System.Windows.Forms.CheckBox chkAudioSizeAndPosition;
        private System.Windows.Forms.CheckBox chkTableSizeAndPosition;
        private System.Windows.Forms.CheckBox chkChartSizeAndPosition;
        private System.Windows.Forms.CheckBox chkSmartartSizeAndPosition;
        private System.Windows.Forms.CheckBox chkLessButtonsForArrangement;
        private System.Windows.Forms.CheckBox chkHideMarginSettingForTextbox;
        private System.Windows.Forms.CheckBox chkSizeAndPosition;
        private System.Windows.Forms.TableLayoutPanel tlpSizeAndPosition;
        private System.Windows.Forms.CheckBox chkCheckUpdateWhenStartUp;
        private System.Windows.Forms.Label txtIconStyle;
        private System.Windows.Forms.ComboBox cboIconStyle;
        private System.Windows.Forms.CheckBox chkAllowDoublePressExtendButton;
    }
}