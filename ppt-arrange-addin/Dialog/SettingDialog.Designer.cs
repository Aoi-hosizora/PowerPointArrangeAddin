
namespace ppt_arrange_addin.Dialog {
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
            this.grpOtherSetting = new System.Windows.Forms.GroupBox();
            this.tlpOtherSetting = new System.Windows.Forms.TableLayoutPanel();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.cboLanguage = new System.Windows.Forms.ComboBox();
            this.tbxDescription = new System.Windows.Forms.TextBox();
            this.tlpButton = new System.Windows.Forms.TableLayoutPanel();
            this.chkVideoSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.chkAudioSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.chkTableSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.chkChartSizeAndPosition = new System.Windows.Forms.CheckBox();
            this.chkSmartartSizeAndPosition = new System.Windows.Forms.CheckBox();
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
            // chkWordArt
            // 
            resources.ApplyResources(this.chkWordArt, "chkWordArt");
            this.chkWordArt.Name = "chkWordArt";
            this.chkWordArt.UseVisualStyleBackColor = true;
            // 
            // chkShapeTextbox
            // 
            resources.ApplyResources(this.chkShapeTextbox, "chkShapeTextbox");
            this.chkShapeTextbox.Name = "chkShapeTextbox";
            this.chkShapeTextbox.UseVisualStyleBackColor = true;
            // 
            // chkShapeSizeAndPosition
            // 
            resources.ApplyResources(this.chkShapeSizeAndPosition, "chkShapeSizeAndPosition");
            this.chkShapeSizeAndPosition.Name = "chkShapeSizeAndPosition";
            this.chkShapeSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkReplacePicture
            // 
            resources.ApplyResources(this.chkReplacePicture, "chkReplacePicture");
            this.chkReplacePicture.Name = "chkReplacePicture";
            this.chkReplacePicture.UseVisualStyleBackColor = true;
            // 
            // chkPictureSizeAndPosition
            // 
            resources.ApplyResources(this.chkPictureSizeAndPosition, "chkPictureSizeAndPosition");
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
            this.tlpGroupVisibility.Controls.Add(this.chkPictureSizeAndPosition, 0, 5);
            this.tlpGroupVisibility.Controls.Add(this.chkReplacePicture, 0, 4);
            this.tlpGroupVisibility.Controls.Add(this.chkShapeSizeAndPosition, 0, 3);
            this.tlpGroupVisibility.Controls.Add(this.chkShapeTextbox, 0, 2);
            this.tlpGroupVisibility.Controls.Add(this.chkArrangement, 0, 1);
            this.tlpGroupVisibility.Controls.Add(this.chkWordArt, 0, 0);
            this.tlpGroupVisibility.Controls.Add(this.chkVideoSizeAndPosition, 0, 6);
            this.tlpGroupVisibility.Controls.Add(this.chkAudioSizeAndPosition, 0, 7);
            this.tlpGroupVisibility.Controls.Add(this.chkTableSizeAndPosition, 0, 8);
            this.tlpGroupVisibility.Controls.Add(this.chkChartSizeAndPosition, 0, 9);
            this.tlpGroupVisibility.Controls.Add(this.chkSmartartSizeAndPosition, 0, 10);
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
            // chkVideoSizeAndPosition
            // 
            resources.ApplyResources(this.chkVideoSizeAndPosition, "chkVideoSizeAndPosition");
            this.chkVideoSizeAndPosition.Name = "chkVideoSizeAndPosition";
            this.chkVideoSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkAudioSizeAndPosition
            // 
            resources.ApplyResources(this.chkAudioSizeAndPosition, "chkAudioSizeAndPosition");
            this.chkAudioSizeAndPosition.Name = "chkAudioSizeAndPosition";
            this.chkAudioSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkTableSizeAndPosition
            // 
            resources.ApplyResources(this.chkTableSizeAndPosition, "chkTableSizeAndPosition");
            this.chkTableSizeAndPosition.Name = "chkTableSizeAndPosition";
            this.chkTableSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkChartSizeAndPosition
            // 
            resources.ApplyResources(this.chkChartSizeAndPosition, "chkChartSizeAndPosition");
            this.chkChartSizeAndPosition.Name = "chkChartSizeAndPosition";
            this.chkChartSizeAndPosition.UseVisualStyleBackColor = true;
            // 
            // chkSmartartSizeAndPosition
            // 
            resources.ApplyResources(this.chkSmartartSizeAndPosition, "chkSmartartSizeAndPosition");
            this.chkSmartartSizeAndPosition.Name = "chkSmartartSizeAndPosition";
            this.chkSmartartSizeAndPosition.UseVisualStyleBackColor = true;
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
    }
}