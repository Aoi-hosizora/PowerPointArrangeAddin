
namespace ppt_arrange_addin {
    partial class ArrangeRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ArrangeRibbon()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.tabHome = this.Factory.CreateRibbonTab();
            this.grpArrange = this.Factory.CreateRibbonGroup();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.btnAlignLeft = this.Factory.CreateRibbonButton();
            this.btnAlignCenter = this.Factory.CreateRibbonButton();
            this.btnAlignRight = this.Factory.CreateRibbonButton();
            this.buttonGroup2 = this.Factory.CreateRibbonButtonGroup();
            this.btnAlignTop = this.Factory.CreateRibbonButton();
            this.btnAlignMiddle = this.Factory.CreateRibbonButton();
            this.btnAlignBottom = this.Factory.CreateRibbonButton();
            this.buttonGroup3 = this.Factory.CreateRibbonButtonGroup();
            this.btnDistributeHorizontal = this.Factory.CreateRibbonButton();
            this.btnDistributeVertical = this.Factory.CreateRibbonButton();
            this.mnuRotate = this.Factory.CreateRibbonMenu();
            this.btnRotateRight90 = this.Factory.CreateRibbonButton();
            this.btnRotateLeft90 = this.Factory.CreateRibbonButton();
            this.btnFlipVertical = this.Factory.CreateRibbonButton();
            this.btnFlipHorizontal = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.buttonGroup7 = this.Factory.CreateRibbonButtonGroup();
            this.btnScaleSameWidth = this.Factory.CreateRibbonButton();
            this.btnScaleSameHeight = this.Factory.CreateRibbonButton();
            this.btnScaleSameSize = this.Factory.CreateRibbonButton();
            this.buttonGroup8 = this.Factory.CreateRibbonButtonGroup();
            this.btnExtendLeft = this.Factory.CreateRibbonButton();
            this.btnExtendRight = this.Factory.CreateRibbonButton();
            this.buttonGroup9 = this.Factory.CreateRibbonButtonGroup();
            this.btnExtendTop = this.Factory.CreateRibbonButton();
            this.btnExtendBottom = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.buttonGroup4 = this.Factory.CreateRibbonButtonGroup();
            this.btnMoveForward = this.Factory.CreateRibbonButton();
            this.btnMoveFront = this.Factory.CreateRibbonButton();
            this.buttonGroup5 = this.Factory.CreateRibbonButtonGroup();
            this.btnMoveBackward = this.Factory.CreateRibbonButton();
            this.btnMoveBack = this.Factory.CreateRibbonButton();
            this.buttonGroup6 = this.Factory.CreateRibbonButtonGroup();
            this.btnGroup = this.Factory.CreateRibbonButton();
            this.btnUngroup = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.buttonGroup10 = this.Factory.CreateRibbonButtonGroup();
            this.btnSnapLeft = this.Factory.CreateRibbonButton();
            this.btnSnapRight = this.Factory.CreateRibbonButton();
            this.buttonGroup11 = this.Factory.CreateRibbonButtonGroup();
            this.btnSnapTop = this.Factory.CreateRibbonButton();
            this.btnSnapBottom = this.Factory.CreateRibbonButton();
            this.tabHome.SuspendLayout();
            this.grpArrange.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.buttonGroup2.SuspendLayout();
            this.buttonGroup3.SuspendLayout();
            this.buttonGroup7.SuspendLayout();
            this.buttonGroup8.SuspendLayout();
            this.buttonGroup9.SuspendLayout();
            this.buttonGroup4.SuspendLayout();
            this.buttonGroup5.SuspendLayout();
            this.buttonGroup6.SuspendLayout();
            this.buttonGroup10.SuspendLayout();
            this.buttonGroup11.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabHome
            // 
            this.tabHome.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabHome.ControlId.OfficeId = "TabHome";
            this.tabHome.Groups.Add(this.grpArrange);
            this.tabHome.Label = "TabHome";
            this.tabHome.Name = "tabHome";
            // 
            // grpArrange
            // 
            this.grpArrange.Items.Add(this.buttonGroup1);
            this.grpArrange.Items.Add(this.buttonGroup2);
            this.grpArrange.Items.Add(this.buttonGroup3);
            this.grpArrange.Items.Add(this.separator1);
            this.grpArrange.Items.Add(this.buttonGroup7);
            this.grpArrange.Items.Add(this.buttonGroup8);
            this.grpArrange.Items.Add(this.buttonGroup9);
            this.grpArrange.Items.Add(this.separator2);
            this.grpArrange.Items.Add(this.buttonGroup4);
            this.grpArrange.Items.Add(this.buttonGroup5);
            this.grpArrange.Items.Add(this.buttonGroup6);
            this.grpArrange.Items.Add(this.separator4);
            this.grpArrange.Items.Add(this.buttonGroup10);
            this.grpArrange.Items.Add(this.buttonGroup11);
            this.grpArrange.Label = "Arrange";
            this.grpArrange.Name = "grpArrange";
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.btnAlignLeft);
            this.buttonGroup1.Items.Add(this.btnAlignCenter);
            this.buttonGroup1.Items.Add(this.btnAlignRight);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // btnAlignLeft
            // 
            this.btnAlignLeft.Image = global::ppt_arrange_addin.Properties.Resources.ObjectsAlignLeft;
            this.btnAlignLeft.Label = "Align left";
            this.btnAlignLeft.Name = "btnAlignLeft";
            this.btnAlignLeft.ScreenTip = "Align left";
            this.btnAlignLeft.ShowImage = true;
            this.btnAlignLeft.ShowLabel = false;
            // 
            // btnAlignCenter
            // 
            this.btnAlignCenter.Image = global::ppt_arrange_addin.Properties.Resources.ObjectsAlignCenterHorizontal;
            this.btnAlignCenter.Label = "Align center";
            this.btnAlignCenter.Name = "btnAlignCenter";
            this.btnAlignCenter.ScreenTip = "Align center";
            this.btnAlignCenter.ShowImage = true;
            this.btnAlignCenter.ShowLabel = false;
            // 
            // btnAlignRight
            // 
            this.btnAlignRight.Image = global::ppt_arrange_addin.Properties.Resources.ObjectsAlignRight;
            this.btnAlignRight.Label = "Align right";
            this.btnAlignRight.Name = "btnAlignRight";
            this.btnAlignRight.ScreenTip = "Align right";
            this.btnAlignRight.ShowImage = true;
            this.btnAlignRight.ShowLabel = false;
            // 
            // buttonGroup2
            // 
            this.buttonGroup2.Items.Add(this.btnAlignTop);
            this.buttonGroup2.Items.Add(this.btnAlignMiddle);
            this.buttonGroup2.Items.Add(this.btnAlignBottom);
            this.buttonGroup2.Name = "buttonGroup2";
            // 
            // btnAlignTop
            // 
            this.btnAlignTop.Image = global::ppt_arrange_addin.Properties.Resources.ObjectsAlignTop;
            this.btnAlignTop.Label = "Align top";
            this.btnAlignTop.Name = "btnAlignTop";
            this.btnAlignTop.ScreenTip = "Align top";
            this.btnAlignTop.ShowImage = true;
            this.btnAlignTop.ShowLabel = false;
            // 
            // btnAlignMiddle
            // 
            this.btnAlignMiddle.Image = global::ppt_arrange_addin.Properties.Resources.ObjectsAlignMiddleVertical;
            this.btnAlignMiddle.Label = "Align middle";
            this.btnAlignMiddle.Name = "btnAlignMiddle";
            this.btnAlignMiddle.ScreenTip = "Align middle";
            this.btnAlignMiddle.ShowImage = true;
            this.btnAlignMiddle.ShowLabel = false;
            // 
            // btnAlignBottom
            // 
            this.btnAlignBottom.Image = global::ppt_arrange_addin.Properties.Resources.ObjectsAlignBottom;
            this.btnAlignBottom.Label = "Align bottom";
            this.btnAlignBottom.Name = "btnAlignBottom";
            this.btnAlignBottom.ScreenTip = "Align bottom";
            this.btnAlignBottom.ShowImage = true;
            this.btnAlignBottom.ShowLabel = false;
            // 
            // buttonGroup3
            // 
            this.buttonGroup3.Items.Add(this.btnDistributeHorizontal);
            this.buttonGroup3.Items.Add(this.btnDistributeVertical);
            this.buttonGroup3.Items.Add(this.mnuRotate);
            this.buttonGroup3.Name = "buttonGroup3";
            // 
            // btnDistributeHorizontal
            // 
            this.btnDistributeHorizontal.Image = global::ppt_arrange_addin.Properties.Resources.AlignDistributeHorizontally;
            this.btnDistributeHorizontal.Label = "Distribute horizontally";
            this.btnDistributeHorizontal.Name = "btnDistributeHorizontal";
            this.btnDistributeHorizontal.ScreenTip = "Distribute horizontally";
            this.btnDistributeHorizontal.ShowImage = true;
            this.btnDistributeHorizontal.ShowLabel = false;
            // 
            // btnDistributeVertical
            // 
            this.btnDistributeVertical.Image = global::ppt_arrange_addin.Properties.Resources.AlignDistributeVertically;
            this.btnDistributeVertical.Label = "Distribute vertically";
            this.btnDistributeVertical.Name = "btnDistributeVertical";
            this.btnDistributeVertical.ScreenTip = "Distribute vertically";
            this.btnDistributeVertical.ShowImage = true;
            this.btnDistributeVertical.ShowLabel = false;
            // 
            // mnuRotate
            // 
            this.mnuRotate.Image = global::ppt_arrange_addin.Properties.Resources.ObjectRotateRight90;
            this.mnuRotate.Items.Add(this.btnRotateRight90);
            this.mnuRotate.Items.Add(this.btnRotateLeft90);
            this.mnuRotate.Items.Add(this.btnFlipVertical);
            this.mnuRotate.Items.Add(this.btnFlipHorizontal);
            this.mnuRotate.Label = "Rotate Shape";
            this.mnuRotate.Name = "mnuRotate";
            this.mnuRotate.ScreenTip = "Rotate Shape";
            this.mnuRotate.ShowImage = true;
            this.mnuRotate.ShowLabel = false;
            // 
            // btnRotateRight90
            // 
            this.btnRotateRight90.Image = global::ppt_arrange_addin.Properties.Resources.ObjectRotateRight90;
            this.btnRotateRight90.Label = "Rotate right with 90 degrees";
            this.btnRotateRight90.Name = "btnRotateRight90";
            this.btnRotateRight90.ShowImage = true;
            // 
            // btnRotateLeft90
            // 
            this.btnRotateLeft90.Image = global::ppt_arrange_addin.Properties.Resources.ObjectRotateLeft90;
            this.btnRotateLeft90.Label = "Rorate left with 90 degrees";
            this.btnRotateLeft90.Name = "btnRotateLeft90";
            this.btnRotateLeft90.ShowImage = true;
            // 
            // btnFlipVertical
            // 
            this.btnFlipVertical.Image = global::ppt_arrange_addin.Properties.Resources.ObjectFlipVertical;
            this.btnFlipVertical.Label = "Flip vertically";
            this.btnFlipVertical.Name = "btnFlipVertical";
            this.btnFlipVertical.ShowImage = true;
            // 
            // btnFlipHorizontal
            // 
            this.btnFlipHorizontal.Image = global::ppt_arrange_addin.Properties.Resources.ObjectFlipHorizontal;
            this.btnFlipHorizontal.Label = "Flip horizontally";
            this.btnFlipHorizontal.Name = "btnFlipHorizontal";
            this.btnFlipHorizontal.ShowImage = true;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // buttonGroup7
            // 
            this.buttonGroup7.Items.Add(this.btnScaleSameWidth);
            this.buttonGroup7.Items.Add(this.btnScaleSameHeight);
            this.buttonGroup7.Items.Add(this.btnScaleSameSize);
            this.buttonGroup7.Name = "buttonGroup7";
            // 
            // btnScaleSameWidth
            // 
            this.btnScaleSameWidth.Label = "Scale to same width";
            this.btnScaleSameWidth.Name = "btnScaleSameWidth";
            this.btnScaleSameWidth.ScreenTip = "Scale to same width";
            // 
            // btnScaleSameHeight
            // 
            this.btnScaleSameHeight.Label = "Scale to same height";
            this.btnScaleSameHeight.Name = "btnScaleSameHeight";
            this.btnScaleSameHeight.ScreenTip = "Scale to same height";
            // 
            // btnScaleSameSize
            // 
            this.btnScaleSameSize.Label = "Scale to same size";
            this.btnScaleSameSize.Name = "btnScaleSameSize";
            this.btnScaleSameSize.ScreenTip = "Scale to same size";
            // 
            // buttonGroup8
            // 
            this.buttonGroup8.Items.Add(this.btnExtendLeft);
            this.buttonGroup8.Items.Add(this.btnExtendRight);
            this.buttonGroup8.Name = "buttonGroup8";
            // 
            // btnExtendLeft
            // 
            this.btnExtendLeft.Label = "Extend to same left";
            this.btnExtendLeft.Name = "btnExtendLeft";
            this.btnExtendLeft.ScreenTip = "Extend to same left";
            // 
            // btnExtendRight
            // 
            this.btnExtendRight.Label = "Extend to same right";
            this.btnExtendRight.Name = "btnExtendRight";
            this.btnExtendRight.ScreenTip = "Extend to same right";
            // 
            // buttonGroup9
            // 
            this.buttonGroup9.Items.Add(this.btnExtendTop);
            this.buttonGroup9.Items.Add(this.btnExtendBottom);
            this.buttonGroup9.Name = "buttonGroup9";
            // 
            // btnExtendTop
            // 
            this.btnExtendTop.Label = "Extend to same top";
            this.btnExtendTop.Name = "btnExtendTop";
            this.btnExtendTop.ScreenTip = "Extend to same top";
            // 
            // btnExtendBottom
            // 
            this.btnExtendBottom.Label = "Extend to same bottom";
            this.btnExtendBottom.Name = "btnExtendBottom";
            this.btnExtendBottom.ScreenTip = "Extend to same bottom";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // buttonGroup4
            // 
            this.buttonGroup4.Items.Add(this.btnMoveForward);
            this.buttonGroup4.Items.Add(this.btnMoveFront);
            this.buttonGroup4.Name = "buttonGroup4";
            // 
            // btnMoveForward
            // 
            this.btnMoveForward.Image = global::ppt_arrange_addin.Properties.Resources.ObjectBringForward;
            this.btnMoveForward.Label = "Move forward";
            this.btnMoveForward.Name = "btnMoveForward";
            this.btnMoveForward.ScreenTip = "Move forward";
            this.btnMoveForward.ShowImage = true;
            this.btnMoveForward.ShowLabel = false;
            // 
            // btnMoveFront
            // 
            this.btnMoveFront.Image = global::ppt_arrange_addin.Properties.Resources.ObjectBringToFront;
            this.btnMoveFront.Label = "Move to front";
            this.btnMoveFront.Name = "btnMoveFront";
            this.btnMoveFront.ScreenTip = "Move to front";
            this.btnMoveFront.ShowImage = true;
            this.btnMoveFront.ShowLabel = false;
            // 
            // buttonGroup5
            // 
            this.buttonGroup5.Items.Add(this.btnMoveBackward);
            this.buttonGroup5.Items.Add(this.btnMoveBack);
            this.buttonGroup5.Name = "buttonGroup5";
            // 
            // btnMoveBackward
            // 
            this.btnMoveBackward.Image = global::ppt_arrange_addin.Properties.Resources.ObjectSendBackward;
            this.btnMoveBackward.Label = "Move backward";
            this.btnMoveBackward.Name = "btnMoveBackward";
            this.btnMoveBackward.ScreenTip = "Move backward";
            this.btnMoveBackward.ShowImage = true;
            this.btnMoveBackward.ShowLabel = false;
            // 
            // btnMoveBack
            // 
            this.btnMoveBack.Image = global::ppt_arrange_addin.Properties.Resources.ObjectSendToBack;
            this.btnMoveBack.Label = "Move to back";
            this.btnMoveBack.Name = "btnMoveBack";
            this.btnMoveBack.ScreenTip = "Move to back";
            this.btnMoveBack.ShowImage = true;
            this.btnMoveBack.ShowLabel = false;
            // 
            // buttonGroup6
            // 
            this.buttonGroup6.Items.Add(this.btnGroup);
            this.buttonGroup6.Items.Add(this.btnUngroup);
            this.buttonGroup6.Name = "buttonGroup6";
            // 
            // btnGroup
            // 
            this.btnGroup.Image = global::ppt_arrange_addin.Properties.Resources.ObjectsGroup;
            this.btnGroup.Label = "Group shapes";
            this.btnGroup.Name = "btnGroup";
            this.btnGroup.ScreenTip = "Group shapes";
            this.btnGroup.ShowImage = true;
            this.btnGroup.ShowLabel = false;
            // 
            // btnUngroup
            // 
            this.btnUngroup.Image = global::ppt_arrange_addin.Properties.Resources.ObjectsUngroup;
            this.btnUngroup.Label = "Ungroup shapes";
            this.btnUngroup.Name = "btnUngroup";
            this.btnUngroup.ScreenTip = "Ungroup shapes";
            this.btnUngroup.ShowImage = true;
            this.btnUngroup.ShowLabel = false;
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // buttonGroup10
            // 
            this.buttonGroup10.Items.Add(this.btnSnapLeft);
            this.buttonGroup10.Items.Add(this.btnSnapRight);
            this.buttonGroup10.Name = "buttonGroup10";
            // 
            // btnSnapLeft
            // 
            this.btnSnapLeft.Label = "Snap to Left";
            this.btnSnapLeft.Name = "btnSnapLeft";
            this.btnSnapLeft.ScreenTip = "Snap to Left";
            // 
            // btnSnapRight
            // 
            this.btnSnapRight.Label = "Snap to right";
            this.btnSnapRight.Name = "btnSnapRight";
            this.btnSnapRight.ScreenTip = "Snap to right";
            // 
            // buttonGroup11
            // 
            this.buttonGroup11.Items.Add(this.btnSnapTop);
            this.buttonGroup11.Items.Add(this.btnSnapBottom);
            this.buttonGroup11.Name = "buttonGroup11";
            // 
            // btnSnapTop
            // 
            this.btnSnapTop.Label = "Snap to top";
            this.btnSnapTop.Name = "btnSnapTop";
            this.btnSnapTop.ScreenTip = "Snap to top";
            // 
            // btnSnapBottom
            // 
            this.btnSnapBottom.Label = "Snap to bottom";
            this.btnSnapBottom.Name = "btnSnapBottom";
            this.btnSnapBottom.ScreenTip = "Snap to bottom";
            // 
            // ArrangeRibbon
            // 
            this.Name = "ArrangeRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabHome);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ArrangeRibbon_Load);
            this.tabHome.ResumeLayout(false);
            this.tabHome.PerformLayout();
            this.grpArrange.ResumeLayout(false);
            this.grpArrange.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.buttonGroup2.ResumeLayout(false);
            this.buttonGroup2.PerformLayout();
            this.buttonGroup3.ResumeLayout(false);
            this.buttonGroup3.PerformLayout();
            this.buttonGroup7.ResumeLayout(false);
            this.buttonGroup7.PerformLayout();
            this.buttonGroup8.ResumeLayout(false);
            this.buttonGroup8.PerformLayout();
            this.buttonGroup9.ResumeLayout(false);
            this.buttonGroup9.PerformLayout();
            this.buttonGroup4.ResumeLayout(false);
            this.buttonGroup4.PerformLayout();
            this.buttonGroup5.ResumeLayout(false);
            this.buttonGroup5.PerformLayout();
            this.buttonGroup6.ResumeLayout(false);
            this.buttonGroup6.PerformLayout();
            this.buttonGroup10.ResumeLayout(false);
            this.buttonGroup10.PerformLayout();
            this.buttonGroup11.ResumeLayout(false);
            this.buttonGroup11.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabHome;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpArrange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtendLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignMiddle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignCenter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDistributeHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDistributeVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtendRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtendTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtendBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnScaleSameWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnScaleSameHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnScaleSameSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUngroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveForward;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveBackward;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveFront;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveBack;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSnapLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSnapRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSnapTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSnapBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup6;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnuRotate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRotateRight90;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRotateLeft90;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFlipVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFlipHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup11;
    }

    partial class ThisRibbonCollection {
        internal ArrangeRibbon ArrangeRibbon {
            get { return this.GetRibbon<ArrangeRibbon>(); }
        }
    }
}
