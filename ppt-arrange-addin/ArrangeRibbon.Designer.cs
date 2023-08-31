
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
            this.grpAlignLR = this.Factory.CreateRibbonButtonGroup();
            this.grpAlignTB = this.Factory.CreateRibbonButtonGroup();
            this.grpDistributeAndRotate = this.Factory.CreateRibbonButtonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.grpScaleWH = this.Factory.CreateRibbonButtonGroup();
            this.grpExtendLR = this.Factory.CreateRibbonButtonGroup();
            this.grpExtendTB = this.Factory.CreateRibbonButtonGroup();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.grpSnapLR = this.Factory.CreateRibbonButtonGroup();
            this.grpSnapTB = this.Factory.CreateRibbonButtonGroup();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.grpMoveF = this.Factory.CreateRibbonButtonGroup();
            this.grpMoveB = this.Factory.CreateRibbonButtonGroup();
            this.grpGroup = this.Factory.CreateRibbonButtonGroup();
            this.btnAlignLeft = this.Factory.CreateRibbonButton();
            this.btnAlignCenter = this.Factory.CreateRibbonButton();
            this.btnAlignRight = this.Factory.CreateRibbonButton();
            this.btnAlignTop = this.Factory.CreateRibbonButton();
            this.btnAlignMiddle = this.Factory.CreateRibbonButton();
            this.btnAlignBottom = this.Factory.CreateRibbonButton();
            this.btnDistributeHorizontal = this.Factory.CreateRibbonButton();
            this.btnDistributeVertical = this.Factory.CreateRibbonButton();
            this.mnuRotate = this.Factory.CreateRibbonMenu();
            this.btnRotateRight90 = this.Factory.CreateRibbonButton();
            this.btnRotateLeft90 = this.Factory.CreateRibbonButton();
            this.btnFlipVertical = this.Factory.CreateRibbonButton();
            this.btnFlipHorizontal = this.Factory.CreateRibbonButton();
            this.btnScaleSameWidth = this.Factory.CreateRibbonButton();
            this.btnScaleSameHeight = this.Factory.CreateRibbonButton();
            this.btnScaleSameSize = this.Factory.CreateRibbonButton();
            this.btnExtendSameLeft = this.Factory.CreateRibbonButton();
            this.btnExtendSameRight = this.Factory.CreateRibbonButton();
            this.btnScalePosition = this.Factory.CreateRibbonButton();
            this.btnExtendSameTop = this.Factory.CreateRibbonButton();
            this.btnExtendSameBottom = this.Factory.CreateRibbonButton();
            this.btnSnapLeft = this.Factory.CreateRibbonButton();
            this.btnSnapRight = this.Factory.CreateRibbonButton();
            this.btnSnapTop = this.Factory.CreateRibbonButton();
            this.btnSnapBottom = this.Factory.CreateRibbonButton();
            this.btnMoveForward = this.Factory.CreateRibbonButton();
            this.btnMoveFront = this.Factory.CreateRibbonButton();
            this.btnMoveBackward = this.Factory.CreateRibbonButton();
            this.btnMoveBack = this.Factory.CreateRibbonButton();
            this.btnGroup = this.Factory.CreateRibbonButton();
            this.btnUngroup = this.Factory.CreateRibbonButton();
            this.tabHome.SuspendLayout();
            this.grpArrange.SuspendLayout();
            this.grpAlignLR.SuspendLayout();
            this.grpAlignTB.SuspendLayout();
            this.grpDistributeAndRotate.SuspendLayout();
            this.grpScaleWH.SuspendLayout();
            this.grpExtendLR.SuspendLayout();
            this.grpExtendTB.SuspendLayout();
            this.grpSnapLR.SuspendLayout();
            this.grpSnapTB.SuspendLayout();
            this.grpMoveF.SuspendLayout();
            this.grpMoveB.SuspendLayout();
            this.grpGroup.SuspendLayout();
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
            this.grpArrange.Items.Add(this.grpAlignLR);
            this.grpArrange.Items.Add(this.grpAlignTB);
            this.grpArrange.Items.Add(this.grpDistributeAndRotate);
            this.grpArrange.Items.Add(this.separator1);
            this.grpArrange.Items.Add(this.grpScaleWH);
            this.grpArrange.Items.Add(this.grpExtendLR);
            this.grpArrange.Items.Add(this.grpExtendTB);
            this.grpArrange.Items.Add(this.separator2);
            this.grpArrange.Items.Add(this.grpSnapLR);
            this.grpArrange.Items.Add(this.grpSnapTB);
            this.grpArrange.Items.Add(this.separator4);
            this.grpArrange.Items.Add(this.grpMoveF);
            this.grpArrange.Items.Add(this.grpMoveB);
            this.grpArrange.Items.Add(this.grpGroup);
            this.grpArrange.Label = "Arrange";
            this.grpArrange.Name = "grpArrange";
            // 
            // grpAlignLR
            // 
            this.grpAlignLR.Items.Add(this.btnAlignLeft);
            this.grpAlignLR.Items.Add(this.btnAlignCenter);
            this.grpAlignLR.Items.Add(this.btnAlignRight);
            this.grpAlignLR.Name = "grpAlignLR";
            // 
            // grpAlignTB
            // 
            this.grpAlignTB.Items.Add(this.btnAlignTop);
            this.grpAlignTB.Items.Add(this.btnAlignMiddle);
            this.grpAlignTB.Items.Add(this.btnAlignBottom);
            this.grpAlignTB.Name = "grpAlignTB";
            // 
            // grpDistributeAndRotate
            // 
            this.grpDistributeAndRotate.Items.Add(this.btnDistributeHorizontal);
            this.grpDistributeAndRotate.Items.Add(this.btnDistributeVertical);
            this.grpDistributeAndRotate.Items.Add(this.mnuRotate);
            this.grpDistributeAndRotate.Name = "grpDistributeAndRotate";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // grpScaleWH
            // 
            this.grpScaleWH.Items.Add(this.btnScaleSameWidth);
            this.grpScaleWH.Items.Add(this.btnScaleSameHeight);
            this.grpScaleWH.Items.Add(this.btnScaleSameSize);
            this.grpScaleWH.Name = "grpScaleWH";
            // 
            // grpExtendLR
            // 
            this.grpExtendLR.Items.Add(this.btnExtendSameLeft);
            this.grpExtendLR.Items.Add(this.btnExtendSameRight);
            this.grpExtendLR.Items.Add(this.btnScalePosition);
            this.grpExtendLR.Name = "grpExtendLR";
            // 
            // grpExtendTB
            // 
            this.grpExtendTB.Items.Add(this.btnExtendSameTop);
            this.grpExtendTB.Items.Add(this.btnExtendSameBottom);
            this.grpExtendTB.Name = "grpExtendTB";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // grpSnapLR
            // 
            this.grpSnapLR.Items.Add(this.btnSnapLeft);
            this.grpSnapLR.Items.Add(this.btnSnapRight);
            this.grpSnapLR.Name = "grpSnapLR";
            // 
            // grpSnapTB
            // 
            this.grpSnapTB.Items.Add(this.btnSnapTop);
            this.grpSnapTB.Items.Add(this.btnSnapBottom);
            this.grpSnapTB.Name = "grpSnapTB";
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // grpMoveF
            // 
            this.grpMoveF.Items.Add(this.btnMoveForward);
            this.grpMoveF.Items.Add(this.btnMoveFront);
            this.grpMoveF.Name = "grpMoveF";
            // 
            // grpMoveB
            // 
            this.grpMoveB.Items.Add(this.btnMoveBackward);
            this.grpMoveB.Items.Add(this.btnMoveBack);
            this.grpMoveB.Name = "grpMoveB";
            // 
            // grpGroup
            // 
            this.grpGroup.Items.Add(this.btnGroup);
            this.grpGroup.Items.Add(this.btnUngroup);
            this.grpGroup.Name = "grpGroup";
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
            // btnScaleSameWidth
            // 
            this.btnScaleSameWidth.Image = global::ppt_arrange_addin.Properties.Resources.ScaleSameWidth;
            this.btnScaleSameWidth.Label = "Scale to same width";
            this.btnScaleSameWidth.Name = "btnScaleSameWidth";
            this.btnScaleSameWidth.ScreenTip = "Scale to same width";
            this.btnScaleSameWidth.ShowImage = true;
            this.btnScaleSameWidth.ShowLabel = false;
            // 
            // btnScaleSameHeight
            // 
            this.btnScaleSameHeight.Image = global::ppt_arrange_addin.Properties.Resources.ScaleSameHeight;
            this.btnScaleSameHeight.Label = "Scale to same height";
            this.btnScaleSameHeight.Name = "btnScaleSameHeight";
            this.btnScaleSameHeight.ScreenTip = "Scale to same height";
            this.btnScaleSameHeight.ShowImage = true;
            this.btnScaleSameHeight.ShowLabel = false;
            // 
            // btnScaleSameSize
            // 
            this.btnScaleSameSize.Image = global::ppt_arrange_addin.Properties.Resources.ScaleSameSize;
            this.btnScaleSameSize.Label = "Scale to same size";
            this.btnScaleSameSize.Name = "btnScaleSameSize";
            this.btnScaleSameSize.ScreenTip = "Scale to same size";
            this.btnScaleSameSize.ShowImage = true;
            this.btnScaleSameSize.ShowLabel = false;
            // 
            // btnExtendSameLeft
            // 
            this.btnExtendSameLeft.Image = global::ppt_arrange_addin.Properties.Resources.ExtendSameLeft;
            this.btnExtendSameLeft.Label = "Extend to same left";
            this.btnExtendSameLeft.Name = "btnExtendSameLeft";
            this.btnExtendSameLeft.ScreenTip = "Extend to same left";
            this.btnExtendSameLeft.ShowImage = true;
            this.btnExtendSameLeft.ShowLabel = false;
            // 
            // btnExtendSameRight
            // 
            this.btnExtendSameRight.Image = global::ppt_arrange_addin.Properties.Resources.ExtendSameRight;
            this.btnExtendSameRight.Label = "Extend to same right";
            this.btnExtendSameRight.Name = "btnExtendSameRight";
            this.btnExtendSameRight.ScreenTip = "Extend to same right";
            this.btnExtendSameRight.ShowImage = true;
            this.btnExtendSameRight.ShowLabel = false;
            // 
            // btnScalePosition
            // 
            this.btnScalePosition.Image = global::ppt_arrange_addin.Properties.Resources.ScaleFromMiddle;
            this.btnScalePosition.Label = "Scale from middle";
            this.btnScalePosition.Name = "btnScalePosition";
            this.btnScalePosition.ScreenTip = "Scale from middle";
            this.btnScalePosition.ShowImage = true;
            this.btnScalePosition.ShowLabel = false;
            // 
            // btnExtendSameTop
            // 
            this.btnExtendSameTop.Image = global::ppt_arrange_addin.Properties.Resources.ExtendSameTop;
            this.btnExtendSameTop.Label = "Extend to same top";
            this.btnExtendSameTop.Name = "btnExtendSameTop";
            this.btnExtendSameTop.ScreenTip = "Extend to same top";
            this.btnExtendSameTop.ShowImage = true;
            this.btnExtendSameTop.ShowLabel = false;
            // 
            // btnExtendSameBottom
            // 
            this.btnExtendSameBottom.Image = global::ppt_arrange_addin.Properties.Resources.ExtendSameBottom;
            this.btnExtendSameBottom.Label = "Extend to same bottom";
            this.btnExtendSameBottom.Name = "btnExtendSameBottom";
            this.btnExtendSameBottom.ScreenTip = "Extend to same bottom";
            this.btnExtendSameBottom.ShowImage = true;
            this.btnExtendSameBottom.ShowLabel = false;
            // 
            // btnSnapLeft
            // 
            this.btnSnapLeft.Image = global::ppt_arrange_addin.Properties.Resources.SnapToLeft;
            this.btnSnapLeft.Label = "Snap to Left";
            this.btnSnapLeft.Name = "btnSnapLeft";
            this.btnSnapLeft.ScreenTip = "Snap to Left";
            this.btnSnapLeft.ShowImage = true;
            this.btnSnapLeft.ShowLabel = false;
            // 
            // btnSnapRight
            // 
            this.btnSnapRight.Image = global::ppt_arrange_addin.Properties.Resources.SnapToRight;
            this.btnSnapRight.Label = "Snap to right";
            this.btnSnapRight.Name = "btnSnapRight";
            this.btnSnapRight.ScreenTip = "Snap to right";
            this.btnSnapRight.ShowImage = true;
            this.btnSnapRight.ShowLabel = false;
            // 
            // btnSnapTop
            // 
            this.btnSnapTop.Image = global::ppt_arrange_addin.Properties.Resources.SnapToTop;
            this.btnSnapTop.Label = "Snap to top";
            this.btnSnapTop.Name = "btnSnapTop";
            this.btnSnapTop.ScreenTip = "Snap to top";
            this.btnSnapTop.ShowImage = true;
            this.btnSnapTop.ShowLabel = false;
            // 
            // btnSnapBottom
            // 
            this.btnSnapBottom.Image = global::ppt_arrange_addin.Properties.Resources.SnapToBottom;
            this.btnSnapBottom.Label = "Snap to bottom";
            this.btnSnapBottom.Name = "btnSnapBottom";
            this.btnSnapBottom.ScreenTip = "Snap to bottom";
            this.btnSnapBottom.ShowImage = true;
            this.btnSnapBottom.ShowLabel = false;
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
            this.grpAlignLR.ResumeLayout(false);
            this.grpAlignLR.PerformLayout();
            this.grpAlignTB.ResumeLayout(false);
            this.grpAlignTB.PerformLayout();
            this.grpDistributeAndRotate.ResumeLayout(false);
            this.grpDistributeAndRotate.PerformLayout();
            this.grpScaleWH.ResumeLayout(false);
            this.grpScaleWH.PerformLayout();
            this.grpExtendLR.ResumeLayout(false);
            this.grpExtendLR.PerformLayout();
            this.grpExtendTB.ResumeLayout(false);
            this.grpExtendTB.PerformLayout();
            this.grpSnapLR.ResumeLayout(false);
            this.grpSnapLR.PerformLayout();
            this.grpSnapTB.ResumeLayout(false);
            this.grpSnapTB.PerformLayout();
            this.grpMoveF.ResumeLayout(false);
            this.grpMoveF.PerformLayout();
            this.grpMoveB.ResumeLayout(false);
            this.grpMoveB.PerformLayout();
            this.grpGroup.ResumeLayout(false);
            this.grpGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabHome;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpArrange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtendSameLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignMiddle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignCenter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDistributeHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDistributeVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtendSameRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtendSameTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtendSameBottom;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpAlignLR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpAlignTB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpDistributeAndRotate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpMoveF;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpMoveB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnuRotate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRotateRight90;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRotateLeft90;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFlipVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFlipHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpExtendLR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpExtendTB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpSnapLR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpSnapTB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpScaleWH;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnScalePosition;
    }

    partial class ThisRibbonCollection {
        internal ArrangeRibbon ArrangeRibbon {
            get { return this.GetRibbon<ArrangeRibbon>(); }
        }
    }
}
