
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ArrangeRibbon));
            this.tabHome = this.Factory.CreateRibbonTab();
            this.grpArrange = this.Factory.CreateRibbonGroup();
            this.grpAlignLR = this.Factory.CreateRibbonButtonGroup();
            this.btnAlignLeft = this.Factory.CreateRibbonButton();
            this.btnAlignCenter = this.Factory.CreateRibbonButton();
            this.btnAlignRight = this.Factory.CreateRibbonButton();
            this.grpAlignTB = this.Factory.CreateRibbonButtonGroup();
            this.btnAlignTop = this.Factory.CreateRibbonButton();
            this.btnAlignMiddle = this.Factory.CreateRibbonButton();
            this.btnAlignBottom = this.Factory.CreateRibbonButton();
            this.grpDistribute = this.Factory.CreateRibbonButtonGroup();
            this.btnDistributeHorizontal = this.Factory.CreateRibbonButton();
            this.btnDistributeVertical = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.grpScaleSize = this.Factory.CreateRibbonButtonGroup();
            this.btnScaleSameWidth = this.Factory.CreateRibbonButton();
            this.btnScaleSameHeight = this.Factory.CreateRibbonButton();
            this.btnScaleSameSize = this.Factory.CreateRibbonButton();
            this.btnScalePosition = this.Factory.CreateRibbonButton();
            this.grpExtendSize = this.Factory.CreateRibbonButtonGroup();
            this.btnExtendSameLeft = this.Factory.CreateRibbonButton();
            this.btnExtendSameRight = this.Factory.CreateRibbonButton();
            this.btnExtendSameTop = this.Factory.CreateRibbonButton();
            this.btnExtendSameBottom = this.Factory.CreateRibbonButton();
            this.grpSnapObjects = this.Factory.CreateRibbonButtonGroup();
            this.btnSnapLeft = this.Factory.CreateRibbonButton();
            this.btnSnapRight = this.Factory.CreateRibbonButton();
            this.btnSnapTop = this.Factory.CreateRibbonButton();
            this.btnSnapBottom = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.grpMoveLayers = this.Factory.CreateRibbonButtonGroup();
            this.btnMoveForward = this.Factory.CreateRibbonButton();
            this.btnMoveFront = this.Factory.CreateRibbonButton();
            this.btnMoveBackward = this.Factory.CreateRibbonButton();
            this.btnMoveBack = this.Factory.CreateRibbonButton();
            this.grpRotate = this.Factory.CreateRibbonButtonGroup();
            this.btnRotateRight90 = this.Factory.CreateRibbonButton();
            this.btnRotateLeft90 = this.Factory.CreateRibbonButton();
            this.btnFlipVertical = this.Factory.CreateRibbonButton();
            this.btnFlipHorizontal = this.Factory.CreateRibbonButton();
            this.grpGroup = this.Factory.CreateRibbonButtonGroup();
            this.btnGroup = this.Factory.CreateRibbonButton();
            this.btnUngroup = this.Factory.CreateRibbonButton();
            this.grpTextbox = this.Factory.CreateRibbonGroup();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.btnAutofitOff = this.Factory.CreateRibbonToggleButton();
            this.btnAutofitText = this.Factory.CreateRibbonToggleButton();
            this.btnResizeShape = this.Factory.CreateRibbonToggleButton();
            this.cbxWrapTextInTextbox = this.Factory.CreateRibbonCheckBox();
            this.btnResetShape = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.box1 = this.Factory.CreateRibbonBox();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.editBox3 = this.Factory.CreateRibbonEditBox();
            this.editBox4 = this.Factory.CreateRibbonEditBox();
            this.tabDrawingToolsFormat = this.Factory.CreateRibbonTab();
            this.tabHome.SuspendLayout();
            this.grpArrange.SuspendLayout();
            this.grpAlignLR.SuspendLayout();
            this.grpAlignTB.SuspendLayout();
            this.grpDistribute.SuspendLayout();
            this.grpScaleSize.SuspendLayout();
            this.grpExtendSize.SuspendLayout();
            this.grpSnapObjects.SuspendLayout();
            this.grpMoveLayers.SuspendLayout();
            this.grpRotate.SuspendLayout();
            this.grpGroup.SuspendLayout();
            this.grpTextbox.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.tabDrawingToolsFormat.SuspendLayout();
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
            this.grpArrange.Items.Add(this.grpDistribute);
            this.grpArrange.Items.Add(this.separator1);
            this.grpArrange.Items.Add(this.grpScaleSize);
            this.grpArrange.Items.Add(this.grpExtendSize);
            this.grpArrange.Items.Add(this.grpSnapObjects);
            this.grpArrange.Items.Add(this.separator2);
            this.grpArrange.Items.Add(this.grpMoveLayers);
            this.grpArrange.Items.Add(this.grpRotate);
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
            // grpAlignTB
            // 
            this.grpAlignTB.Items.Add(this.btnAlignTop);
            this.grpAlignTB.Items.Add(this.btnAlignMiddle);
            this.grpAlignTB.Items.Add(this.btnAlignBottom);
            this.grpAlignTB.Name = "grpAlignTB";
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
            // grpDistribute
            // 
            this.grpDistribute.Items.Add(this.btnDistributeHorizontal);
            this.grpDistribute.Items.Add(this.btnDistributeVertical);
            this.grpDistribute.Name = "grpDistribute";
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
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // grpScaleSize
            // 
            this.grpScaleSize.Items.Add(this.btnScaleSameWidth);
            this.grpScaleSize.Items.Add(this.btnScaleSameHeight);
            this.grpScaleSize.Items.Add(this.btnScaleSameSize);
            this.grpScaleSize.Items.Add(this.btnScalePosition);
            this.grpScaleSize.Name = "grpScaleSize";
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
            // btnScalePosition
            // 
            this.btnScalePosition.Image = global::ppt_arrange_addin.Properties.Resources.ScaleFromMiddle;
            this.btnScalePosition.Label = "Scale from middle";
            this.btnScalePosition.Name = "btnScalePosition";
            this.btnScalePosition.ScreenTip = "Scale from middle";
            this.btnScalePosition.ShowImage = true;
            this.btnScalePosition.ShowLabel = false;
            // 
            // grpExtendSize
            // 
            this.grpExtendSize.Items.Add(this.btnExtendSameLeft);
            this.grpExtendSize.Items.Add(this.btnExtendSameRight);
            this.grpExtendSize.Items.Add(this.btnExtendSameTop);
            this.grpExtendSize.Items.Add(this.btnExtendSameBottom);
            this.grpExtendSize.Name = "grpExtendSize";
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
            // grpSnapObjects
            // 
            this.grpSnapObjects.Items.Add(this.btnSnapLeft);
            this.grpSnapObjects.Items.Add(this.btnSnapRight);
            this.grpSnapObjects.Items.Add(this.btnSnapTop);
            this.grpSnapObjects.Items.Add(this.btnSnapBottom);
            this.grpSnapObjects.Name = "grpSnapObjects";
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
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // grpMoveLayers
            // 
            this.grpMoveLayers.Items.Add(this.btnMoveForward);
            this.grpMoveLayers.Items.Add(this.btnMoveFront);
            this.grpMoveLayers.Items.Add(this.btnMoveBackward);
            this.grpMoveLayers.Items.Add(this.btnMoveBack);
            this.grpMoveLayers.Name = "grpMoveLayers";
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
            // grpRotate
            // 
            this.grpRotate.Items.Add(this.btnRotateRight90);
            this.grpRotate.Items.Add(this.btnRotateLeft90);
            this.grpRotate.Items.Add(this.btnFlipVertical);
            this.grpRotate.Items.Add(this.btnFlipHorizontal);
            this.grpRotate.Name = "grpRotate";
            // 
            // btnRotateRight90
            // 
            this.btnRotateRight90.Image = global::ppt_arrange_addin.Properties.Resources.ObjectRotateRight90;
            this.btnRotateRight90.Label = "Rotate right with 90 degrees";
            this.btnRotateRight90.Name = "btnRotateRight90";
            this.btnRotateRight90.ScreenTip = "Rotate right with 90 degrees";
            this.btnRotateRight90.ShowImage = true;
            this.btnRotateRight90.ShowLabel = false;
            // 
            // btnRotateLeft90
            // 
            this.btnRotateLeft90.Image = global::ppt_arrange_addin.Properties.Resources.ObjectRotateLeft90;
            this.btnRotateLeft90.Label = "Rorate left with 90 degrees";
            this.btnRotateLeft90.Name = "btnRotateLeft90";
            this.btnRotateLeft90.ScreenTip = "Rorate left with 90 degrees";
            this.btnRotateLeft90.ShowImage = true;
            this.btnRotateLeft90.ShowLabel = false;
            // 
            // btnFlipVertical
            // 
            this.btnFlipVertical.Image = global::ppt_arrange_addin.Properties.Resources.ObjectFlipVertical;
            this.btnFlipVertical.Label = "Flip vertically";
            this.btnFlipVertical.Name = "btnFlipVertical";
            this.btnFlipVertical.ScreenTip = "Flip vertically";
            this.btnFlipVertical.ShowImage = true;
            this.btnFlipVertical.ShowLabel = false;
            // 
            // btnFlipHorizontal
            // 
            this.btnFlipHorizontal.Image = global::ppt_arrange_addin.Properties.Resources.ObjectFlipHorizontal;
            this.btnFlipHorizontal.Label = "Flip horizontally";
            this.btnFlipHorizontal.Name = "btnFlipHorizontal";
            this.btnFlipHorizontal.ScreenTip = "Flip horizontally";
            this.btnFlipHorizontal.ShowImage = true;
            this.btnFlipHorizontal.ShowLabel = false;
            // 
            // grpGroup
            // 
            this.grpGroup.Items.Add(this.btnGroup);
            this.grpGroup.Items.Add(this.btnUngroup);
            this.grpGroup.Name = "grpGroup";
            // 
            // btnGroup
            // 
            this.btnGroup.Image = ((System.Drawing.Image)(resources.GetObject("btnGroup.Image")));
            this.btnGroup.Label = "Group Shapes";
            this.btnGroup.Name = "btnGroup";
            this.btnGroup.ScreenTip = "Group Shapes";
            this.btnGroup.ShowImage = true;
            this.btnGroup.ShowLabel = false;
            // 
            // btnUngroup
            // 
            this.btnUngroup.Image = ((System.Drawing.Image)(resources.GetObject("btnUngroup.Image")));
            this.btnUngroup.Label = "Ungroup Shapes";
            this.btnUngroup.Name = "btnUngroup";
            this.btnUngroup.ScreenTip = "Ungroup Shapes";
            this.btnUngroup.ShowImage = true;
            this.btnUngroup.ShowLabel = false;
            // 
            // grpTextbox
            // 
            this.grpTextbox.Items.Add(this.buttonGroup1);
            this.grpTextbox.Items.Add(this.cbxWrapTextInTextbox);
            this.grpTextbox.Items.Add(this.btnResetShape);
            this.grpTextbox.Items.Add(this.separator3);
            this.grpTextbox.Items.Add(this.box1);
            this.grpTextbox.Items.Add(this.box2);
            this.grpTextbox.Label = "Textbox";
            this.grpTextbox.Name = "grpTextbox";
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.btnAutofitOff);
            this.buttonGroup1.Items.Add(this.btnAutofitText);
            this.buttonGroup1.Items.Add(this.btnResizeShape);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // btnAutofitOff
            // 
            this.btnAutofitOff.Label = "Autofit off";
            this.btnAutofitOff.Name = "btnAutofitOff";
            // 
            // btnAutofitText
            // 
            this.btnAutofitText.Checked = true;
            this.btnAutofitText.Label = "Autofit text";
            this.btnAutofitText.Name = "btnAutofitText";
            // 
            // btnResizeShape
            // 
            this.btnResizeShape.Label = "Resize shape";
            this.btnResizeShape.Name = "btnResizeShape";
            // 
            // cbxWrapTextInTextbox
            // 
            this.cbxWrapTextInTextbox.Checked = true;
            this.cbxWrapTextInTextbox.Label = "Wrap text in textbox";
            this.cbxWrapTextInTextbox.Name = "cbxWrapTextInTextbox";
            // 
            // btnResetShape
            // 
            this.btnResetShape.Label = "Reset shape";
            this.btnResetShape.Name = "btnResetShape";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.editBox1);
            this.box1.Items.Add(this.editBox2);
            this.box1.Name = "box1";
            // 
            // editBox1
            // 
            this.editBox1.Label = "Left Margin:";
            this.editBox1.Name = "editBox1";
            this.editBox1.Text = "0.00 cm";
            // 
            // editBox2
            // 
            this.editBox2.Label = "Right Margin:";
            this.editBox2.Name = "editBox2";
            this.editBox2.Text = "0.00 cm";
            // 
            // box2
            // 
            this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box2.Items.Add(this.editBox3);
            this.box2.Items.Add(this.editBox4);
            this.box2.Name = "box2";
            // 
            // editBox3
            // 
            this.editBox3.Label = "Top Margin:";
            this.editBox3.Name = "editBox3";
            this.editBox3.Text = "0.00 cm";
            // 
            // editBox4
            // 
            this.editBox4.Label = "Bottom Margin:";
            this.editBox4.Name = "editBox4";
            this.editBox4.Text = "0.00 cm";
            // 
            // tabDrawingToolsFormat
            // 
            this.tabDrawingToolsFormat.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabDrawingToolsFormat.ControlId.OfficeId = "TabDrawingToolsFormat";
            this.tabDrawingToolsFormat.Groups.Add(this.grpTextbox);
            this.tabDrawingToolsFormat.Label = "TabDrawingToolsFormat";
            this.tabDrawingToolsFormat.Name = "tabDrawingToolsFormat";
            // 
            // ArrangeRibbon
            // 
            this.Name = "ArrangeRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabHome);
            this.Tabs.Add(this.tabDrawingToolsFormat);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ArrangeRibbon_Load);
            this.tabHome.ResumeLayout(false);
            this.tabHome.PerformLayout();
            this.grpArrange.ResumeLayout(false);
            this.grpArrange.PerformLayout();
            this.grpAlignLR.ResumeLayout(false);
            this.grpAlignLR.PerformLayout();
            this.grpAlignTB.ResumeLayout(false);
            this.grpAlignTB.PerformLayout();
            this.grpDistribute.ResumeLayout(false);
            this.grpDistribute.PerformLayout();
            this.grpScaleSize.ResumeLayout(false);
            this.grpScaleSize.PerformLayout();
            this.grpExtendSize.ResumeLayout(false);
            this.grpExtendSize.PerformLayout();
            this.grpSnapObjects.ResumeLayout(false);
            this.grpSnapObjects.PerformLayout();
            this.grpMoveLayers.ResumeLayout(false);
            this.grpMoveLayers.PerformLayout();
            this.grpRotate.ResumeLayout(false);
            this.grpRotate.PerformLayout();
            this.grpGroup.ResumeLayout(false);
            this.grpGroup.PerformLayout();
            this.grpTextbox.ResumeLayout(false);
            this.grpTextbox.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.tabDrawingToolsFormat.ResumeLayout(false);
            this.tabDrawingToolsFormat.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveForward;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveBackward;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveFront;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveBack;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSnapLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSnapRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSnapTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSnapBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpAlignLR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpAlignTB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpDistribute;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpMoveLayers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRotateRight90;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRotateLeft90;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFlipVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFlipHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpExtendSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpSnapObjects;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpRotate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpScaleSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnScalePosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup grpGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUngroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTextbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnAutofitOff;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnResizeShape;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnAutofitText;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbxWrapTextInTextbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox4;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabDrawingToolsFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetShape;
    }

    partial class ThisRibbonCollection {
        internal ArrangeRibbon ArrangeRibbon {
            get { return this.GetRibbon<ArrangeRibbon>(); }
        }
    }
}
