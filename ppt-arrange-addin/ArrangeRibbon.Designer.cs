
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
            this.btnAlignLeft = this.Factory.CreateRibbonButton();
            this.btnAlignCenter = this.Factory.CreateRibbonButton();
            this.btnAlignRight = this.Factory.CreateRibbonButton();
            this.btnAlignTop = this.Factory.CreateRibbonButton();
            this.btnAlignMiddle = this.Factory.CreateRibbonButton();
            this.btnAlignBottom = this.Factory.CreateRibbonButton();
            this.btnDistributeHorizontal = this.Factory.CreateRibbonButton();
            this.btnDistributeVertical = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnExtendLeft = this.Factory.CreateRibbonButton();
            this.btnExtendRight = this.Factory.CreateRibbonButton();
            this.btnExtendTop = this.Factory.CreateRibbonButton();
            this.btnExtendBottom = this.Factory.CreateRibbonButton();
            this.btnScaleHorizontal = this.Factory.CreateRibbonButton();
            this.btnScaleVertical = this.Factory.CreateRibbonButton();
            this.btnScaleSameSize = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnGroup = this.Factory.CreateRibbonButton();
            this.btnUngroup = this.Factory.CreateRibbonButton();
            this.btnBringForward = this.Factory.CreateRibbonButton();
            this.btnBringBackward = this.Factory.CreateRibbonButton();
            this.btnBringFront = this.Factory.CreateRibbonButton();
            this.bringBehind = this.Factory.CreateRibbonButton();
            this.tabHome.SuspendLayout();
            this.grpArrange.SuspendLayout();
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
            this.grpArrange.Items.Add(this.btnAlignLeft);
            this.grpArrange.Items.Add(this.btnAlignCenter);
            this.grpArrange.Items.Add(this.btnAlignRight);
            this.grpArrange.Items.Add(this.btnAlignTop);
            this.grpArrange.Items.Add(this.btnAlignMiddle);
            this.grpArrange.Items.Add(this.btnAlignBottom);
            this.grpArrange.Items.Add(this.btnDistributeHorizontal);
            this.grpArrange.Items.Add(this.btnDistributeVertical);
            this.grpArrange.Items.Add(this.separator1);
            this.grpArrange.Items.Add(this.btnScaleHorizontal);
            this.grpArrange.Items.Add(this.btnScaleVertical);
            this.grpArrange.Items.Add(this.btnScaleSameSize);
            this.grpArrange.Items.Add(this.btnExtendLeft);
            this.grpArrange.Items.Add(this.btnExtendRight);
            this.grpArrange.Items.Add(this.btnExtendTop);
            this.grpArrange.Items.Add(this.btnExtendBottom);
            this.grpArrange.Items.Add(this.separator2);
            this.grpArrange.Items.Add(this.btnGroup);
            this.grpArrange.Items.Add(this.btnUngroup);
            this.grpArrange.Items.Add(this.btnBringForward);
            this.grpArrange.Items.Add(this.btnBringBackward);
            this.grpArrange.Items.Add(this.btnBringFront);
            this.grpArrange.Items.Add(this.bringBehind);
            this.grpArrange.Label = "Arrange";
            this.grpArrange.Name = "grpArrange";
            // 
            // btnAlignLeft
            // 
            this.btnAlignLeft.Label = "Left";
            this.btnAlignLeft.Name = "btnAlignLeft";
            // 
            // btnAlignCenter
            // 
            this.btnAlignCenter.Label = "Center";
            this.btnAlignCenter.Name = "btnAlignCenter";
            // 
            // btnAlignRight
            // 
            this.btnAlignRight.Label = "Right";
            this.btnAlignRight.Name = "btnAlignRight";
            // 
            // btnAlignTop
            // 
            this.btnAlignTop.Label = "Top";
            this.btnAlignTop.Name = "btnAlignTop";
            // 
            // btnAlignMiddle
            // 
            this.btnAlignMiddle.Label = "Middle";
            this.btnAlignMiddle.Name = "btnAlignMiddle";
            // 
            // btnAlignBottom
            // 
            this.btnAlignBottom.Label = "Bottom";
            this.btnAlignBottom.Name = "btnAlignBottom";
            // 
            // btnDistributeHorizontal
            // 
            this.btnDistributeHorizontal.Label = "Horizontal";
            this.btnDistributeHorizontal.Name = "btnDistributeHorizontal";
            // 
            // btnDistributeVertical
            // 
            this.btnDistributeVertical.Label = "Vertical";
            this.btnDistributeVertical.Name = "btnDistributeVertical";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnExtendLeft
            // 
            this.btnExtendLeft.Label = "Left";
            this.btnExtendLeft.Name = "btnExtendLeft";
            // 
            // btnExtendRight
            // 
            this.btnExtendRight.Label = "Right";
            this.btnExtendRight.Name = "btnExtendRight";
            // 
            // btnExtendTop
            // 
            this.btnExtendTop.Label = "Top";
            this.btnExtendTop.Name = "btnExtendTop";
            // 
            // btnExtendBottom
            // 
            this.btnExtendBottom.Label = "Bottom";
            this.btnExtendBottom.Name = "btnExtendBottom";
            // 
            // btnScaleHorizontal
            // 
            this.btnScaleHorizontal.Label = "Horizontal";
            this.btnScaleHorizontal.Name = "btnScaleHorizontal";
            // 
            // btnScaleVertical
            // 
            this.btnScaleVertical.Label = "Vertical";
            this.btnScaleVertical.Name = "btnScaleVertical";
            // 
            // btnScaleSameSize
            // 
            this.btnScaleSameSize.Label = "SameSize";
            this.btnScaleSameSize.Name = "btnScaleSameSize";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnGroup
            // 
            this.btnGroup.Label = "Group";
            this.btnGroup.Name = "btnGroup";
            // 
            // btnUngroup
            // 
            this.btnUngroup.Label = "Ungroup";
            this.btnUngroup.Name = "btnUngroup";
            // 
            // btnBringForward
            // 
            this.btnBringForward.Label = "Forward";
            this.btnBringForward.Name = "btnBringForward";
            // 
            // btnBringBackward
            // 
            this.btnBringBackward.Label = "Backward";
            this.btnBringBackward.Name = "btnBringBackward";
            // 
            // btnBringFront
            // 
            this.btnBringFront.Label = "Front";
            this.btnBringFront.Name = "btnBringFront";
            // 
            // bringBehind
            // 
            this.bringBehind.Label = "Behind";
            this.bringBehind.Name = "bringBehind";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnScaleHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnScaleVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnScaleSameSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUngroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBringForward;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBringBackward;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBringFront;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bringBehind;
    }

    partial class ThisRibbonCollection {
        internal ArrangeRibbon ArrangeRibbon {
            get { return this.GetRibbon<ArrangeRibbon>(); }
        }
    }
}
