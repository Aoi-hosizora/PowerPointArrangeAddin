using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ppt_arrange_addin {

    public partial class ArrangeRibbon {

        private void ArrangeRibbon_Load(object sender, RibbonUIEventArgs e) {
            AdjustButtonsAccessibility();
            btnAlignLeft.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignCenter.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignRight.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignTop.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignMiddle.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignBottom.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnDistributeHorizontal.Click += new RibbonControlEventHandler(BtnDistribute_Click);
            btnDistributeVertical.Click += new RibbonControlEventHandler(BtnDistribute_Click);
            btnScaleSameWidth.Click += new RibbonControlEventHandler(BtnScale_Click);
            btnScaleSameHeight.Click += new RibbonControlEventHandler(BtnScale_Click);
            btnScaleSameSize.Click += new RibbonControlEventHandler(BtnScale_Click);
            btnScalePosition.Click += new RibbonControlEventHandler(BtnScalePosition_Click);
            btnExtendSameLeft.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnExtendSameRight.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnExtendSameTop.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnExtendSameBottom.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnSnapLeft.Click += new RibbonControlEventHandler(BtnSnap_Click);
            btnSnapRight.Click += new RibbonControlEventHandler(BtnSnap_Click);
            btnSnapTop.Click += new RibbonControlEventHandler(BtnSnap_Click);
            btnSnapBottom.Click += new RibbonControlEventHandler(BtnSnap_Click);
            btnMoveForward.Click += new RibbonControlEventHandler(BtnMove_Click);
            btnMoveBackward.Click += new RibbonControlEventHandler(BtnMove_Click);
            btnMoveFront.Click += new RibbonControlEventHandler(BtnMove_Click);
            btnMoveBack.Click += new RibbonControlEventHandler(BtnMove_Click);
            btnRotateLeft90.Click += new RibbonControlEventHandler(BtnRotate_Click);
            btnRotateRight90.Click += new RibbonControlEventHandler(BtnRotate_Click);
            btnFlipHorizontal.Click += new RibbonControlEventHandler(BtnFlip_Click);
            btnFlipVertical.Click += new RibbonControlEventHandler(BtnFlip_Click);
            btnGroup.Click += new RibbonControlEventHandler(BtnGroup_Click);
            btnUngroup.Click += new RibbonControlEventHandler(BtnGroup_Click);
        }

        public void AdjustButtonsAccessibility() {
            int selectedCount;
            try {
                var shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.OfType<PowerPoint.Shape>().ToArray();
                selectedCount = shapeRange.Count();
            } catch (Exception) {
                selectedCount = 0;
            }
            btnAlignLeft.Enabled = selectedCount >= 1;
            btnAlignCenter.Enabled = selectedCount >= 1;
            btnAlignRight.Enabled = selectedCount >= 1;
            btnAlignTop.Enabled = selectedCount >= 1;
            btnAlignMiddle.Enabled = selectedCount >= 1;
            btnAlignBottom.Enabled = selectedCount >= 1;
            btnDistributeHorizontal.Enabled = selectedCount >= 3;
            btnDistributeVertical.Enabled = selectedCount >= 3;
            btnRotateLeft90.Enabled = selectedCount >= 1;
            btnRotateRight90.Enabled = selectedCount >= 1;
            btnFlipHorizontal.Enabled = selectedCount >= 1;
            btnFlipVertical.Enabled = selectedCount >= 1;
            btnScalePosition.Enabled = true;
            if (btnScalePosition.Tag == null) btnScalePosition.Tag = Office.MsoScaleFrom.msoScaleFromMiddle;
            btnScaleSameWidth.Enabled = selectedCount >= 2;
            btnScaleSameHeight.Enabled = selectedCount >= 2;
            btnScaleSameSize.Enabled = selectedCount >= 2;
            btnExtendSameLeft.Enabled = selectedCount >= 2;
            btnExtendSameRight.Enabled = selectedCount >= 2;
            btnExtendSameTop.Enabled = selectedCount >= 2;
            btnExtendSameBottom.Enabled = selectedCount >= 2;
            btnMoveForward.Enabled = selectedCount >= 1;
            btnMoveBackward.Enabled = selectedCount >= 1;
            btnMoveFront.Enabled = selectedCount >= 1;
            btnMoveBack.Enabled = selectedCount >= 1;
            btnSnapLeft.Enabled = selectedCount >= 2;
            btnSnapRight.Enabled = selectedCount >= 2;
            btnSnapTop.Enabled = selectedCount >= 2;
            btnSnapBottom.Enabled = selectedCount >= 2;
            btnGroup.Enabled = selectedCount >= 1;
            btnUngroup.Enabled = selectedCount >= 1;
        }

        private PowerPoint.ShapeRange GetShapeRange(int mustMoreThanOrEqualTo = 1) {
            var shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            if (shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            return shapeRange;
        }

        private void StartNewUndoEntry() {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
        }

        private void BtnAlign_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            Office.MsoAlignCmd cmd;
            switch (e.Control.Id) {
            case "btnAlignLeft":
                cmd = Office.MsoAlignCmd.msoAlignLefts;
                break;
            case "btnAlignCenter":
                cmd = Office.MsoAlignCmd.msoAlignCenters;
                break;
            case "btnAlignRight":
                cmd = Office.MsoAlignCmd.msoAlignRights;
                break;
            case "btnAlignTop":
                cmd = Office.MsoAlignCmd.msoAlignTops;
                break;
            case "btnAlignMiddle":
                cmd = Office.MsoAlignCmd.msoAlignMiddles;
                break;
            case "btnAlignBottom":
                cmd = Office.MsoAlignCmd.msoAlignBottoms;
                break;
            default:
                return;
            }

            StartNewUndoEntry();
            var flag = shapeRange.Count == 1 ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            shapeRange.Align(cmd, flag);
        }

        private void BtnDistribute_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = GetShapeRange(mustMoreThanOrEqualTo: 3);
            if (shapeRange == null) {
                return;
            }

            Office.MsoDistributeCmd cmd;
            switch (e.Control.Id) {
            case "btnDistributeHorizontal":
                cmd = Office.MsoDistributeCmd.msoDistributeHorizontally;
                break;
            case "btnDistributeVertical":
                cmd = Office.MsoDistributeCmd.msoDistributeVertically;
                break;
            default:
                return;
            }

            StartNewUndoEntry();
            shapeRange.Distribute(cmd, Office.MsoTriState.msoFalse);
        }

        private void BtnRotate_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            StartNewUndoEntry();
            switch (e.Control.Id) {
            case "btnRotateLeft90":
                shapeRange.IncrementRotation(-90);
                break;
            case "btnRotateRight90":
                shapeRange.IncrementRotation(90);
                break;
            }
        }

        private void BtnFlip_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            StartNewUndoEntry();
            switch (e.Control.Id) {
            case "btnFlipHorizontal":
                shapeRange.Flip(Office.MsoFlipCmd.msoFlipHorizontal);
                break;
            case "btnFlipVertical":
                shapeRange.Flip(Office.MsoFlipCmd.msoFlipVertical);
                break;
            }
        }

        private void BtnScalePosition_Click(object sender, RibbonControlEventArgs e) {
            var tag = (btnScalePosition.Tag as Office.MsoScaleFrom?) ?? Office.MsoScaleFrom.msoScaleFromMiddle;
            if (tag == Office.MsoScaleFrom.msoScaleFromMiddle) {
                btnScalePosition.Tag = Office.MsoScaleFrom.msoScaleFromTopLeft;
                btnScalePosition.Image = Properties.Resources.ScaleFromLeftTop;
                btnScalePosition.Label = "Scale from left top";
                btnScalePosition.ScreenTip = "Scale from left top";
            } else {
                btnScalePosition.Tag = Office.MsoScaleFrom.msoScaleFromMiddle;
                btnScalePosition.Image = Properties.Resources.ScaleFromMiddle;
                btnScalePosition.Label = "Scale from middle";
                btnScalePosition.ScreenTip = "Scale from middle";
            }
        }

        private void BtnScale_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = GetShapeRange(mustMoreThanOrEqualTo: 2);
            if (shapeRange == null) {
                return;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
            var (firstWidth, firstHeight) = (shapes[0].Width, shapes[0].Height);
            var scaleFrom = (btnScalePosition.Tag as Office.MsoScaleFrom?) ?? Office.MsoScaleFrom.msoScaleFromMiddle;

            StartNewUndoEntry();
            switch (e.Control.Id) {
            case "btnScaleSameWidth":
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, scaleFrom);
                }
                break;
            case "btnScaleSameHeight":
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstHeight / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, scaleFrom);
                }
                break;
            case "btnScaleSameSize":
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, scaleFrom);
                    ratio = firstHeight / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, scaleFrom);
                }
                break;
            }
        }

        private void BtnExtend_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = GetShapeRange(mustMoreThanOrEqualTo: 2);
            if (shapeRange == null) {
                return;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
            float minLeft = 0x7fffffff, minTop = 0x7fffffff, maxLeftWidth = -1, maxTopHeight = -1;
            foreach (var shape in shapes) {
                minLeft = Math.Min(minLeft, shape.Left);
                minTop = Math.Min(minTop, shape.Top);
                maxLeftWidth = Math.Max(maxLeftWidth, shape.Left + shape.Width);
                maxTopHeight = Math.Max(maxTopHeight, shape.Top + shape.Height);
            }

            StartNewUndoEntry();
            switch (e.Control.Id) {
            case "btnExtendSameLeft":
                foreach (var shape in shapes) {
                    var newWidth = shape.Width + shape.Left - minLeft;
                    var ratio = newWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case "btnExtendSameRight":
                foreach (var shape in shapes) {
                    var newWidth = maxLeftWidth - shape.Left;
                    var ratio = newWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            case "btnExtendSameTop":
                foreach (var shape in shapes) {
                    var newTop = shape.Height + shape.Top - minTop;
                    var ratio = newTop / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case "btnExtendSameBottom":
                foreach (var shape in shapes) {
                    var newHeight = maxTopHeight - shape.Top;
                    var ratio = newHeight / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            }
        }

        private void BtnMove_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            Office.MsoZOrderCmd cmd;
            switch (e.Control.Id) {
            case "btnMoveForward":
                cmd = Office.MsoZOrderCmd.msoBringForward;
                break;
            case "btnMoveBackward":
                cmd = Office.MsoZOrderCmd.msoSendBackward;
                break;
            case "btnMoveFront":
                cmd = Office.MsoZOrderCmd.msoBringToFront;
                break;
            case "btnMoveBack":
                cmd = Office.MsoZOrderCmd.msoSendToBack;
                break;
            default:
                return;
            }

            StartNewUndoEntry();
            shapeRange.ZOrder(cmd);
        }

        private void BtnSnap_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = GetShapeRange(mustMoreThanOrEqualTo: 2);
            if (shapeRange == null) {
                return;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
            var (lastLeft, lastTop) = (shapes[0].Left, shapes[0].Top);
            var (lastWidth, lastHeight) = (shapes[0].Width, shapes[0].Height);

            StartNewUndoEntry();
            switch (e.Control.Id) {
            case "btnSnapLeft":
                for (var i = 1; i < shapes.Count(); i++) {
                    shapes[i].Left = lastLeft + lastWidth;
                    lastLeft = shapes[i].Left;
                    lastWidth = shapes[i].Width;
                }
                break;
            case "btnSnapRight":
                for (var i = 1; i < shapes.Count(); i++) {
                    lastWidth = shapes[i].Width;
                    shapes[i].Left = lastLeft - lastWidth;
                    lastLeft = shapes[i].Left;
                }
                break;
            case "btnSnapTop":
                for (var i = 1; i < shapes.Count(); i++) {
                    shapes[i].Top = lastTop + lastHeight;
                    lastTop = shapes[i].Top;
                    lastHeight = shapes[i].Height;
                }
                break;
            case "btnSnapBottom":
                for (var i = 1; i < shapes.Count(); i++) {
                    lastHeight = shapes[i].Height;
                    shapes[i].Top = lastTop - lastHeight;
                    lastTop = shapes[i].Top;
                }
                break;
            default:
                return;
            }
        }
        private void BtnGroup_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            switch (e.Control.Id) {
            case "btnGroup":
                if (shapeRange.Count >= 2) {
                    StartNewUndoEntry();
                    var grouped = shapeRange.Group();
                    grouped.Select();
                    AdjustButtonsAccessibility();
                }
                break;
            case "btnUngroup":
                StartNewUndoEntry();
                var ungrouped = shapeRange.Ungroup();
                ungrouped.Select();
                AdjustButtonsAccessibility();
                break;
            }
        }

    }

}
