using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ppt_arrange_addin {

    public partial class ArrangeRibbon {

        private void ArrangeRibbon_Load(object sender, RibbonUIEventArgs e) {
            AdjustButtonsEnabled(); // TODO useless, how to initialize ???
            btnAlignLeft.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignCenter.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignRight.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignTop.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignMiddle.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignBottom.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnDistributeHorizontal.Click += new RibbonControlEventHandler(BtnDistribute_Click);
            btnDistributeVertical.Click += new RibbonControlEventHandler(BtnDistribute_Click);
            btnRotateLeft90.Click += new RibbonControlEventHandler(BtnRotate_Click);
            btnRotateRight90.Click += new RibbonControlEventHandler(BtnRotate_Click);
            btnFlipHorizontal.Click += new RibbonControlEventHandler(BtnFlip_Click);
            btnFlipVertical.Click += new RibbonControlEventHandler(BtnFlip_Click);
            btnScaleSameWidth.Click += new RibbonControlEventHandler(BtnScale_Click);
            btnScaleSameHeight.Click += new RibbonControlEventHandler(BtnScale_Click);
            btnScaleSameSize.Click += new RibbonControlEventHandler(BtnScale_Click);
            btnExtendLeft.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnExtendRight.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnExtendTop.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnExtendBottom.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnMoveForward.Click += new RibbonControlEventHandler(BtnMove_Click);
            btnMoveBackward.Click += new RibbonControlEventHandler(BtnMove_Click);
            btnMoveFront.Click += new RibbonControlEventHandler(BtnMove_Click);
            btnMoveBack.Click += new RibbonControlEventHandler(BtnMove_Click);
            btnGroup.Click += new RibbonControlEventHandler(BtnGroup_Click);
            btnUngroup.Click += new RibbonControlEventHandler(BtnGroup_Click);
            btnSnapLeft.Click += new RibbonControlEventHandler(BtnSnap_Click);
            btnSnapRight.Click += new RibbonControlEventHandler(BtnSnap_Click);
            btnSnapTop.Click += new RibbonControlEventHandler(BtnSnap_Click);
            btnSnapBottom.Click += new RibbonControlEventHandler(BtnSnap_Click);
        }

        public void AdjustButtonsEnabled() {
            Shape[] shapeRange;
            int selectedCount;
            try {
                shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.OfType<Shape>().ToArray();
                selectedCount = shapeRange.Count();
            } catch (Exception) {
                return;
            }
            btnAlignLeft.Enabled = selectedCount >= 1;
            btnAlignCenter.Enabled = selectedCount >= 1;
            btnAlignRight.Enabled = selectedCount >= 1;
            btnAlignTop.Enabled = selectedCount >= 1;
            btnAlignMiddle.Enabled = selectedCount >= 1;
            btnAlignBottom.Enabled = selectedCount >= 1;
            btnDistributeHorizontal.Enabled = selectedCount >= 3;
            btnDistributeVertical.Enabled = selectedCount >= 3;
            mnuRotate.Enabled = selectedCount >= 1;
            btnRotateLeft90.Enabled = selectedCount >= 1;
            btnRotateRight90.Enabled = selectedCount >= 1;
            btnFlipHorizontal.Enabled = selectedCount >= 1;
            btnFlipVertical.Enabled = selectedCount >= 1;
            btnScaleSameWidth.Enabled = selectedCount >= 2;
            btnScaleSameHeight.Enabled = selectedCount >= 2;
            btnScaleSameSize.Enabled = selectedCount >= 2;
            btnExtendLeft.Enabled = selectedCount >= 2;
            btnExtendRight.Enabled = selectedCount >= 2;
            btnExtendTop.Enabled = selectedCount >= 2;
            btnExtendBottom.Enabled = selectedCount >= 2;
            btnMoveForward.Enabled = selectedCount >= 1;
            btnMoveBackward.Enabled = selectedCount >= 1;
            btnMoveFront.Enabled = selectedCount >= 1;
            btnMoveBack.Enabled = selectedCount >= 1;
            btnGroup.Enabled = selectedCount >= 2;
            btnUngroup.Enabled = selectedCount >= 1; // TODO check ungroup enabled ???
            //System.Diagnostics.Debug.WriteLine($"shapeRange: {shapeRange.Count()}");
            //System.Diagnostics.Debug.WriteLine($"=> {string.Join(",", shapeRange.Select((s) => s.GroupItems))}");
            //btnUngroup.Enabled = selectedCount >= 1 && shapeRange.Any((s) => s.GroupItems.Count >= 2); // <<<
            btnSnapLeft.Enabled = selectedCount >= 2;
            btnSnapRight.Enabled = selectedCount >= 2;
            btnSnapTop.Enabled = selectedCount >= 2;
            btnSnapBottom.Enabled = selectedCount >= 2;
        }

        private ShapeRange getShapeRange(int mustMoreThanOrEqualTo = 1) {
            var shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            if (shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            return shapeRange;
        }

        private void BtnAlign_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange();
            if (shapeRange == null) {
                return;
            }

            Microsoft.Office.Core.MsoAlignCmd cmd;
            switch (e.Control.Id) {
            case "btnAlignLeft":
                cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignLefts;
                break;
            case "btnAlignCenter":
                cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters;
                break;
            case "btnAlignRight":
                cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignRights;
                break;
            case "btnAlignTop":
                cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignTops;
                break;
            case "btnAlignMiddle":
                cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignMiddles;
                break;
            case "btnAlignBottom":
                cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignBottoms;
                break;
            default:
                return;
            }
            var flag = shapeRange.Count == 1 ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            shapeRange.Align(cmd, flag);
        }

        private void BtnDistribute_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange(mustMoreThanOrEqualTo: 3);
            if (shapeRange == null) {
                return;
            }

            Microsoft.Office.Core.MsoDistributeCmd cmd;
            switch (e.Control.Id) {
            case "btnDistributeHorizontal":
                cmd = Microsoft.Office.Core.MsoDistributeCmd.msoDistributeHorizontally;
                break;
            case "btnDistributeVertical":
                cmd = Microsoft.Office.Core.MsoDistributeCmd.msoDistributeVertically;
                break;
            default:
                return;
            }

            shapeRange.Distribute(cmd, Microsoft.Office.Core.MsoTriState.msoFalse);
        }

        private void BtnRotate_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange();
            if (shapeRange == null) {
                return;
            }

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
            var shapeRange = getShapeRange();
            if (shapeRange == null) {
                return;
            }

            switch (e.Control.Id) {
            case "btnFlipHorizontal":
                shapeRange.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipHorizontal);
                break;
            case "btnFlipVertical":
                shapeRange.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipVertical);
                break;
            }
        }

        private void BtnScale_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange(mustMoreThanOrEqualTo: 2);
            if (shapeRange == null) {
                return;
            }

            var shapes = shapeRange.OfType<Shape>().ToArray();
            var (firstWidth, firstHeight) = (shapes[0].Width, shapes[0].Height);
            switch (e.Control.Id) {
            case "btnScaleSameWidth":
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstWidth / shape.Width;
                    shape.ScaleWidth(ratio, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                }
                break;
            case "btnScaleSameHeight":
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstHeight / shape.Height;
                    shape.ScaleHeight(ratio, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                }
                break;
            case "btnScaleSameSize":
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstWidth / shape.Width;
                    shape.ScaleWidth(ratio, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                    ratio = firstHeight / shape.Height;
                    shape.ScaleHeight(ratio, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                }
                break;
            }
        }

        private void BtnExtend_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange(mustMoreThanOrEqualTo: 2);
            if (shapeRange == null) {
                return;
            }

            var shapes = shapeRange.OfType<Shape>().ToArray();
            float minLeft = 0x7fffffff, minTop = 0x7fffffff, maxLeftWidth = -1, maxTopHeight = -1;
            foreach (var shape in shapes) {
                minLeft = Math.Min(minLeft, shape.Left);
                minTop = Math.Min(minTop, shape.Top);
                maxLeftWidth = Math.Max(maxLeftWidth, shape.Left + shape.Width);
                maxTopHeight = Math.Max(maxTopHeight, shape.Top + shape.Height);
            }

            switch (e.Control.Id) {
            case "btnExtendLeft":
                foreach (var shape in shapes) {
                    var newWidth = shape.Width + shape.Left - minLeft;
                    var ratio = newWidth / shape.Width;
                    shape.ScaleWidth(ratio, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case "btnExtendRight":
                foreach (var shape in shapes) {
                    var newWidth = maxLeftWidth - shape.Left;
                    var ratio = newWidth / shape.Width;
                    shape.ScaleWidth(ratio, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            case "btnExtendTop":
                foreach (var shape in shapes) {
                    var newTop = shape.Height + shape.Top - minTop;
                    var ratio = newTop / shape.Height;
                    shape.ScaleHeight(ratio, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case "btnExtendBottom":
                foreach (var shape in shapes) {
                    var newHeight = maxTopHeight - shape.Top;
                    var ratio = newHeight / shape.Height;
                    shape.ScaleHeight(ratio, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            }
        }

        private void BtnMove_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange();
            if (shapeRange == null) {
                return;
            }

            Microsoft.Office.Core.MsoZOrderCmd cmd;
            switch (e.Control.Id) {
            case "btnMoveForward":
                cmd = Microsoft.Office.Core.MsoZOrderCmd.msoBringForward;
                break;
            case "btnMoveBackward":
                cmd = Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward;
                break;
            case "btnMoveFront":
                cmd = Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront;
                break;
            case "btnMoveBack":
                cmd = Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack;
                break;
            default:
                return;
            }
            shapeRange.ZOrder(cmd);
        }

        private void BtnGroup_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange();
            if (shapeRange == null) {
                return;
            }

            switch (e.Control.Id) {
            case "btnGroup":
                if (shapeRange.Count >= 2) {
                    var grouped = shapeRange.Group();
                    grouped.Select();
                    AdjustButtonsEnabled();
                }
                break;
            case "btnUngroup":
                var ungrouped = shapeRange.Ungroup();
                ungrouped.Select();
                AdjustButtonsEnabled();
                break;
            }
        }


        private void BtnSnap_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange(mustMoreThanOrEqualTo: 2);
            if (shapeRange == null) {
                return;
            }

            var shapes = shapeRange.OfType<Shape>().ToArray();
            var (lastLeft, lastTop) = (shapes[0].Left, shapes[0].Top);
            var (lastWidth, lastHeight) = (shapes[0].Width, shapes[0].Height);

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

    }

}
