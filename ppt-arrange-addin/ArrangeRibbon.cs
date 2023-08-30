using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ppt_arrange_addin {

    public partial class ArrangeRibbon {

        private void ArrangeRibbon_Load(object sender, RibbonUIEventArgs e) {
            btnAlignLeft.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignCenter.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignRight.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignTop.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignMiddle.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnAlignBottom.Click += new RibbonControlEventHandler(BtnAlign_Click);
            btnDistributeHorizontal.Click += new RibbonControlEventHandler(BtnDistribute_Click);
            btnDistributeVertical.Click += new RibbonControlEventHandler(BtnDistribute_Click);
            btnScaleHorizontal.Click += new RibbonControlEventHandler(BtnScale_Click);
            btnScaleVertical.Click += new RibbonControlEventHandler(BtnScale_Click);
            btnScaleSameSize.Click += new RibbonControlEventHandler(BtnScale_Click);
            btnExtendLeft.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnExtendRight.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnExtendTop.Click += new RibbonControlEventHandler(BtnExtend_Click);
            btnExtendBottom.Click += new RibbonControlEventHandler(BtnExtend_Click);
        }

        private ShapeRange getShapeRange(int mustMoreThanOrEqualTo = 1) {
            var shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            if (shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            return shapeRange;
        }

        private void BtnAlign_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange(mustMoreThanOrEqualTo: 2);
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
            shapeRange.Align(cmd, Microsoft.Office.Core.MsoTriState.msoFalse);
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

        private void BtnScale_Click(object sender, RibbonControlEventArgs e) {
            var shapeRange = getShapeRange(mustMoreThanOrEqualTo: 2);
            if (shapeRange == null) {
                return;
            }

            var shapes = shapeRange.OfType<Shape>().ToArray();
            var (firstWidth, firstHeight) = (shapes[0].Width, shapes[0].Height);
            switch (e.Control.Id) {
            case "btnScaleHorizontal":
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstWidth / shape.Width;
                    shape.ScaleWidth(ratio, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                }
                break;
            case "btnScaleVertical":
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

    }

}
