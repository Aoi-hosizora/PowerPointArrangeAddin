using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ppt_arrange_addin {

    public partial class ArrangeRibbon {

        private Office.IRibbonUI ribbon;

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            ribbon = ribbonUI;
        }

        private Dictionary<string, Func<int, bool>> elementsAvailabilityRules = new Dictionary<string, Func<int, bool>>() {
            { btnAlignLeft, (count) => count >= 1 },
            { btnAlignCenter, (count) => count >= 1 },
            { btnAlignRight, (count) => count >= 1 },
            { btnAlignTop, (count) => count >= 1 },
            { btnAlignMiddle, (count) => count >= 1 },
            { btnAlignBottom, (count) => count >= 1 },
            { btnDistributeHorizontal, (count) => count >= 3 },
            { btnDistributeVertical, (count) => count >= 3 },
            { btnScaleSameWidth, (count) => count >= 2 },
            { btnScaleSameHeight, (count) => count >= 2 },
            { btnScaleSameSize, (count) => count >= 2 },
            { btnScalePosition, (count) => count >= 1 },
            { btnExtendSameLeft, (count) => count >= 2 },
            { btnExtendSameRight, (count) => count >= 2 },
            { btnExtendSameTop, (count) => count >= 2 },
            { btnExtendSameBottom, (count) => count >= 2 },
            { btnSnapLeft, (count) => count >= 2 },
            { btnSnapRight, (count) => count >= 2 },
            { btnSnapTop, (count) => count >= 2 },
            { btnSnapBottom, (count) => count >= 2 },
            { btnMoveForward, (count) => count >= 1 },
            { btnMoveFront, (count) => count >= 1 },
            { btnMoveBackward, (count) => count >= 1 },
            { btnMoveBack, (count) => count >= 1 },
            { btnRotateRight90, (count) => count >= 1 },
            { btnRotateLeft90, (count) => count >= 1 },
            { btnFlipVertical, (count) => count >= 1 },
            { btnFlipHorizontal, (count) => count >= 1 },
            { btnGroup, (count) => count >= 2 },
            { btnUngroup, (count) => count >= 1 },
        };

        public bool GetEnabled(Office.IRibbonControl ribbonControl) {
            int selectedCount = 0;
            try {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) {
                    selectedCount = selection.ShapeRange.Count;
                }
            } catch (Exception e) {
                System.Diagnostics.Debug.WriteLine(e);
            }
            elementsAvailabilityRules.TryGetValue(ribbonControl.Id, out Func<int, bool> checker);
            return checker?.Invoke(selectedCount) ?? true;
        }

        public void AdjustRibbonButtonsAvailability() {
            ribbon.Invalidate();
        }

        private PowerPoint.ShapeRange GetShapeRange(int mustMoreThanOrEqualTo = 1, bool mustHasTextFrame = false) {
            var shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            if (shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            if (mustHasTextFrame && shapeRange.HasTextFrame != Office.MsoTriState.msoTrue) {
                return null;
            }
            return shapeRange;
        }

        private void StartNewUndoEntry() {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
        }

        public void BtnAlign_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            Office.MsoAlignCmd cmd;
            switch (ribbonControl.Id) {
            case btnAlignLeft:
                cmd = Office.MsoAlignCmd.msoAlignLefts;
                break;
            case btnAlignCenter:
                cmd = Office.MsoAlignCmd.msoAlignCenters;
                break;
            case btnAlignRight:
                cmd = Office.MsoAlignCmd.msoAlignRights;
                break;
            case btnAlignTop:
                cmd = Office.MsoAlignCmd.msoAlignTops;
                break;
            case btnAlignMiddle:
                cmd = Office.MsoAlignCmd.msoAlignMiddles;
                break;
            case btnAlignBottom:
                cmd = Office.MsoAlignCmd.msoAlignBottoms;
                break;
            default:
                return;
            }

            StartNewUndoEntry();
            var flag = shapeRange.Count == 1 ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            shapeRange.Align(cmd, flag);
        }

        public void BtnDistribute_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange(mustMoreThanOrEqualTo: 3);
            if (shapeRange == null) {
                return;
            }

            Office.MsoDistributeCmd cmd;
            switch (ribbonControl.Id) {
            case btnDistributeHorizontal:
                cmd = Office.MsoDistributeCmd.msoDistributeHorizontally;
                break;
            case btnDistributeVertical:
                cmd = Office.MsoDistributeCmd.msoDistributeVertically;
                break;
            default:
                return;
            }

            StartNewUndoEntry();
            shapeRange.Distribute(cmd, Office.MsoTriState.msoFalse);
        }

        private Office.MsoScaleFrom scaleFromFlag = Office.MsoScaleFrom.msoScaleFromMiddle;

        public void BtnScalePosition_Click(Office.IRibbonControl ribbonControl) {
            if (scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle) {
                scaleFromFlag = Office.MsoScaleFrom.msoScaleFromTopLeft;
            } else {
                scaleFromFlag = Office.MsoScaleFrom.msoScaleFromMiddle;
            }
            ribbon.Invalidate();
        }

        public string GetBtnScalePositionLabel(Office.IRibbonControl ribbonControl) {
            return scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle ? ArrangeRibbonResources.btnScalePosition_Middle : ArrangeRibbonResources.btnScalePosition_TopLeft;
        }

        public System.Drawing.Image GetBtnScalePositionImage(Office.IRibbonControl ribbonControl) {
            return scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle ? Properties.Resources.ScaleFromMiddle : Properties.Resources.ScaleFromTopLeft;
        }

        public void BtnScale_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange(mustMoreThanOrEqualTo: 2);
            if (shapeRange == null) {
                return;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
            var (firstWidth, firstHeight) = (shapes[0].Width, shapes[0].Height);
            var scaleFrom = scaleFromFlag;

            StartNewUndoEntry();
            switch (ribbonControl.Id) {
            case btnScaleSameWidth:
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, scaleFrom);
                }
                break;
            case btnScaleSameHeight:
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstHeight / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, scaleFrom);
                }
                break;
            case btnScaleSameSize:
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

        public void BtnExtend_Click(Office.IRibbonControl ribbonControl) {
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
            switch (ribbonControl.Id) {
            case btnExtendSameLeft:
                foreach (var shape in shapes) {
                    var newWidth = shape.Width + shape.Left - minLeft;
                    var ratio = newWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case btnExtendSameRight:
                foreach (var shape in shapes) {
                    var newWidth = maxLeftWidth - shape.Left;
                    var ratio = newWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            case btnExtendSameTop:
                foreach (var shape in shapes) {
                    var newTop = shape.Height + shape.Top - minTop;
                    var ratio = newTop / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case btnExtendSameBottom:
                foreach (var shape in shapes) {
                    var newHeight = maxTopHeight - shape.Top;
                    var ratio = newHeight / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            }
        }

        public void BtnSnap_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange(mustMoreThanOrEqualTo: 2);
            if (shapeRange == null) {
                return;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
            var (lastLeft, lastTop) = (shapes[0].Left, shapes[0].Top);
            var (lastWidth, lastHeight) = (shapes[0].Width, shapes[0].Height);

            StartNewUndoEntry();
            switch (ribbonControl.Id) {
            case btnSnapLeft:
                for (var i = 1; i < shapes.Count(); i++) {
                    shapes[i].Left = lastLeft + lastWidth;
                    lastLeft = shapes[i].Left;
                    lastWidth = shapes[i].Width;
                }
                break;
            case btnSnapRight:
                for (var i = 1; i < shapes.Count(); i++) {
                    lastWidth = shapes[i].Width;
                    shapes[i].Left = lastLeft - lastWidth;
                    lastLeft = shapes[i].Left;
                }
                break;
            case btnSnapTop:
                for (var i = 1; i < shapes.Count(); i++) {
                    shapes[i].Top = lastTop + lastHeight;
                    lastTop = shapes[i].Top;
                    lastHeight = shapes[i].Height;
                }
                break;
            case btnSnapBottom:
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

        public void BtnMove_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            Office.MsoZOrderCmd cmd;
            switch (ribbonControl.Id) {
            case btnMoveForward:
                cmd = Office.MsoZOrderCmd.msoBringForward;
                break;
            case btnMoveBackward:
                cmd = Office.MsoZOrderCmd.msoSendBackward;
                break;
            case btnMoveFront:
                cmd = Office.MsoZOrderCmd.msoBringToFront;
                break;
            case btnMoveBack:
                cmd = Office.MsoZOrderCmd.msoSendToBack;
                break;
            default:
                return;
            }

            StartNewUndoEntry();
            shapeRange.ZOrder(cmd);
        }

        public void BtnRotate_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            StartNewUndoEntry();
            switch (ribbonControl.Id) {
            case btnRotateLeft90:
                shapeRange.IncrementRotation(-90);
                break;
            case btnRotateRight90:
                shapeRange.IncrementRotation(90);
                break;
            }
        }

        public void BtnFlip_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            StartNewUndoEntry();
            switch (ribbonControl.Id) {
            case btnFlipHorizontal:
                shapeRange.Flip(Office.MsoFlipCmd.msoFlipHorizontal);
                break;
            case btnFlipVertical:
                shapeRange.Flip(Office.MsoFlipCmd.msoFlipVertical);
                break;
            }
        }


        public void BtnGroup_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            switch (ribbonControl.Id) {
            case btnGroup:
                if (shapeRange.Count >= 2) {
                    StartNewUndoEntry();
                    var grouped = shapeRange.Group();
                    grouped.Select();
                    //AdjustButtonsAccessibility();
                }
                break;
            case btnUngroup:
                StartNewUndoEntry();
                var ungrouped = shapeRange.Ungroup();
                ungrouped.Select();
                //AdjustButtonsAccessibility();
                break;
            }
        }

        public void BtnAutofit_Click(Office.IRibbonControl ribbonControl, bool pressed) {
            var shapeRange = GetShapeRange(mustHasTextFrame: true);
            if (shapeRange == null) {
                return;
            }

            switch (ribbonControl.Id) {
            case btnAutofitOff:
                shapeRange.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
                break;
            case btnAutofitText:
                shapeRange.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
                break;
            case btnAutoResize:
                shapeRange.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
                break;
            }
            ribbon.InvalidateControl(btnAutofitOff);
            ribbon.InvalidateControl(btnAutofitText);
            ribbon.InvalidateControl(btnAutoResize);
        }

        public bool GetBtnAutofitPressed(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange(mustHasTextFrame: true);
            if (shapeRange == null) {
                return false;
            }

            switch (ribbonControl.Id) {
            case btnAutofitOff:
                return shapeRange.TextFrame2.AutoSize == Office.MsoAutoSize.msoAutoSizeNone;
            case btnAutofitText:
                return shapeRange.TextFrame2.AutoSize == Office.MsoAutoSize.msoAutoSizeTextToFitShape;
            case btnAutoResize:
                return shapeRange.TextFrame2.AutoSize == Office.MsoAutoSize.msoAutoSizeShapeToFitText;
            }
            return false;
        }

        public void CbxWrapTextbox_Click(Office.IRibbonControl ribbonControl, bool pressed) {
            var shapeRange = GetShapeRange(mustHasTextFrame: true);
            if (shapeRange == null) {
                return;
            }

            shapeRange.TextFrame.WordWrap = pressed ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            ribbon.InvalidateControl(ribbonControl.Id);
        }

        public bool GetCbxWrapTextboxChecked(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange(mustHasTextFrame: true);
            if (shapeRange == null) {
                return false;
            }

            return shapeRange.TextFrame.WordWrap == Office.MsoTriState.msoTrue;
        }

        private float CmToPt(float cm) => (float) (cm * 720 / 25.4);

        private float PtToCm(float pt) => (float) (pt * 25.4 / 720);

        private const float defaultMarginHorizontalPt = 7.2F;
        private const float defaultMarginVerticalPt = 3.6F;

        public void EdtMargin_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var shapeRange = GetShapeRange(mustHasTextFrame: true);
            if (shapeRange == null) {
                return;
            }

            text = text.Replace("cm", "").Trim();
            if (text.Length == 0) text = "0";
            if (float.TryParse(text, out float input)) {
                var pt = CmToPt(input);
                switch (ribbonControl.Id) {
                case edtMarginLeft:
                    shapeRange.TextFrame.MarginLeft = pt;
                    break;
                case edtMarginRight:
                    shapeRange.TextFrame.MarginRight = pt;
                    break;
                case edtMarginTop:
                    shapeRange.TextFrame.MarginTop = pt;
                    break;
                case edtMarginBottom:
                    shapeRange.TextFrame.MarginBottom = pt;
                    break;
                }
            }

            ribbon.InvalidateControl(ribbonControl.Id);
        }

        public string GetEdtMarginText(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange(mustHasTextFrame: true);
            if (shapeRange == null) {
                return "";
            }

            float pt = 0;
            switch (ribbonControl.Id) {
            case edtMarginLeft:
                pt = shapeRange.TextFrame.MarginLeft;
                break;
            case edtMarginRight:
                pt = shapeRange.TextFrame.MarginRight;
                break;
            case edtMarginTop:
                pt = shapeRange.TextFrame.MarginTop;
                break;
            case edtMarginBottom:
                pt = shapeRange.TextFrame.MarginBottom;
                break;
            }
            if (pt < 0) {
                return "";
            }
            return $"{Math.Round(PtToCm(pt), 2)} cm";
        }

        public void BtnResetMargin_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange(mustHasTextFrame: true);
            if (shapeRange == null) {
                return;
            }

            switch (ribbonControl.Id) {
            case btnResetMarginHorizontal:
                shapeRange.TextFrame.MarginLeft = defaultMarginHorizontalPt;
                shapeRange.TextFrame.MarginRight = defaultMarginHorizontalPt;
                break;
            case btnResetMarginVertical:
                shapeRange.TextFrame.MarginTop = defaultMarginVerticalPt;
                shapeRange.TextFrame.MarginBottom = defaultMarginVerticalPt;
                break;
            }

            ribbon.InvalidateControl(edtMarginLeft);
            ribbon.InvalidateControl(edtMarginRight);
            ribbon.InvalidateControl(edtMarginTop);
            ribbon.InvalidateControl(edtMarginBottom);
        }

    }

}
