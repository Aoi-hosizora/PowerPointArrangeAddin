using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ppt_arrange_addin {

    public partial class ArrangeRibbon {

        private Office.IRibbonUI ribbon;

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            ribbon = ribbonUI;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private struct Selection {
            public PowerPoint.ShapeRange ShapeRange { get; set; }
            public PowerPoint.Shape TextShape { get; set; }
            public PowerPoint.TextRange TextRange { get; set; }
            public PowerPoint.TextFrame TextFrame { get; set; }
            public PowerPoint.TextFrame2 TextFrame2 { get; set; }
        }

        private Selection GetSelection(bool onlyShapeRange) {
            // 1. application
            PowerPoint.Selection selection = null;
            try {
                var application = Globals.ThisAddIn.Application;
                if (application.Windows.Count > 0 && GetForegroundWindow().ToInt32() == application.HWND) {
                    selection = application.ActiveWindow.Selection;
                }
            } catch (Exception e) {
                System.Diagnostics.Debug.WriteLine(e);
            }
            if (selection == null) {
                return new Selection();
            }

            // 2. shape range
            PowerPoint.ShapeRange shapeRange = null;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) {
                shapeRange = selection.ShapeRange;
            } else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText) {
                try {
                    shapeRange = selection.ShapeRange;
                } catch (Exception) { }
            }
            if (onlyShapeRange) {
                return new Selection() { ShapeRange = shapeRange };
            }

            // 3. text range
            PowerPoint.TextRange textRange = null;
            PowerPoint.TextFrame textFrame = null;
            PowerPoint.Shape textShape = null;
            PowerPoint.TextFrame2 textFrame2 = null;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText) {
                textRange = selection.TextRange;
                if (textRange.Parent is PowerPoint.TextFrame textFrame_) {
                    textFrame = textFrame_;
                    if (textFrame.Parent is PowerPoint.Shape textShape_) {
                        textShape = textShape_;
                        textFrame2 = textShape.TextFrame2;
                    }
                }
            } else if (shapeRange != null && shapeRange.HasTextFrame != Office.MsoTriState.msoFalse) {
                textFrame = shapeRange.TextFrame;
                textRange = textFrame.TextRange;
                textFrame2 = shapeRange.TextFrame2;
                textShape = null;
            }

            // 4. return selection
            return new Selection() {
                ShapeRange = shapeRange,
                TextRange = textRange,
                TextShape = textShape,
                TextFrame = textFrame,
                TextFrame2 = textFrame2,
            };
        }

        private delegate bool AvailabilityRule(bool hasShape, int shapesCount, bool hasTextFrame);
        private Dictionary<string, AvailabilityRule> availabilityRules;

        private void InitializeAvailabilityRules() {
            availabilityRules = new Dictionary<string, AvailabilityRule>() {
                { btnAlignLeft, (_, cnt, __) => cnt >= 1 },
                { btnAlignCenter, (_, cnt, __) => cnt >= 1 },
                { btnAlignRight, (_, cnt, __) => cnt >= 1 },
                { btnAlignTop, (_, cnt, __) => cnt >= 1 },
                { btnAlignMiddle, (_, cnt, __) => cnt >= 1 },
                { btnAlignBottom, (_, cnt, __) => cnt >= 1 },
                { btnDistributeHorizontal, (_, cnt, __) => cnt >= 3 },
                { btnDistributeVertical, (_, cnt, __) => cnt >= 3 },
                { btnScaleSameWidth, (_, cnt, __) => cnt >= 2 },
                { btnScaleSameHeight, (_, cnt, __) => cnt >= 2 },
                { btnScaleSameSize, (_, cnt, __) => cnt >= 2 },
                { btnScalePosition, (_, cnt, __) => cnt >= 1 },
                { btnExtendSameLeft, (_, cnt, __) => cnt >= 2 },
                { btnExtendSameRight, (_, cnt, __) => cnt >= 2 },
                { btnExtendSameTop, (_, cnt, __) => cnt >= 2 },
                { btnExtendSameBottom, (_, cnt, __) => cnt >= 2 },
                { btnSnapLeft, (_, cnt, __) => cnt >= 2 },
                { btnSnapRight, (_, cnt, __) => cnt >= 2 },
                { btnSnapTop, (_, cnt, __) => cnt >= 2 },
                { btnSnapBottom, (_, cnt, __) => cnt >= 2 },
                { btnMoveForward, (_, cnt, __) => cnt >= 1 },
                { btnMoveFront, (_, cnt, __) => cnt >= 1 },
                { btnMoveBackward, (_, cnt, __) => cnt >= 1 },
                { btnMoveBack, (_, cnt, __) => cnt >= 1 },
                { btnRotateRight90, (_, cnt, __) => cnt >= 1 },
                { btnRotateLeft90, (_, cnt, __) => cnt >= 1 },
                { btnFlipVertical, (_, cnt, __) => cnt >= 1 },
                { btnFlipHorizontal, (_, cnt, __) => cnt >= 1 },
                { btnGroup, (_, cnt, __) => cnt >= 2 },
                { btnUngroup, (_, cnt, __) => cnt >= 1 },
                { edtShapePositionX, (_, cnt, __) => cnt >= 1 },
                { edtShapePositionY, (_, cnt, __) => cnt >= 1 },
                { btnShapePositionCopy, (_, cnt, __) => cnt == 1 },
                { btnShapePositionPaste, (_, cnt, __) => cnt >= 1 && copiedPositionXPt >= 0 && copiedPositionYPt >= 0 },
                { btnAutofitOff, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnAutofitText, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnAutoResize, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnWrapText, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginLeft, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginRight, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginTop, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginBottom, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnResetMarginHorizontal, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnResetMarginVertical, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
            };
        }

        public bool GetEnabled(Office.IRibbonControl ribbonControl) {
            var selection = GetSelection(onlyShapeRange: false);
            var hasShape = selection.ShapeRange != null;
            var shapesCount = selection.ShapeRange?.Count ?? 0;
            var hasTextFrame = selection.TextFrame != null;
            availabilityRules.TryGetValue(ribbonControl.Id, out AvailabilityRule checker);
            return checker?.Invoke(hasShape, shapesCount, hasTextFrame) ?? true;
        }

        public void AdjustRibbonButtonsAvailability(bool onlyForDrag = false) {
            if (!onlyForDrag) {
                ribbon.Invalidate();
            } else {
                // TODO
                ribbon.InvalidateControl(edtShapePositionX);
                ribbon.InvalidateControl(edtShapePositionY);
            }
        }

        private PowerPoint.ShapeRange GetShapeRange(int mustMoreThanOrEqualTo = 1) {
            var selection = GetSelection(onlyShapeRange: true);
            var shapeRange = selection.ShapeRange;
            if (shapeRange == null || shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            return shapeRange;
        }

        private (PowerPoint.TextFrame, PowerPoint.TextFrame2) GetTextFrame() {
            var selection = GetSelection(onlyShapeRange: false);
            return (selection.TextFrame, selection.TextFrame2);
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

        public void BtnScalePosition_Click(Office.IRibbonControl _) {
            if (scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle) {
                scaleFromFlag = Office.MsoScaleFrom.msoScaleFromTopLeft;
            } else {
                scaleFromFlag = Office.MsoScaleFrom.msoScaleFromMiddle;
            }
            ribbon.Invalidate();
        }

        public string GetBtnScalePositionLabel(Office.IRibbonControl _) {
            return scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle ? ArrangeRibbonResources.btnScalePosition_Middle : ArrangeRibbonResources.btnScalePosition_TopLeft;
        }

        public System.Drawing.Image GetBtnScalePositionImage(Office.IRibbonControl _) {
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
                    AdjustRibbonButtonsAvailability();
                }
                break;
            case btnUngroup:
                if (shapeRange.OfType<PowerPoint.Shape>().Any((s) => s.Type == Office.MsoShapeType.msoGroup)) {
                    StartNewUndoEntry();
                    var ungrouped = shapeRange.Ungroup();
                    ungrouped.Select();
                    AdjustRibbonButtonsAvailability();
                }
                break;
            }
        }

        private float CmToPt(float cm) => (float) (cm * 720 / 25.4);

        private float PtToCm(float pt) => (float) (pt * 25.4 / 720);

        public void EdtShapePosition_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            text = text.Replace("cm", "").Trim();
            if (text.Length == 0) text = "0";

            StartNewUndoEntry();
            if (float.TryParse(text, out float input)) {
                var pt = CmToPt(input);
                switch (ribbonControl.Id) {
                case edtShapePositionX:
                    shapeRange.Left = pt;
                    break;
                case edtShapePositionY:
                    shapeRange.Top = pt;
                    break;
                }
            }

            ribbon.InvalidateControl(ribbonControl.Id);
        }

        public string GetEdtShapePositionText(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return "";
            }

            // TODO call this callback when drag the shape to change it's position
            float pt = 0;
            switch (ribbonControl.Id) {
            case edtShapePositionX:
                pt = shapeRange.Left;
                break;
            case edtShapePositionY:
                pt = shapeRange.Top;
                break;
            }
            if (pt < 0) {
                return "";
            }
            return $"{Math.Round(PtToCm(pt), 2)} cm";
        }

        private float copiedPositionXPt = -1; // for shape and image
        private float copiedPositionYPt = -1; // for shape and image

        public void BtnShapePositionCopy_Click(Office.IRibbonControl _) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null || shapeRange.Count > 1) {
                return;
            }

            copiedPositionXPt = shapeRange.Left;
            copiedPositionYPt = shapeRange.Top;
            ribbon.InvalidateControl(btnShapePositionPaste);
        }

        public void BtnShapePositionPaste_Click(Office.IRibbonControl _) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            if (copiedPositionXPt >= 0 && copiedPositionYPt >= 0) {
                shapeRange.Left = copiedPositionXPt;
                shapeRange.Top = copiedPositionYPt;
            }
            ribbon.InvalidateControl(edtShapePositionX);
            ribbon.InvalidateControl(edtShapePositionY);
        }

        public void BtnAutofit_Click(Office.IRibbonControl ribbonControl, bool _) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return;
            }

            StartNewUndoEntry();
            switch (ribbonControl.Id) {
            case btnAutofitOff:
                textFrame.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
                break;
            case btnAutofitText:
                textFrame.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
                break;
            case btnAutoResize:
                textFrame.AutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
                break;
            }
            ribbon.InvalidateControl(btnAutofitOff);
            ribbon.InvalidateControl(btnAutofitText);
            ribbon.InvalidateControl(btnAutoResize);
        }

        public bool GetBtnAutofitPressed(Office.IRibbonControl ribbonControl) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return false;
            }

            switch (ribbonControl.Id) {
            case btnAutofitOff:
                return textFrame.AutoSize == Office.MsoAutoSize.msoAutoSizeNone;
            case btnAutofitText:
                return textFrame.AutoSize == Office.MsoAutoSize.msoAutoSizeTextToFitShape;
            case btnAutoResize:
                return textFrame.AutoSize == Office.MsoAutoSize.msoAutoSizeShapeToFitText;
            }
            return false;
        }

        public void BtnWrapText_Click(Office.IRibbonControl ribbonControl, bool _) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return;
            }

            StartNewUndoEntry();
            if (textFrame.WordWrap != Office.MsoTriState.msoTrue) {
                textFrame.WordWrap = Office.MsoTriState.msoTrue;
            } else {
                textFrame.WordWrap = Office.MsoTriState.msoFalse;
            }
            ribbon.InvalidateControl(ribbonControl.Id);
        }

        public bool GetBtnWrapTextPressed(Office.IRibbonControl _) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return false;
            }

            return textFrame.WordWrap == Office.MsoTriState.msoTrue;
        }

        private const float defaultMarginHorizontalPt = 7.2F;
        private const float defaultMarginVerticalPt = 3.6F;

        public void EdtMargin_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return;
            }

            text = text.Replace("cm", "").Trim();
            if (text.Length == 0) text = "0";

            StartNewUndoEntry();
            if (float.TryParse(text, out float input)) {
                var pt = CmToPt(input);
                switch (ribbonControl.Id) {
                case edtMarginLeft:
                    textFrame.MarginLeft = pt;
                    break;
                case edtMarginRight:
                    textFrame.MarginRight = pt;
                    break;
                case edtMarginTop:
                    textFrame.MarginTop = pt;
                    break;
                case edtMarginBottom:
                    textFrame.MarginBottom = pt;
                    break;
                }
            }

            ribbon.InvalidateControl(ribbonControl.Id);
        }

        public string GetEdtMarginText(Office.IRibbonControl ribbonControl) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return "";
            }

            float pt = 0;
            switch (ribbonControl.Id) {
            case edtMarginLeft:
                pt = textFrame.MarginLeft;
                break;
            case edtMarginRight:
                pt = textFrame.MarginRight;
                break;
            case edtMarginTop:
                pt = textFrame.MarginTop;
                break;
            case edtMarginBottom:
                pt = textFrame.MarginBottom;
                break;
            }
            if (pt < 0) {
                return "";
            }
            return $"{Math.Round(PtToCm(pt), 2)} cm";
        }

        public void BtnResetMargin_Click(Office.IRibbonControl ribbonControl) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return;
            }

            StartNewUndoEntry();
            switch (ribbonControl.Id) {
            case btnResetMarginHorizontal:
                textFrame.MarginLeft = defaultMarginHorizontalPt;
                textFrame.MarginRight = defaultMarginHorizontalPt;
                break;
            case btnResetMarginVertical:
                textFrame.MarginTop = defaultMarginVerticalPt;
                textFrame.MarginBottom = defaultMarginVerticalPt;
                break;
            }

            ribbon.InvalidateControl(edtMarginLeft);
            ribbon.InvalidateControl(edtMarginRight);
            ribbon.InvalidateControl(edtMarginTop);
            ribbon.InvalidateControl(edtMarginBottom);
        }

    }

}
