using System;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

#nullable enable

namespace PowerPointArrangeAddin.Helper {

    public static class ArrangementHelper {

        public enum AlignRelativeFlag {
            RelativeToObjects,
            RelativeToFirstObject,
            RelativeToSlide
        }

        public static void Align(PowerPoint.ShapeRange? shapeRange, Office.MsoAlignCmd? cmd, AlignRelativeFlag? relativeFlag = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (cmd == null) {
                return;
            }

            relativeFlag ??= AlignRelativeFlag.RelativeToObjects; // defaults to relative to objects
            if (shapeRange.Count == 1) {
                relativeFlag = AlignRelativeFlag.RelativeToSlide; // relative to slide when only single shape is selected
            }

            switch (relativeFlag!) {
            case AlignRelativeFlag.RelativeToObjects:
                if (shapeRange.Count >= 2) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    shapeRange.Align(cmd.Value, Office.MsoTriState.msoFalse);
                }
                break;
            case AlignRelativeFlag.RelativeToSlide:
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                shapeRange.Align(cmd.Value, Office.MsoTriState.msoTrue);
                break;
            case AlignRelativeFlag.RelativeToFirstObject:
                if (shapeRange.Count >= 2) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
                    var firstShape = shapes[0];
                    for (var i = 1; i < shapes.Length; i++) {
                        shapes[i].Left = cmd! switch {
                            Office.MsoAlignCmd.msoAlignLefts => firstShape.Left,
                            Office.MsoAlignCmd.msoAlignCenters => firstShape.Left + (firstShape.Width - shapes[i].Width) / 2,
                            Office.MsoAlignCmd.msoAlignRights => firstShape.Left + firstShape.Width - shapes[i].Width,
                            _ => shapes[i].Left
                        };
                        shapes[i].Top = cmd! switch {
                            Office.MsoAlignCmd.msoAlignTops => firstShape.Top,
                            Office.MsoAlignCmd.msoAlignMiddles => firstShape.Top + (firstShape.Height - shapes[i].Height) / 2,
                            Office.MsoAlignCmd.msoAlignBottoms => firstShape.Top + firstShape.Height - shapes[i].Height,
                            _ => shapes[i].Top
                        };
                    }
                }
                break;
            }
        }

        public static bool IsDistributable(PowerPoint.ShapeRange? shapeRange, AlignRelativeFlag? relativeCmd) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return false;
            }
            if (shapeRange.Count == 1 || relativeCmd == AlignRelativeFlag.RelativeToSlide) {
                return true; // ignore relative cmd, which is regarded with relative to slide
            }
            if (relativeCmd == null || relativeCmd == AlignRelativeFlag.RelativeToFirstObject) {
                return false; // always disable distribution, when relative to the first object
            }
            return shapeRange.Count >= 3; // for "relative to shapes" when select more than 3 objects
        }

        public static void Distribute(PowerPoint.ShapeRange? shapeRange, Office.MsoDistributeCmd? cmd, AlignRelativeFlag? relativeFlag = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (cmd == null) {
                return;
            }

            relativeFlag ??= AlignRelativeFlag.RelativeToObjects; // defaults to relative to objects
            if (shapeRange.Count == 1) {
                relativeFlag = AlignRelativeFlag.RelativeToSlide; // relative to slide when only single shape is selected
            }

            switch (relativeFlag!) {
            case AlignRelativeFlag.RelativeToObjects:
                if (shapeRange.Count >= 3) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    shapeRange.Distribute(cmd.Value, Office.MsoTriState.msoFalse);
                }
                break;
            case AlignRelativeFlag.RelativeToSlide:
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                shapeRange.Distribute(cmd.Value, Office.MsoTriState.msoTrue);
                break;
            case AlignRelativeFlag.RelativeToFirstObject:
                break;
            }
        }

        public static void UpdateAppAlignRelative(AlignRelativeFlag flag) {
            var mso = flag == AlignRelativeFlag.RelativeToSlide
                ? "ObjectsAlignRelativeToContainerSmart"
                : "ObjectsAlignSelectedSmart";
            try {
                if (Globals.ThisAddIn.Application.CommandBars.GetEnabledMso(mso)) {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso(mso);
                }
            } catch (Exception) {
                // ignored
            }
        }

        public enum ScaleSizeCmd {
            SameWidth,
            SameHeight,
            SameSize
        }

        public static void ScaleSize(PowerPoint.ShapeRange? shapeRange, ScaleSizeCmd? cmd, Office.MsoScaleFrom? scaleFromFlag) {
            if (shapeRange == null || shapeRange.Count < 2) {
                return;
            }
            if (cmd == null) {
                return;
            }

            scaleFromFlag ??= Office.MsoScaleFrom.msoScaleFromTopLeft;

            var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
            var (firstWidth, firstHeight) = (shapes[0].Width, shapes[0].Height); // select the first shape as final size

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (cmd!) {
            case ScaleSizeCmd.SameWidth:
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    shape.ScaleWidthTo(firstWidth, scaleFromFlag.Value);
                }
                break;
            case ScaleSizeCmd.SameHeight:
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    shape.ScaleHeightTo(firstHeight, scaleFromFlag.Value);
                }
                break;
            case ScaleSizeCmd.SameSize:
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    shape.ScaleSizeTo(firstWidth, firstHeight, scaleFromFlag.Value);
                }
                break;
            }

        }

        public enum ExtendSizeCmd {
            ExtendToLeft,
            ExtendToRight,
            ExtendToTop,
            ExtendToBottom
        }

        public static void ExtendSize(PowerPoint.ShapeRange? shapeRange, ExtendSizeCmd? cmd) {
            if (shapeRange == null || shapeRange.Count < 2) {
                return;
            }
            if (cmd == null) {
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

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (cmd!) {
            case ExtendSizeCmd.ExtendToLeft:
                foreach (var shape in shapes) {
                    var newWidth = shape.Width + shape.Left - minLeft;
                    shape.ScaleWidthTo(newWidth, Office.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case ExtendSizeCmd.ExtendToRight:
                foreach (var shape in shapes) {
                    var newWidth = maxLeftWidth - shape.Left;
                    shape.ScaleWidthTo(newWidth, Office.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            case ExtendSizeCmd.ExtendToTop:
                foreach (var shape in shapes) {
                    var newHeight = shape.Height + shape.Top - minTop;
                    shape.ScaleHeightTo(newHeight, Office.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case ExtendSizeCmd.ExtendToBottom:
                foreach (var shape in shapes) {
                    var newHeight = maxTopHeight - shape.Top;
                    shape.ScaleHeightTo(newHeight, Office.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            }
        }

        public enum SnapCmd {
            SnapLeftToRight,
            SnapRightToLeft,
            SnapTopToBottom,
            SnapBottomToTop
        }

        public static void Snap(PowerPoint.ShapeRange? shapeRange, SnapCmd? cmd) {
            if (shapeRange == null || shapeRange.Count < 2) {
                return;
            }
            if (cmd == null) {
                return;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
            var (previousLeft, previousTop) = (shapes[0].Left, shapes[0].Top);
            var (previousWidth, previousHeight) = (shapes[0].Width, shapes[0].Height);

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (cmd!) {
            case SnapCmd.SnapLeftToRight:
                for (var i = 1; i < shapes.Length; i++) {
                    shapes[i].Left = previousLeft + previousWidth;
                    previousLeft = shapes[i].Left;
                    previousWidth = shapes[i].Width;
                }
                break;
            case SnapCmd.SnapRightToLeft:
                for (var i = 1; i < shapes.Length; i++) {
                    previousWidth = shapes[i].Width;
                    shapes[i].Left = previousLeft - previousWidth;
                    previousLeft = shapes[i].Left;
                }
                break;
            case SnapCmd.SnapTopToBottom:
                for (var i = 1; i < shapes.Length; i++) {
                    shapes[i].Top = previousTop + previousHeight;
                    previousTop = shapes[i].Top;
                    previousHeight = shapes[i].Height;
                }
                break;
            case SnapCmd.SnapBottomToTop:
                for (var i = 1; i < shapes.Length; i++) {
                    previousHeight = shapes[i].Height;
                    shapes[i].Top = previousTop - previousHeight;
                    previousTop = shapes[i].Top;
                }
                break;
            }
        }

        public static void LayerMove(PowerPoint.ShapeRange? shapeRange, Office.MsoZOrderCmd? cmd) {
            if (shapeRange == null) {
                return;
            }
            if (cmd == null) {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            shapeRange.ZOrder(cmd.Value);
        }

        public enum RotateCmd {
            RotateLeft90,
            RotateRight90
        }

        public static void Rotate(PowerPoint.ShapeRange? shapeRange, RotateCmd? cmd) {
            if (shapeRange == null) {
                return;
            }
            if (cmd == null) {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (cmd!) {
            case RotateCmd.RotateLeft90:
                shapeRange.IncrementRotation(-90);
                break;
            case RotateCmd.RotateRight90:
                shapeRange.IncrementRotation(90);
                break;
            }
        }

        public static void Flip(PowerPoint.ShapeRange? shapeRange, Office.MsoFlipCmd? cmd) {
            if (shapeRange == null) {
                return;
            }
            if (cmd == null) {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            shapeRange.Flip(cmd.Value);
        }

        public enum GroupCmd {
            Group,
            Ungroup
        }


        public static bool IsUngroupable(PowerPoint.ShapeRange? shapeRange) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return false;
            }
            return shapeRange.OfType<PowerPoint.Shape>().Any(s => s.Type == Office.MsoShapeType.msoGroup);
        }

        public static void Group(PowerPoint.ShapeRange? shapeRange, GroupCmd? cmd, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (cmd == null) {
                return;
            }

            switch (cmd!) {
            case GroupCmd.Group:
                if (shapeRange.Count >= 2) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    var grouped = shapeRange.Group();
                    grouped.Select();
                    uiInvalidator?.Invoke();
                }
                break;
            case GroupCmd.Ungroup:
                if (IsUngroupable(shapeRange)) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    var ungrouped = shapeRange.Ungroup();
                    ungrouped.Select();
                    uiInvalidator?.Invoke();
                }
                break;
            }
        }

        public static void GridSettingDialog() {
            const string mso = "GridSettings";
            try {
                if (Globals.ThisAddIn.Application.CommandBars.GetEnabledMso(mso)) {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso(mso);
                }
            } catch (Exception) {
                // ignored
            }
        }

    }

}
