using System;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

#nullable enable

namespace ppt_arrange_addin.Helper {

    public static class ArrangementHelper {

        public static void Align(PowerPoint.ShapeRange? shapeRange, Office.MsoAlignCmd? cmd, Office.MsoTriState? relativeToSlide = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (cmd == null) {
                return;
            }

            if (relativeToSlide == null) {
                relativeToSlide = Office.MsoTriState.msoFalse; // defaults to relative to objects
                if (shapeRange.Count == 1) {
                    relativeToSlide = Office.MsoTriState.msoTrue; // relative to slide when only single shape is selected
                }
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            shapeRange.Align(cmd.Value, relativeToSlide.Value);
        }

        public static void Distribute(PowerPoint.ShapeRange? shapeRange, Office.MsoDistributeCmd? cmd, Office.MsoTriState? relativeToSlide = null) {
            if (shapeRange == null || shapeRange.Count <= 0 || shapeRange.Count == 2) {
                return;
            }
            if (cmd == null) {
                return;
            }

            if (relativeToSlide == null) {
                relativeToSlide = Office.MsoTriState.msoFalse; // defaults to relative to objects
                if (shapeRange.Count == 1) {
                    relativeToSlide = Office.MsoTriState.msoTrue; // relative to slide when only single shape is selected
                }
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            shapeRange.Distribute(cmd.Value, relativeToSlide.Value);
        }

        public enum ScaleSizeCmd {
            SameWidth,
            SameHeight,
            SameSize
        }

        public static void ScaleSize(PowerPoint.ShapeRange? shapeRange, ScaleSizeCmd? cmd, Office.MsoScaleFrom scaleFromFlag) {
            if (shapeRange == null || shapeRange.Count < 2) {
                return;
            }
            if (cmd == null) {
                return;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
            var (firstWidth, firstHeight) = (shapes[0].Width, shapes[0].Height); // select the first shape as final size

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (cmd!) {
            case ScaleSizeCmd.SameWidth:
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, scaleFromFlag);
                }
                break;
            case ScaleSizeCmd.SameHeight:
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstHeight / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, scaleFromFlag);
                }
                break;
            case ScaleSizeCmd.SameSize:
                for (var i = 1; i < shapes.Length; i++) {
                    var shape = shapes[i];
                    var ratio = firstWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, scaleFromFlag);
                    ratio = firstHeight / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, scaleFromFlag);
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
                    var ratio = newWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case ExtendSizeCmd.ExtendToRight:
                foreach (var shape in shapes) {
                    var newWidth = maxLeftWidth - shape.Left;
                    var ratio = newWidth / shape.Width;
                    shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            case ExtendSizeCmd.ExtendToTop:
                foreach (var shape in shapes) {
                    var newTop = shape.Height + shape.Top - minTop;
                    var ratio = newTop / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight);
                }
                break;
            case ExtendSizeCmd.ExtendToBottom:
                foreach (var shape in shapes) {
                    var newHeight = maxTopHeight - shape.Top;
                    var ratio = newHeight / shape.Height;
                    shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft);
                }
                break;
            }
        }

        public enum SnapCmd {
            SnapToLeft,
            SnapToRight,
            SnapToTop,
            SnapToBottom
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
            case SnapCmd.SnapToLeft:
                for (var i = 1; i < shapes.Length; i++) {
                    shapes[i].Left = previousLeft + previousWidth;
                    previousLeft = shapes[i].Left;
                    previousWidth = shapes[i].Width;
                }
                break;
            case SnapCmd.SnapToRight:
                for (var i = 1; i < shapes.Length; i++) {
                    previousWidth = shapes[i].Width;
                    shapes[i].Left = previousLeft - previousWidth;
                    previousLeft = shapes[i].Left;
                }
                break;
            case SnapCmd.SnapToTop:
                for (var i = 1; i < shapes.Length; i++) {
                    shapes[i].Top = previousTop + previousHeight;
                    previousTop = shapes[i].Top;
                    previousHeight = shapes[i].Height;
                }
                break;
            case SnapCmd.SnapToBottom:
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

        public static void Group(PowerPoint.ShapeRange? shapeRange, GroupCmd? cmd, Action? uiInvalidator) {
            if (shapeRange == null) {
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

    }

}
