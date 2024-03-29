﻿using System;
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
            if (relativeCmd == null) {
                return false;
            }
            if (shapeRange.Count == 1 || relativeCmd == AlignRelativeFlag.RelativeToSlide) {
                return true; // ignore relative cmd, which is regarded with relative to slide
            }
            return shapeRange.Count >= 3; // for "relative to shapes" and "relative to the first object", when select more than 3 objects
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
                var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
                if (shapes.Length <= 2) {
                    break;
                }
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                var (firstShape, secondShape) = (shapes[0], shapes[1]);
                if (cmd.Value == Office.MsoDistributeCmd.msoDistributeHorizontally) {
                    var horizontalDistance = secondShape.Left - firstShape.Left;
                    if (horizontalDistance == 0) {
                        for (var i = 2; i < shapes.Length; i++) {
                            shapes[i].Left = firstShape.Left; // same left
                        }
                        break;
                    }
                    var sortedShapes = shapes.Skip(2).ToArray();
                    if (horizontalDistance > 0) {
                        // first -> second -> ...
                        Array.Sort(sortedShapes, (a, b) => a.Left.CompareTo(b.Left));
                        var separated = false;
                        if (horizontalDistance >= firstShape.Width) {
                            horizontalDistance -= firstShape.Width;
                            separated = true; // shapes are seperated
                        }
                        var left = secondShape.Left + horizontalDistance + (separated ? secondShape.Width : 0);
                        foreach (var shape in sortedShapes) {
                            shape.Left = left;
                            left += horizontalDistance + (separated ? shape.Width : 0);
                        }
                    } else if (horizontalDistance < 0) {
                        // ... <- second <- first
                        Array.Sort(sortedShapes, (a, b) => b.Left.CompareTo(a.Left));
                        horizontalDistance = -horizontalDistance;
                        var separated = false;
                        if (horizontalDistance >= secondShape.Width) {
                            horizontalDistance -= secondShape.Width;
                            separated = true; // shapes are seperated
                        }
                        var left = secondShape.Left - horizontalDistance;
                        foreach (var shape in sortedShapes) {
                            shape.Left = left - (separated ? shape.Width : 0);
                            left -= horizontalDistance + (separated ? shape.Width : 0);
                        }
                    }
                } else if (cmd.Value == Office.MsoDistributeCmd.msoDistributeVertically) {
                    var verticalDistance = secondShape.Top - firstShape.Top;
                    if (verticalDistance == 0) {
                        for (var i = 2; i < shapes.Length; i++) {
                            shapes[i].Top = firstShape.Top; // same top
                        }
                        break;
                    }
                    var sortedShapes = shapes.Skip(2).ToArray();
                    if (verticalDistance > 0) {
                        // first -> second -> ...
                        Array.Sort(sortedShapes, (a, b) => a.Top.CompareTo(b.Top));
                        var separated = false;
                        if (verticalDistance >= firstShape.Top) {
                            verticalDistance -= firstShape.Top;
                            separated = true; // shapes are seperated 
                        }
                        var top = secondShape.Top + verticalDistance + (separated ? secondShape.Height : 0);
                        foreach (var shape in sortedShapes) {
                            shape.Top = top;
                            top += verticalDistance + (separated ? shape.Top : 0);
                        }
                    } else if (verticalDistance < 0) {
                        // ... <- second <- first
                        Array.Sort(sortedShapes, (a, b) => b.Top.CompareTo(a.Top));
                        verticalDistance = -verticalDistance;
                        var separated = false;
                        if (verticalDistance >= secondShape.Top) {
                            verticalDistance -= secondShape.Top;
                            separated = true; // shapes are seperated 
                        }
                        var top = secondShape.Top - verticalDistance;
                        foreach (var shape in sortedShapes) {
                            shape.Top = top - (separated ? shape.Height : 0);
                            top -= verticalDistance + (separated ? shape.Height : 0);
                        }
                    }
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

        public static void ScaleSize(PowerPoint.ShapeRange? shapeRange, ScaleSizeCmd? cmd, SizeAndPositionHelper.ScaleFromFlag? scaleFromFlag) {
            if (shapeRange == null || shapeRange.Count < 2) {
                return;
            }
            if (cmd == null) {
                return;
            }

            scaleFromFlag ??= SizeAndPositionHelper.ScaleFromFlag.FromTopLeft;

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

        public static void ExtendSize(PowerPoint.ShapeRange? shapeRange, ExtendSizeCmd? cmd, bool extendToFirstObject) {
            if (shapeRange == null || shapeRange.Count < 2) {
                return;
            }
            if (cmd == null) {
                return;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
            float minLeft = 0x7fffffff, minTop = 0x7fffffff, maxLeftWidth = -1, maxTopHeight = -1;
            if (!extendToFirstObject) {
                foreach (var shape in shapes) {
                    minLeft = Math.Min(minLeft, shape.Left);
                    minTop = Math.Min(minTop, shape.Top);
                    maxLeftWidth = Math.Max(maxLeftWidth, shape.Left + shape.Width);
                    maxTopHeight = Math.Max(maxTopHeight, shape.Top + shape.Height);
                }
            } else {
                minLeft = shapes[0].Left;
                minTop = shapes[0].Top;
                maxLeftWidth = shapes[0].Left + shapes[0].Width;
                maxTopHeight = shapes[0].Top + shapes[0].Height;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (cmd!) {
            case ExtendSizeCmd.ExtendToLeft:
                foreach (var shape in shapes) {
                    var newWidth = shape.Width + shape.Left - minLeft;
                    shape.ScaleWidthTo(newWidth, SizeAndPositionHelper.ScaleFromFlag.FromBottomRight);
                }
                break;
            case ExtendSizeCmd.ExtendToRight:
                foreach (var shape in shapes) {
                    var newWidth = maxLeftWidth - shape.Left;
                    shape.ScaleWidthTo(newWidth, SizeAndPositionHelper.ScaleFromFlag.FromTopLeft);
                }
                break;
            case ExtendSizeCmd.ExtendToTop:
                foreach (var shape in shapes) {
                    var newHeight = shape.Height + shape.Top - minTop;
                    shape.ScaleHeightTo(newHeight, SizeAndPositionHelper.ScaleFromFlag.FromBottomRight);
                }
                break;
            case ExtendSizeCmd.ExtendToBottom:
                foreach (var shape in shapes) {
                    var newHeight = maxTopHeight - shape.Top;
                    shape.ScaleHeightTo(newHeight, SizeAndPositionHelper.ScaleFromFlag.FromTopLeft);
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
