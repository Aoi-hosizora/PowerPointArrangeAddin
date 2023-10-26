using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

#nullable enable

namespace PowerPointArrangeAddin.Helper {
    public static class SelectionGetter {

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public struct Selection {
            public PowerPoint.ShapeRange? ShapeRange { get; init; }
            public PowerPoint.Shape? TextShape { get; init; }
            public PowerPoint.TextRange? TextRange { get; init; }
            public PowerPoint.TextFrame? TextFrame { get; init; }
            public PowerPoint.TextFrame2? TextFrame2 { get; init; }
        }

        public static Selection GetSelection(bool onlyShapeRange) {
            // 1. application
            PowerPoint.Selection? selection = null;
            try {
                var application = Globals.ThisAddIn.Application;
                if (application.Windows.Count > 0) {
                    // GetForegroundWindow().ToInt32() == application.HWND
                    selection = application.ActiveWindow.Selection;
                }
            } catch (Exception) { /* ignored */
            }
            if (selection == null) {
                return new Selection();
            }

            // 2. shape range
            PowerPoint.ShapeRange? shapeRange = null;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) {
                shapeRange = selection.ShapeRange;
            } else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText) {
                try {
                    shapeRange = selection.ShapeRange;
                } catch (Exception) { /* ignored */
                }
            }
            if (onlyShapeRange) {
                return new Selection { ShapeRange = shapeRange };
            }

            // 3. text range
            PowerPoint.TextRange? textRange = null;
            PowerPoint.TextFrame? textFrame = null;
            PowerPoint.Shape? textShape = null;
            PowerPoint.TextFrame2? textFrame2 = null;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText) {
                textRange = selection.TextRange;
                if (textRange.Parent is PowerPoint.TextFrame frame) {
                    textFrame = frame;
                    if (textFrame.Parent is PowerPoint.Shape shape) {
                        textShape = shape;
                        textFrame2 = shape.TextFrame2;
                    }
                }
            } else if (shapeRange != null && shapeRange.HasTextFrame != Office.MsoTriState.msoFalse) {
                textFrame = shapeRange.TextFrame;
                textFrame2 = shapeRange.TextFrame2;
                try {
                    textRange = textFrame.TextRange; // may throw when selecting different type objects
                } catch (Exception) { }
            }

            // 4. return selection
            return new Selection {
                ShapeRange = shapeRange,
                TextRange = textRange,
                TextShape = textShape,
                TextFrame = textFrame,
                TextFrame2 = textFrame2
            };
        }

    }

}
