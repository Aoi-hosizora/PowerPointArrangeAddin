using System;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

#nullable enable

namespace ppt_arrange_addin.Helper {

    public static class TextboxHelper {

        public enum TextboxStatusCmd {
            AutofitOff,
            AutoShrinkText,
            AutoResizeShape,
            WrapTextOnOff,
        }

        public static void ChangeAutofitStatus(PowerPoint.TextFrame2? textFrame, TextboxStatusCmd? cmd, Action? uiInvalidator = null) {
            if (textFrame == null) {
                return;
            }
            if (cmd == null) {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (cmd!) {
            case TextboxStatusCmd.AutofitOff:
                textFrame.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
                break;
            case TextboxStatusCmd.AutoShrinkText:
                textFrame.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
                break;
            case TextboxStatusCmd.AutoResizeShape:
                textFrame.AutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
                break;
            case TextboxStatusCmd.WrapTextOnOff:
                textFrame.WordWrap = textFrame.WordWrap != Office.MsoTriState.msoTrue
                    ? Office.MsoTriState.msoTrue
                    : Office.MsoTriState.msoFalse;
                break;
            }
            uiInvalidator?.Invoke();
        }

        public static bool GetAutofitStatus(PowerPoint.TextFrame2? textFrame, TextboxStatusCmd? cmd) {
            if (textFrame == null) {
                return false;
            }
            if (cmd == null) {
                return false;
            }

            return cmd! switch {
                TextboxStatusCmd.AutofitOff => textFrame.AutoSize == Office.MsoAutoSize.msoAutoSizeNone,
                TextboxStatusCmd.AutoShrinkText => textFrame.AutoSize == Office.MsoAutoSize.msoAutoSizeTextToFitShape,
                TextboxStatusCmd.AutoResizeShape => textFrame.AutoSize == Office.MsoAutoSize.msoAutoSizeShapeToFitText,
                TextboxStatusCmd.WrapTextOnOff => textFrame.WordWrap == Office.MsoTriState.msoTrue,
                _ => false
            };
        }

        public enum MarginKind {
            Left,
            Right,
            Top,
            Bottom,
        }

        public static void ChangeMarginOfString(PowerPoint.TextFrame? textFrame, MarginKind? marginKind, string? input, Action? uiInvalidator = null) {
            if (textFrame == null) {
                return;
            }
            if (marginKind == null || input == null) {
                return;
            }
            var (valueInPt, ok) = UnitConverter.ParseStringToPtValue(input);
            if (!ok) {
                uiInvalidator?.Invoke(); // reset input
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (marginKind) {
            case MarginKind.Left:
                textFrame.MarginLeft = valueInPt;
                break;
            case MarginKind.Right:
                textFrame.MarginRight = valueInPt;
                break;
            case MarginKind.Top:
                textFrame.MarginTop = valueInPt;
                break;
            case MarginKind.Bottom:
                textFrame.MarginBottom = valueInPt;
                break;
            }
            uiInvalidator?.Invoke();
        }

        public static (string, bool) GetMarginOfString(PowerPoint.TextFrame? textFrame, MarginKind? marginKind) {
            if (textFrame == null) {
                return ("", false);
            }
            if (marginKind == null) {
                return ("", false);
            }

            var valueInPt = marginKind! switch {
                MarginKind.Left => textFrame.MarginLeft,
                MarginKind.Right => textFrame.MarginRight,
                MarginKind.Top => textFrame.MarginTop,
                MarginKind.Bottom => textFrame.MarginBottom,
                _ => -1
            };

            var text = "";
            if (valueInPt >= 0) {
                text = UnitConverter.FormatPtValueToString(valueInPt);
            }
            return (text, true);
        }

        public enum ResetMarginCmd {
            Horizontal,
            Vertical,
        }

        private static readonly float DefaultMarginHorizontalPt = 7.2F;
        private static readonly float DefaultMarginVerticalPt = 3.6F;

        public static void ResetMargin(PowerPoint.TextFrame? textFrame, ResetMarginCmd? cmd, Action? uiInvalidator = null) {
            if (textFrame == null) {
                return;
            }
            if (cmd == null) {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (cmd!) {
            case ResetMarginCmd.Horizontal:
                textFrame.MarginLeft = DefaultMarginHorizontalPt;
                textFrame.MarginRight = DefaultMarginHorizontalPt;
                break;
            case ResetMarginCmd.Vertical:
                textFrame.MarginTop = DefaultMarginVerticalPt;
                textFrame.MarginBottom = DefaultMarginVerticalPt;
                break;
            }
            uiInvalidator?.Invoke();
        }

    }

}
