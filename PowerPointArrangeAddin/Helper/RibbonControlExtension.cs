using System;
using System.Collections.Generic;
using Microsoft.Office.Core;

#pragma warning disable CS0618

namespace PowerPointArrangeAddin.Helper {

    public static class RibbonControlExtension {

        public static string TagOrId(this IRibbonControl ribbonControl) {
            // Microsoft.Office.Core.IRibbonControl;
            return !string.IsNullOrWhiteSpace(ribbonControl.Tag)
                ? ribbonControl.Tag
                : ribbonControl.Id;
        }

        private const string Separator = "·";

        public static string Id(this IRibbonControl ribbonControl) {
            var id = ribbonControl.Id;
            var parts = id.Split(new[] { Separator }, 2, StringSplitOptions.RemoveEmptyEntries);
            return parts[0];
        }

        public static string Group(this IRibbonControl ribbonControl) {
            var id = ribbonControl.Id;
            var parts = id.Split(new[] { Separator }, 2, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length < 2 ? "" : parts[1];
        }

        public static void InvalidateControl(this IRibbonUI ribbonUi, string controlId, string parentName) {
            ribbonUi.InvalidateControl($"{controlId}{Separator}{parentName}");
        }

        public static void InvalidateControl(this IRibbonUI ribbonUi, string controlId, IEnumerable<string> parentNames) {
            foreach (var parentName in parentNames) {
                ribbonUi.InvalidateControl(controlId, parentName);
            }
        }

    }

}
