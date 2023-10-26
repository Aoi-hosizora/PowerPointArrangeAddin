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

        public static void InvalidateControl(this IRibbonUI ribbonUi, string controlId, string parentId) {
            ribbonUi.InvalidateControl($"{controlId}{Separator}{parentId}");
        }

        public static void InvalidateControl(this IRibbonUI ribbonUi, string controlId, (string, string) parentIds) {
            ribbonUi.InvalidateControl($"{controlId}{Separator}{parentIds.Item1}{Separator}{parentIds.Item2}");
        }

        public static void InvalidateControls(this IRibbonUI ribbonUi, string controlId, IEnumerable<string> parentIds) {
            foreach (var parentId in parentIds) {
                ribbonUi.InvalidateControl(controlId, parentId);
            }
        }

        public static void InvalidateControls(this IRibbonUI ribbonUi, string controlId, IEnumerable<(string, string)> parentIds) {
            foreach (var (parentId1, parentId2) in parentIds) {
                ribbonUi.InvalidateControl(controlId, (parentId1, parentId2));
            }
        }

    }

}
