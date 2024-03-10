using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Forms = System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

#nullable enable

namespace PowerPointArrangeAddin.Helper {

    public static class ReplacePictureHelper {

        public enum ReplacePictureCmd {
            WithClipboard,
            WithFile
        }

        [Flags]
        public enum ReplacePictureFlag : ushort {
            None = 0,
            ReplaceToFill = 1 << 0,
            ReplaceToContain = 1 << 1,
            ReplaceToMiddle = 1 << 2
        }

        public static void ReplacePicture(PowerPoint.ShapeRange? shapeRange, ReplacePictureCmd? cmd, ReplacePictureFlag? flag, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (cmd == null) {
                return;
            }

            var pictures = shapeRange.OfType<PowerPoint.Shape>().Where(shape => shape.Type == Office.MsoShapeType.msoPicture).ToArray();
            if (pictures.Length == 0) {
                return;
            }
            var slideShapes = GetSlideShapes(shapeRange);
            if (slideShapes == null) {
                return;
            }

            var (filepath, needCleanup) = cmd! switch {
                ReplacePictureCmd.WithClipboard => (GetFilepathForReplacingWithClipboard(), true),
                ReplacePictureCmd.WithFile => (GetFilepathForReplacingWithFile(), false),
                _ => (null, false)
            };
            if (filepath == null) {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var newShapes = InternalReplacePicture(filepath, pictures, slideShapes, flag); // <<<
            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            foreach (var shape in newShapes) {
                shape.Select(Office.MsoTriState.msoFalse);
            }

            if (needCleanup) {
                try { File.Delete(filepath); } catch (Exception) { }
            }
            uiInvalidator?.Invoke();
        }

        private static List<PowerPoint.Shape> InternalReplacePicture(string filepath, PowerPoint.Shape[] pictures, PowerPoint.Shapes slideShapes, ReplacePictureFlag? flag) {
            var replaceToFill = (flag & ReplacePictureFlag.ReplaceToFill) != 0;
            var replaceToContain = (flag & ReplacePictureFlag.ReplaceToContain) != 0;
            var replaceToMiddle = (flag & ReplacePictureFlag.ReplaceToMiddle) != 0;

            var newShapes = new List<PowerPoint.Shape>();
            foreach (var shape in pictures) {
                try {
                    var (toLink, toSaveWith) = (Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
                    var newShape = slideShapes.AddPicture(filepath, toLink, toSaveWith, shape.Left, shape.Top); // <<<
                    ApplySizeAndPositionToNewShape(shape, newShape, replaceToFill, replaceToContain, replaceToMiddle);
                    ApplyFormatAndAnimationToNewShape(shape, newShape);
                    newShapes.Add(newShape);
                    shape.Delete();
                    Marshal.ReleaseComObject(shape); // must release object to avoid 0x800a01a8 error
                } catch (Exception) {
                    // ignored
                }
            }

            return newShapes;
        }

        private static PowerPoint.Shapes? GetSlideShapes(PowerPoint.ShapeRange? shapeRange) {
            if (shapeRange?.Parent is PowerPoint.Slide slide) {
                return slide.Shapes;
            }
            return null;
        }

        private static string? GetFilepathForReplacingWithClipboard() {
            var image = Forms.Clipboard.GetImage();
            if (image == null) {
                Forms.MessageBox.Show(
                    Ribbon.ArrangeRibbonResources.dlgNoPictureInClipboard, Ribbon.ArrangeRibbonResources.dlgReplacePicture,
                    Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error);
                return null;
            }

            var path = Path.GetTempFileName();
            try {
                image.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            } catch (Exception) {
                return null;
            }
            return path;
        }

        private static string? GetFilepathForReplacingWithFile() {
            var dlg = Globals.ThisAddIn.Application.FileDialog[Office.MsoFileDialogType.msoFileDialogFilePicker];
            dlg.Title = Ribbon.ArrangeRibbonResources.dlgSelectPictureToReplace;
            dlg.AllowMultiSelect = false;

            const string imageFilter = "*.jpg; *.jpeg; *.png; *.bmp; *.gif; *.tif; *.tiff";
            const string allFilesFilter = "*.*";
            dlg.Filters.Add(Ribbon.ArrangeRibbonResources.dlgImageFilesFilter, imageFilter);
            dlg.Filters.Add(Ribbon.ArrangeRibbonResources.dlgAllFilesFilter, allFilesFilter);

            var result = dlg.Show();
            if (result != -1 || dlg.SelectedItems.Count == 0) {
                return null;
            }

            var path = dlg.SelectedItems.Item(1);
            if (string.IsNullOrEmpty(path)) {
                return null;
            }
            return path;
        }

        private static void ApplySizeAndPositionToNewShape(PowerPoint.Shape oldShape, PowerPoint.Shape newShape, bool replaceToFill, bool replaceToContain, bool replaceToMiddle) {
            var oldLockAspectRatio = newShape.LockAspectRatio;
            newShape.LockAspectRatio = Office.MsoTriState.msoFalse;
            var (oldWidth, oldHeight) = (oldShape.Width, oldShape.Height);
            var (oldLeft, oldTop) = (oldShape.Left, oldShape.Top);
            var (newWidth, newHeight) = (newShape.Width, newShape.Height);
            if (replaceToFill) {
                newHeight = oldHeight;
                newWidth = oldWidth;
            } else if (replaceToContain) {
                var widthHeightRate = newWidth / newHeight;
                if (oldHeight * widthHeightRate <= oldWidth) {
                    newHeight = oldHeight;
                    newWidth = oldHeight * widthHeightRate;
                } else {
                    newWidth = oldWidth;
                    newHeight = oldWidth / widthHeightRate;
                }
            }
            newShape.Width = newWidth;
            newShape.Height = newHeight;
            if (replaceToMiddle) {
                newShape.Left = oldLeft - (newWidth - oldWidth) / 2;
                newShape.Top = oldTop - (newHeight - oldHeight) / 2;
            }
            newShape.LockAspectRatio = oldLockAspectRatio;
        }

        private static void ApplyFormatAndAnimationToNewShape(PowerPoint.Shape oldShape, PowerPoint.Shape newShape) {
            newShape.LockAspectRatio = oldShape.LockAspectRatio;
            try {
                oldShape.PickUp();
                newShape.Apply();
            } catch (Exception) {
                // ignored
            }
            try {
                if (oldShape.AnimationSettings.Animate == Office.MsoTriState.msoTrue) {
                    oldShape.PickupAnimation();
                    newShape.ApplyAnimation();
                }
            } catch (Exception) {
                // ignored
            }
        }

    }

}
