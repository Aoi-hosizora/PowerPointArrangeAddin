using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Forms = System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ppt_arrange_addin.Ribbon;


#nullable enable

namespace ppt_arrange_addin.Helper {

    public static class ReplacePictureHelper {

        public enum ReplacePictureCmd {
            WithClipboard,
            WithFile
        }

        [Flags]
        public enum ReplacePictureFlag : ushort {
            Default = 0,
            ReserveOriginalSize = 1 << 0,
            ReplaceToMiddle = 1 << 1
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

            var (filepath, needCleanup) = ("", false);
            switch (cmd!) {
            case ReplacePictureCmd.WithClipboard:
                filepath = GetFilepathForReplacingWithClipboard();
                needCleanup = true;
                break;
            case ReplacePictureCmd.WithFile:
                filepath = GetFilepathForReplacingWithFile();
                needCleanup = false;
                break;
            }
            if (filepath == null) {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var newShapes = InternalReplacePicture(filepath, pictures, slideShapes, flag);
            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            foreach (var shape in newShapes) {
                shape.Select(Office.MsoTriState.msoFalse);
            }

            if (needCleanup) {
                try {
                    File.Delete(filepath);
                } catch (Exception) {
                    // ignored
                }
            }

            uiInvalidator?.Invoke();
        }

        private static List<PowerPoint.Shape> InternalReplacePicture(string filepath, IEnumerable<PowerPoint.Shape> pictures, PowerPoint.Shapes slideShapes, ReplacePictureFlag? flag) {
            var reserveOriginalSize = (flag & ReplacePictureFlag.ReserveOriginalSize) != 0;
            var replaceToMiddle = (flag & ReplacePictureFlag.ReplaceToMiddle) != 0;

            var newShapes = new List<PowerPoint.Shape>();
            foreach (var shape in pictures) {
                try {
                    var (toLink, toSaveWith) = (Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
                    var newShape = slideShapes.AddPicture(filepath, toLink, toSaveWith, shape.Left, shape.Top);
                    newShape.LockAspectRatio = shape.LockAspectRatio;
                    // TODO apply old format

                    var (oldWidth, oldHeight) = (shape.Width, shape.Height);
                    var (oldLeft, oldTop) = (shape.Left, shape.Top);
                    var (newWidth, newHeight) = (newShape.Width, newShape.Height);

                    if (reserveOriginalSize) {
                        var widthHeightRate = newWidth / newHeight;
                        if (oldHeight * widthHeightRate <= oldWidth) {
                            newHeight = oldHeight;
                            newWidth = oldHeight * widthHeightRate;
                        } else {
                            newWidth = oldWidth;
                            newHeight = oldWidth / widthHeightRate;
                        }
                        newShape.Width = newWidth;
                        newShape.Height = newHeight;
                    }

                    if (replaceToMiddle) {
                        newShape.Left = oldLeft - (newWidth - oldWidth) / 2;
                        newShape.Top = oldTop - (newHeight - oldHeight) / 2;
                    }

                    newShapes.Add(newShape);
                    shape.Delete();
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
                    ArrangeRibbonResources.dlgNoPictureInClipboard, ArrangeRibbonResources.dlgReplacePicture,
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
            dlg.Title = ArrangeRibbonResources.dlgSelectPictureToReplace;
            dlg.AllowMultiSelect = false;

            const string imageFilter = "*.jpg; *.jpeg; *.png; *.bmp; *.gif; *.tif; *.tiff";
            const string allFilesFilter = "*.*";
            dlg.Filters.Add(ArrangeRibbonResources.dlgImageFilesFilter, imageFilter);
            dlg.Filters.Add(ArrangeRibbonResources.dlgAllFilesFilter, allFilesFilter);

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

    }

}
