using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.WindowsAPICodePack.Dialogs;
using Forms = System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointArrangeAddin.Helper;
using PowerPointArrangeAddin.Misc;

#nullable enable

namespace PowerPointArrangeAddin.Ribbon {

    public partial class ArrangeRibbon {

        private ArrangeRibbon() {
            (_ribbonUiElements, _specialRibbonUiElements) = GenerateNewUiElements();
            _availabilityRules = GenerateNewAvailabilityRules();
        }

        private static ArrangeRibbon? _instance;

        public static ArrangeRibbon Instance {
            get {
                _instance ??= new ArrangeRibbon();
                return _instance;
            }
        }

        private const string ArrangeRibbonXmlName = "PowerPointArrangeAddin.Ribbon.ArrangeRibbon.UI.xml";
        private const string ArrangeRibbonMenuXmlName = "PowerPointArrangeAddin.Ribbon.ArrangeRibbon.Menu.xml";

        private Office.IRibbonUI? _ribbon;

        public void Ribbon_Load(Office.IRibbonUI ribbonUi) {
            _ribbon = ribbonUi;

            // auto check for updates when PowerPoint starts up
            if (AddInSetting.Instance.CheckUpdateWhenStartUp) {
                Task.Run(async () => {
                    await Task.Delay(TimeSpan.FromSeconds(5));
                    var hwnd = new IntPtr(Globals.ThisAddIn.Application.HWND);
                    var _ = await AddInVersion.Instance.CheckUpdateAutomatically(hwnd);
                });
            }

            // apply long pressable flag from setting
            _extendDoublePressable.EnableDoublePress = AddInSetting.Instance.AllowDoublePressExtendButton;
        }

        private delegate bool AvailabilityRule(PowerPoint.ShapeRange? shapeRange, int shapesCount, bool hasTextFrame);
        private readonly Dictionary<string, AvailabilityRule> _availabilityRules;

        private Dictionary<string, AvailabilityRule> GenerateNewAvailabilityRules() {
            var map = new Dictionary<string, AvailabilityRule>();

            void Register(string id, AvailabilityRule rule) {
                map[id] = rule;
            }

            // grpArrange
            Register(btnAlignRelative, (_, cnt, _) => cnt >= 1);
            Register(btnScaleAnchor, (_, _, _) => true);
            Register(mnuArrangement, (_, _, _) => true);
            Register(btnAddInSetting, (_, _, _) => true);
            // grpAlignment
            Register(btnAlignLeft, (_, cnt, _) => cnt >= 1);
            Register(btnAlignCenter, (_, cnt, _) => cnt >= 1);
            Register(btnAlignRight, (_, cnt, _) => cnt >= 1);
            Register(btnAlignTop, (_, cnt, _) => cnt >= 1);
            Register(btnAlignMiddle, (_, cnt, _) => cnt >= 1);
            Register(btnAlignBottom, (_, cnt, _) => cnt >= 1);
            Register(btnDistributeHorizontal, (shapeRange, cnt, _) => cnt >= 1 && ArrangementHelper.IsDistributable(shapeRange, _alignRelativeFlag));
            Register(btnDistributeVertical, (shapeRange, cnt, _) => cnt >= 1 && ArrangementHelper.IsDistributable(shapeRange, _alignRelativeFlag));
            Register(btnSnapLeft, (_, cnt, _) => cnt >= 2);
            Register(btnSnapRight, (_, cnt, _) => cnt >= 2);
            Register(btnSnapTop, (_, cnt, _) => cnt >= 2);
            Register(btnSnapBottom, (_, cnt, _) => cnt >= 2);
            Register(btnGridSwitcher, (_, _, _) => true);
            Register(btnGridSetting, (_, _, _) => true);
            Register(btnAlignRelative_ToObjects, (_, cnt, _) => cnt >= 2);
            Register(btnAlignRelative_ToFirstObject, (_, cnt, _) => cnt >= 2);
            Register(btnAlignRelative_ToSlide, (_, cnt, _) => cnt >= 1);
            Register(btnSizeAndPosition, (_, cnt, _) => cnt >= 1);
            // grpResizing
            Register(btnScaleSameWidth, (_, cnt, _) => cnt >= 2);
            Register(btnScaleSameHeight, (_, cnt, _) => cnt >= 2);
            Register(btnScaleSameSize, (_, cnt, _) => cnt >= 2);
            Register(btnExtendSameLeft, (_, cnt, _) => cnt >= 2);
            Register(btnExtendSameRight, (_, cnt, _) => cnt >= 2);
            Register(btnExtendSameTop, (_, cnt, _) => cnt >= 2);
            Register(btnExtendSameBottom, (_, cnt, _) => cnt >= 2);
            Register(chkExtendToFirstObject, (_, _, _) => true);
            Register(btnScaleAnchor_FromTopLeft, (_, _, _) => true);
            Register(btnScaleAnchor_FromCenter, (_, _, _) => true);
            Register(btnScaleAnchor_FromBottomRight, (_, _, _) => true);
            // grpRotateAndFlip
            Register(btnRotateRight90, (_, cnt, _) => cnt >= 1);
            Register(btnRotateLeft90, (_, cnt, _) => cnt >= 1);
            Register(btnFlipVertical, (_, cnt, _) => cnt >= 1);
            Register(btnFlipHorizontal, (_, cnt, _) => cnt >= 1);
            Register(edtAngle, (_, cnt, _) => cnt >= 1);
            Register(btnCopyAngle, (sr, cnt, _) => cnt >= 1 && RotationHelper.IsAngleCopyable(sr));
            Register(btnPasteAngle, (_, cnt, _) => cnt >= 1 && RotationHelper.IsValidCopiedAngleValue());
            Register(btnResetAngle, (_, cnt, _) => cnt >= 1);
            // grpObjectArrange
            Register(btnMoveFront, (_, cnt, _) => cnt >= 1);
            Register(btnMoveBack, (_, cnt, _) => cnt >= 1);
            Register(btnMoveForward, (_, cnt, _) => cnt >= 1);
            Register(btnMoveBackward, (_, cnt, _) => cnt >= 1);
            Register(btnGroup, (_, cnt, _) => cnt >= 2);
            Register(btnUngroup, (sr, cnt, _) => cnt >= 1 && ArrangementHelper.IsUngroupable(sr));
            // grpObjectSize
            Register(btnResetSize, (sr, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsSizeResettable(sr));
            Register(btnLockAspectRatio, (_, cnt, _) => cnt >= 1);
            Register(edtSizeHeight, (_, cnt, _) => cnt >= 1);
            Register(edtSizeWidth, (_, cnt, _) => cnt >= 1);
            Register(btnCopySize, (sr, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsSizeCopyable(sr));
            Register(btnPasteSize, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedSizeValue());
            // grpObjectPosition
            Register(edtPositionX, (_, cnt, _) => cnt >= 1);
            Register(edtPositionY, (_, cnt, _) => cnt >= 1);
            Register(btnCopyPosition, (sr, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsPositionCopyable(sr));
            Register(btnPastePosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue());
            Register(btnDistanceType_RightLeft, (_, _, _) => true);
            Register(btnDistanceType_LeftLeft, (_, _, _) => true);
            Register(btnDistanceType_RightRight, (_, _, _) => true);
            Register(btnDistanceType_LeftRight, (_, _, _) => true);
            Register(btnCopyDistanceH, (_, cnt, _) => cnt == 2);
            Register(btnPasteDistanceH, (_, cnt, _) => cnt == 2 && SizeAndPositionHelper.IsValidCopiedDistanceHValue());
            Register(btnCopyDistanceV, (_, cnt, _) => cnt == 2);
            Register(btnPasteDistanceV, (_, cnt, _) => cnt == 2 && SizeAndPositionHelper.IsValidCopiedDistanceVValue());
            // grpTextbox
            Register(btnAutofitOff, (_, cnt, tf) => cnt >= 1 && tf);
            Register(btnAutoShrinkText, (_, cnt, tf) => cnt >= 1 && tf);
            Register(btnAutoResizeShape, (_, cnt, tf) => cnt >= 1 && tf);
            Register(btnWrapText, (_, cnt, tf) => cnt >= 1 && tf);
            Register(btnResetHorizontalMargin, (_, cnt, tf) => cnt >= 1 && tf);
            Register(edtMarginLeft, (_, cnt, tf) => cnt >= 1 && tf);
            Register(edtMarginRight, (_, cnt, tf) => cnt >= 1 && tf);
            Register(btnResetVerticalMargin, (_, cnt, tf) => cnt >= 1 && tf);
            Register(edtMarginTop, (_, cnt, tf) => cnt >= 1 && tf);
            Register(edtMarginBottom, (_, cnt, tf) => cnt >= 1 && tf);
            // grpReplacePicture
            Register(btnReplaceWithClipboard, (_, cnt, _) => cnt >= 1);
            Register(btnReplaceWithFile, (_, cnt, _) => cnt >= 1);
            Register(chkReplaceToFill, (_, _, _) => true);
            Register(chkReplaceToContain, (_, _, _) => true);
            Register(chkReplaceToMiddle, (_, _, _) => true);

            return map;
        }

        public bool GetEnabled(Office.IRibbonControl ribbonControl) {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: false);
            var shapesCount = selection.ShapeRange?.Count ?? 0;
            var hasTextFrame = selection.TextFrame != null;
            _availabilityRules.TryGetValue(ribbonControl.Id(), out var checker);
            return checker?.Invoke(selection.ShapeRange, shapesCount, hasTextFrame) ?? true;
        }

        public bool GetControlVisible(Office.IRibbonControl ribbonControl) {
            var arrangementControls = new[] { sepRotate, bgpMoveLayers, bgpRotate, bgpGroupObjects, sepArrangement, mnuArrangement };
            if (arrangementControls.Contains(ribbonControl.Id()) && ribbonControl.Group() == grpArrange) {
                return !AddInSetting.Instance.LessButtonsForArrangementGroup;
            }
            var textboxControls = new[] { sepHorizontalMargin, bgpHorizontalMargin, edtMarginLeft, edtMarginRight, sepVerticalMargin, bgpVerticalMargin, edtMarginTop, edtMarginBottom };
            if (textboxControls.Contains(ribbonControl.Id()) && ribbonControl.Group() == grpTextbox) {
                return !AddInSetting.Instance.HideMarginSettingForTextboxGroup;
            }
            return true;
        }

        public bool GetGroupVisible(Office.IRibbonControl ribbonControl) {
            return ribbonControl.Id() switch {
                grpWordArt => AddInSetting.Instance.ShowWordArtGroup,
                grpArrange => AddInSetting.Instance.ShowArrangementGroup,
                grpAddInSetting => true,
                grpAlignment => true,
                grpResizing => true,
                grpRotateAndFlip => true,
                grpObjectArrange => true,
                grpObjectSize => true,
                grpObjectPosition => true,
                grpTextbox => AddInSetting.Instance.ShowShapeTextboxGroup,
                grpReplacePicture => AddInSetting.Instance.ShowReplacePictureGroup,
                grpShapeSizeAndPosition => AddInSetting.Instance.ShowShapeSizeAndPositionGroup2,
                grpPictureSizeAndPosition => AddInSetting.Instance.ShowPictureSizeAndPositionGroup2,
                grpVideoSizeAndPosition => AddInSetting.Instance.ShowVideoSizeAndPositionGroup2,
                grpAudioSizeAndPosition => AddInSetting.Instance.ShowAudioSizeAndPositionGroup2,
                grpTableSizeAndPosition => AddInSetting.Instance.ShowTableSizeAndPositionGroup2,
                grpChartSizeAndPosition => AddInSetting.Instance.ShowChartSizeAndPositionGroup2,
                grpSmartartSizeAndPosition => AddInSetting.Instance.ShowSmartartSizeAndPositionGroup2,
                _ => true
            };
        }

        public void InvalidateRibbon(bool onlyForDrag = false) {
            if (!onlyForDrag) {
                _ribbon?.Invalidate();
            } else {
                // currently callback that only for dragging to change the position is unavailable
                _ribbon?.InvalidateControls(edtPositionX, grpObjectPosition, _sizeAndPositionGroups);
                _ribbon?.InvalidateControls(edtPositionY, grpObjectPosition, _sizeAndPositionGroups);
            }
        }

        private PowerPoint.ShapeRange? GetShapeRange(int mustMoreThanOrEqualTo = 1) {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: true);
            var shapeRange = selection.ShapeRange;
            if (shapeRange == null || shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            return shapeRange;
        }

        private (PowerPoint.TextFrame?, PowerPoint.TextFrame2?) GetTextFrame() {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: false);
            return (selection.TextFrame, selection.TextFrame2);
        }

        public void BtnAddInSetting_Click(Office.IRibbonControl _) {
            var oldLanguage = AddInSetting.Instance.Language;
            var oldIconStyle = AddInSetting.Instance.IconStyle;
            var dlg = new Dialog.SettingDialog();
            var result = dlg.ShowDialog();
            if (result != Forms.DialogResult.OK) {
                return;
            }

            _extendDoublePressable.EnableDoublePress = AddInSetting.Instance.AllowDoublePressExtendButton;
            if (AddInSetting.Instance.Language != oldLanguage) {
                // include updating elements and invalidating ribbon
                AddInLanguageChanger.ChangeLanguage(AddInSetting.Instance.Language);
            }
            if (AddInSetting.Instance.IconStyle != oldIconStyle) {
                // update elements for icons and invalidating ribbon
                UpdateUiElementsAndInvalidate();
            }
            if (AddInSetting.Instance.Language == oldLanguage && AddInSetting.Instance.IconStyle == oldIconStyle) {
                // just invalidate ribbon
                _ribbon?.Invalidate();
            }
        }

        public void BtnAddInCheckUpdate_Click(Office.IRibbonControl _) {
            Task.Run(async () => {
                var hwnd = new IntPtr(Globals.ThisAddIn.Application.HWND);
                var __ = await AddInVersion.Instance.CheckUpdateManually(hwnd);
            });
        }

        public void BtnAddInHomepage_Click(Office.IRibbonControl _) {
            using (new EnableThemingInScope(true)) {
                var dialog = new TaskDialog();
                dialog.Caption = AddInDescription.Instance.Title;
                dialog.InstructionText = ArrangeRibbonResources.dlgChooseToVisit;
                dialog.Icon = TaskDialogStandardIcon.Information;
                dialog.OwnerWindowHandle = new IntPtr(Globals.ThisAddIn.Application.HWND);
                dialog.StandardButtons = TaskDialogStandardButtons.Cancel;
                var lnkGhHomepage = new TaskDialogCommandLink("GitHub Homepage", ArrangeRibbonResources.dlgGitHubHomepage);
                var lnkGhRelease = new TaskDialogCommandLink("GitHub Release", ArrangeRibbonResources.dlgGitHubRelease);
                var lnkAcRelease = new TaskDialogCommandLink("AppCenter Release", ArrangeRibbonResources.dlgAppCenterRelease);
                lnkGhHomepage.Click += (_, _) => AccessAndClose(dialog, AddInDescription.Instance.Homepage);
                lnkGhRelease.Click += (_, _) => AccessAndClose(dialog, AddInDescription.Instance.GitHubReleaseUrl);
                lnkAcRelease.Click += (_, _) => AccessAndClose(dialog, AddInDescription.Instance.AppCenterReleaseUrl);
                dialog.Controls.Add(lnkGhHomepage);
                dialog.Controls.Add(lnkGhRelease);
                dialog.Controls.Add(lnkAcRelease);
                dialog.Show();
            }

            static void AccessAndClose(TaskDialog dlg, string url) {
                Process.Start(url);
                dlg.Close();
            }
        }

        public void BtnAddInFeedback_Click(Office.IRibbonControl _) {
            Process.Start(AddInDescription.Instance.GitHubFeedbackUrl);
        }

        public void BtnAlign_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            Office.MsoAlignCmd? cmd = ribbonControl.Id() switch {
                btnAlignLeft => Office.MsoAlignCmd.msoAlignLefts,
                btnAlignCenter => Office.MsoAlignCmd.msoAlignCenters,
                btnAlignRight => Office.MsoAlignCmd.msoAlignRights,
                btnAlignTop => Office.MsoAlignCmd.msoAlignTops,
                btnAlignMiddle => Office.MsoAlignCmd.msoAlignMiddles,
                btnAlignBottom => Office.MsoAlignCmd.msoAlignBottoms,
                _ => null
            };
            ArrangementHelper.Align(shapeRange, cmd, _alignRelativeFlag);
        }

        public void BtnDistribute_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            Office.MsoDistributeCmd? cmd = ribbonControl.Id() switch {
                btnDistributeHorizontal => Office.MsoDistributeCmd.msoDistributeHorizontally,
                btnDistributeVertical => Office.MsoDistributeCmd.msoDistributeVertically,
                _ => null
            };
            ArrangementHelper.Distribute(shapeRange, cmd, _alignRelativeFlag);
        }

        public void BtnSnap_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.SnapCmd? cmd = ribbonControl.Id() switch {
                btnSnapLeft => ArrangementHelper.SnapCmd.SnapLeftToRight,
                btnSnapRight => ArrangementHelper.SnapCmd.SnapRightToLeft,
                btnSnapTop => ArrangementHelper.SnapCmd.SnapTopToBottom,
                btnSnapBottom => ArrangementHelper.SnapCmd.SnapBottomToTop,
                _ => null
            };
            ArrangementHelper.Snap(shapeRange, cmd);
        }


        public void BtnGridSwitcher_Click(Office.IRibbonControl _, bool __) {
            const string mso = "ViewGridlinesPowerPoint";
            try {
                if (Globals.ThisAddIn.Application.CommandBars.GetEnabledMso(mso)) {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso(mso);
                }
            } catch (Exception) {
                // ignored
            } finally {
                Task.Run(async () => {
                    await Task.Delay(50);
                    _ribbon?.InvalidateControls(btnGridSwitcher, grpAlignment, mnuArrangement, (mnuArrangement, mnuAlignment));
                });
            }
        }

        public bool BtnGridSwitcher_GetPressed(Office.IRibbonControl _) {
            const string mso = "ViewGridlinesPowerPoint";
            try {
                if (Globals.ThisAddIn.Application.CommandBars.GetEnabledMso(mso)) {
                    return Globals.ThisAddIn.Application.CommandBars.GetPressedMso(mso);
                }
            } catch (Exception) {
                // ignored
            }
            return false;
        }

        public void BtnGridSetting_Click(Office.IRibbonControl _) {
            ArrangementHelper.GridSettingDialog();
        }

        // This flag is used to adjust alignment relative behavior, is used by BtnAlign_Click and BtnDistribute_Click.
        private ArrangementHelper.AlignRelativeFlag _alignRelativeFlag = ArrangementHelper.AlignRelativeFlag.RelativeToObjects;

        // Only for btnAlignRelative* and btnScaleAnchor*.
        private bool IsOptionRibbonButton(Office.IRibbonControl ribbonControl) {
            return ribbonControl.Id().Contains("_To") || ribbonControl.Id().Contains("_From"); // just check by id
        }

        public void BtnAlignRelative_Click(Office.IRibbonControl ribbonControl, bool _ = false) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null || shapeRange.Count <= 1) {
                _ribbon?.InvalidateControl(ribbonControl.Id(), ribbonControl.Group());
                return; // not change if no more than 1 shape is selected
            }
            if (!IsOptionRibbonButton(ribbonControl)) {
                _alignRelativeFlag = _alignRelativeFlag switch {
                    ArrangementHelper.AlignRelativeFlag.RelativeToObjects => ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject,
                    ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject => ArrangementHelper.AlignRelativeFlag.RelativeToSlide,
                    ArrangementHelper.AlignRelativeFlag.RelativeToSlide => ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                    _ => ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject
                };
            } else {
                _alignRelativeFlag = ribbonControl.Id() switch {
                    btnAlignRelative_ToObjects => ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                    btnAlignRelative_ToFirstObject => ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject,
                    btnAlignRelative_ToSlide => ArrangementHelper.AlignRelativeFlag.RelativeToSlide,
                    _ => ArrangementHelper.AlignRelativeFlag.RelativeToObjects
                };
            }
            ArrangementHelper.UpdateAppAlignRelative(_alignRelativeFlag);
            _ribbon?.InvalidateControl(btnAlignRelative, grpArrange);
            _ribbon?.InvalidateControls(btnDistributeHorizontal, grpArrange, grpAlignment);
            _ribbon?.InvalidateControls(btnDistributeVertical, grpArrange, grpAlignment);
            _ribbon?.InvalidateControls(btnAlignRelative_ToObjects, grpAlignment, (mnuArrangement, mnuAlignment));
            _ribbon?.InvalidateControls(btnAlignRelative_ToFirstObject, grpAlignment, (mnuArrangement, mnuAlignment));
            _ribbon?.InvalidateControls(btnAlignRelative_ToSlide, grpAlignment, (mnuArrangement, mnuAlignment));
        }

        public bool BtnAlignRelative_GetPressed(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                return false;
            }
            var shapeRange = GetShapeRange();
            if (shapeRange?.Count == 1) {
                return ribbonControl.Id() == btnAlignRelative_ToSlide; // when single shape is selected
            }
            return ribbonControl.Id() switch {
                btnAlignRelative_ToObjects => _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                btnAlignRelative_ToFirstObject => _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject,
                btnAlignRelative_ToSlide => _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToSlide,
                _ => false
            };
        }

        public string BtnAlignRelative_GetLabel(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                var shapeRange = GetShapeRange();
                if (shapeRange?.Count == 1) {
                    return GetLabel(btnAlignRelative_ToSlide); // when single shape is selected
                }
                return _alignRelativeFlag switch {
                    ArrangementHelper.AlignRelativeFlag.RelativeToObjects => GetLabel(btnAlignRelative_ToObjects),
                    ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject => GetLabel(btnAlignRelative_ToFirstObject),
                    ArrangementHelper.AlignRelativeFlag.RelativeToSlide => GetLabel(btnAlignRelative_ToSlide),
                    _ => GetLabel(btnAlignRelative_ToObjects)
                };
            }
            return GetLabel(ribbonControl.Id());
        }

        public System.Drawing.Image? BtnAlignRelative_GetImage(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                var shapeRange = GetShapeRange();
                if (shapeRange?.Count == 1) {
                    return GetImage(btnAlignRelative_ToSlide); // when single shape is selected
                }
                return _alignRelativeFlag switch {
                    ArrangementHelper.AlignRelativeFlag.RelativeToObjects => GetImage(btnAlignRelative_ToObjects),
                    ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject => GetImage(btnAlignRelative_ToFirstObject),
                    ArrangementHelper.AlignRelativeFlag.RelativeToSlide => GetImage(btnAlignRelative_ToSlide),
                    _ => GetImage(btnAlignRelative_ToObjects)
                };
            }
            return GetImage(ribbonControl.Id());
        }

        public void BtnSizeAndPosition_Click(Office.IRibbonControl _) {
            SizeAndPositionHelper.SizeAndPositionDialog();
        }

        public void BtnScale_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.ScaleSizeCmd? cmd = ribbonControl.Id() switch {
                btnScaleSameWidth => ArrangementHelper.ScaleSizeCmd.SameWidth,
                btnScaleSameHeight => ArrangementHelper.ScaleSizeCmd.SameHeight,
                btnScaleSameSize => ArrangementHelper.ScaleSizeCmd.SameSize,
                _ => null
            };
            ArrangementHelper.ScaleSize(shapeRange, cmd, _scaleFromFlag);
        }

        private DoublePressableHandler _extendDoublePressable = new() { EnableDoublePress = true };

        public void BtnExtend_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            _extendDoublePressable.CheckPress(() => BtnExtend_SinglePress(ribbonControl), () => BtnExtend_DoublePress(ribbonControl));
        }

        public void BtnExtend_SinglePress(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.ExtendSizeCmd? cmd = ribbonControl.Id() switch {
                btnExtendSameLeft => ArrangementHelper.ExtendSizeCmd.ExtendToLeft,
                btnExtendSameRight => ArrangementHelper.ExtendSizeCmd.ExtendToRight,
                btnExtendSameTop => ArrangementHelper.ExtendSizeCmd.ExtendToTop,
                btnExtendSameBottom => ArrangementHelper.ExtendSizeCmd.ExtendToBottom,
                _ => null
            };
            ArrangementHelper.ExtendSize(shapeRange, cmd, _extendToFirstObject);
        }

        public void BtnExtend_DoublePress(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.ExtendSizeCmd? cmd = ribbonControl.Id() switch {
                btnExtendSameLeft => ArrangementHelper.ExtendSizeCmd.ExtendToLeft,
                btnExtendSameRight => ArrangementHelper.ExtendSizeCmd.ExtendToRight,
                btnExtendSameTop => ArrangementHelper.ExtendSizeCmd.ExtendToTop,
                btnExtendSameBottom => ArrangementHelper.ExtendSizeCmd.ExtendToBottom,
                _ => null
            };
            ArrangementHelper.ExtendSize(shapeRange, cmd, !_extendToFirstObject);
        }

        private bool _extendToFirstObject;

        public void ChkExtendToFirstObject_Click(Office.IRibbonControl ribbonControl, bool _) {
            _extendToFirstObject = !_extendToFirstObject;
            _ribbon?.InvalidateControls(chkExtendToFirstObject, grpResizing, (mnuArrangement, mnuResizing));
        }

        public bool ChkExtendToFirstObject_GetPressed(Office.IRibbonControl ribbonControl) {
            return _extendToFirstObject;
        }

        // This flag is used to control scale and size behavior, is used by BtnScale_Click, EdtSize_TextChanged, BtnCopyAndPasteSize_Click and BtnResetMediaSize_Click.
        private SizeAndPositionHelper.ScaleFromFlag _scaleFromFlag = SizeAndPositionHelper.ScaleFromFlag.FromTopLeft;

        public void BtnScaleAnchor_Click(Office.IRibbonControl ribbonControl, bool _ = false) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                _scaleFromFlag = _scaleFromFlag switch {
                    SizeAndPositionHelper.ScaleFromFlag.FromTopLeft => SizeAndPositionHelper.ScaleFromFlag.FromTopRight,
                    SizeAndPositionHelper.ScaleFromFlag.FromTopRight => SizeAndPositionHelper.ScaleFromFlag.FromBottomLeft,
                    SizeAndPositionHelper.ScaleFromFlag.FromBottomLeft => SizeAndPositionHelper.ScaleFromFlag.FromBottomRight,
                    SizeAndPositionHelper.ScaleFromFlag.FromBottomRight => SizeAndPositionHelper.ScaleFromFlag.FromCenter,
                    SizeAndPositionHelper.ScaleFromFlag.FromCenter => SizeAndPositionHelper.ScaleFromFlag.FromTopLeft,
                    _ => SizeAndPositionHelper.ScaleFromFlag.FromTopLeft
                };
            } else {
                _scaleFromFlag = ribbonControl.Id() switch {
                    btnScaleAnchor_FromTopLeft => SizeAndPositionHelper.ScaleFromFlag.FromTopLeft,
                    btnScaleAnchor_FromTopRight => SizeAndPositionHelper.ScaleFromFlag.FromTopRight,
                    btnScaleAnchor_FromBottomLeft => SizeAndPositionHelper.ScaleFromFlag.FromBottomLeft,
                    btnScaleAnchor_FromBottomRight => SizeAndPositionHelper.ScaleFromFlag.FromBottomRight,
                    btnScaleAnchor_FromLeft => SizeAndPositionHelper.ScaleFromFlag.FromLeft,
                    btnScaleAnchor_FromRight => SizeAndPositionHelper.ScaleFromFlag.FromRight,
                    btnScaleAnchor_FromTop => SizeAndPositionHelper.ScaleFromFlag.FromTop,
                    btnScaleAnchor_FromBottom => SizeAndPositionHelper.ScaleFromFlag.FromBottom,
                    btnScaleAnchor_FromCenter => SizeAndPositionHelper.ScaleFromFlag.FromCenter,
                    _ => _scaleFromFlag
                };
            }
            _ribbon?.InvalidateControls(btnScaleAnchor, grpArrange, _sizeAndPositionGroups);
            _ribbon?.InvalidateControls(btnScaleAnchor_FromTopLeft, grpResizing, (mnuArrangement, mnuResizing));
            _ribbon?.InvalidateControls(btnScaleAnchor_FromTopRight, grpResizing, (mnuArrangement, mnuResizing));
            _ribbon?.InvalidateControls(btnScaleAnchor_FromBottomLeft, grpResizing, (mnuArrangement, mnuResizing));
            _ribbon?.InvalidateControls(btnScaleAnchor_FromBottomRight, grpResizing, (mnuArrangement, mnuResizing));
            _ribbon?.InvalidateControls(btnScaleAnchor_FromLeft, grpResizing, (mnuArrangement, mnuResizing));
            _ribbon?.InvalidateControls(btnScaleAnchor_FromRight, grpResizing, (mnuArrangement, mnuResizing));
            _ribbon?.InvalidateControls(btnScaleAnchor_FromTop, grpResizing, (mnuArrangement, mnuResizing));
            _ribbon?.InvalidateControls(btnScaleAnchor_FromBottom, grpResizing, (mnuArrangement, mnuResizing));
            _ribbon?.InvalidateControls(btnScaleAnchor_FromCenter, grpResizing, (mnuArrangement, mnuResizing));
        }

        public bool BtnScaleAnchor_GetPressed(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                return false;
            }
            return ribbonControl.Id() switch {
                btnScaleAnchor_FromTopLeft => _scaleFromFlag == SizeAndPositionHelper.ScaleFromFlag.FromTopLeft,
                btnScaleAnchor_FromTopRight => _scaleFromFlag == SizeAndPositionHelper.ScaleFromFlag.FromTopRight,
                btnScaleAnchor_FromBottomLeft => _scaleFromFlag == SizeAndPositionHelper.ScaleFromFlag.FromBottomLeft,
                btnScaleAnchor_FromBottomRight => _scaleFromFlag == SizeAndPositionHelper.ScaleFromFlag.FromBottomRight,
                btnScaleAnchor_FromLeft => _scaleFromFlag == SizeAndPositionHelper.ScaleFromFlag.FromLeft,
                btnScaleAnchor_FromRight => _scaleFromFlag == SizeAndPositionHelper.ScaleFromFlag.FromRight,
                btnScaleAnchor_FromTop => _scaleFromFlag == SizeAndPositionHelper.ScaleFromFlag.FromTop,
                btnScaleAnchor_FromBottom => _scaleFromFlag == SizeAndPositionHelper.ScaleFromFlag.FromBottom,
                btnScaleAnchor_FromCenter => _scaleFromFlag == SizeAndPositionHelper.ScaleFromFlag.FromCenter,
                _ => false
            };
        }

        public string BtnScaleAnchor_GetLabel(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                return _scaleFromFlag switch {
                    SizeAndPositionHelper.ScaleFromFlag.FromTopLeft => GetLabel(btnScaleAnchor_FromTopLeft),
                    SizeAndPositionHelper.ScaleFromFlag.FromTopRight => GetLabel(btnScaleAnchor_FromTopRight),
                    SizeAndPositionHelper.ScaleFromFlag.FromBottomLeft => GetLabel(btnScaleAnchor_FromBottomLeft),
                    SizeAndPositionHelper.ScaleFromFlag.FromBottomRight => GetLabel(btnScaleAnchor_FromBottomRight),
                    SizeAndPositionHelper.ScaleFromFlag.FromLeft => GetLabel(btnScaleAnchor_FromLeft),
                    SizeAndPositionHelper.ScaleFromFlag.FromRight => GetLabel(btnScaleAnchor_FromRight),
                    SizeAndPositionHelper.ScaleFromFlag.FromTop => GetLabel(btnScaleAnchor_FromTop),
                    SizeAndPositionHelper.ScaleFromFlag.FromBottom => GetLabel(btnScaleAnchor_FromBottom),
                    SizeAndPositionHelper.ScaleFromFlag.FromCenter => GetLabel(btnScaleAnchor_FromCenter),
                    _ => GetLabel(btnScaleAnchor_FromTopLeft)
                };
            }
            return GetLabel(ribbonControl.Id());
        }

        public System.Drawing.Image? BtnScaleAnchor_GetImage(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                return _scaleFromFlag switch {
                    SizeAndPositionHelper.ScaleFromFlag.FromTopLeft => GetImage(btnScaleAnchor_FromTopLeft),
                    SizeAndPositionHelper.ScaleFromFlag.FromTopRight => GetImage(btnScaleAnchor_FromTopRight),
                    SizeAndPositionHelper.ScaleFromFlag.FromBottomLeft => GetImage(btnScaleAnchor_FromBottomLeft),
                    SizeAndPositionHelper.ScaleFromFlag.FromBottomRight => GetImage(btnScaleAnchor_FromBottomRight),
                    SizeAndPositionHelper.ScaleFromFlag.FromLeft => GetImage(btnScaleAnchor_FromLeft),
                    SizeAndPositionHelper.ScaleFromFlag.FromRight => GetImage(btnScaleAnchor_FromRight),
                    SizeAndPositionHelper.ScaleFromFlag.FromTop => GetImage(btnScaleAnchor_FromTop),
                    SizeAndPositionHelper.ScaleFromFlag.FromBottom => GetImage(btnScaleAnchor_FromBottom),
                    SizeAndPositionHelper.ScaleFromFlag.FromCenter => GetImage(btnScaleAnchor_FromCenter),
                    _ => GetImage(btnScaleAnchor_FromTopLeft)
                };
            }
            return GetImage(ribbonControl.Id());
        }


        public void BtnRotate_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.RotateCmd? cmd = ribbonControl.Id() switch {
                btnRotateLeft90 => ArrangementHelper.RotateCmd.RotateLeft90,
                btnRotateRight90 => ArrangementHelper.RotateCmd.RotateRight90,
                _ => null
            };
            ArrangementHelper.Rotate(shapeRange, cmd);
            _ribbon?.InvalidateControl(edtAngle, grpRotateAndFlip);
        }

        public void BtnFlip_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            Office.MsoFlipCmd? cmd = ribbonControl.Id() switch {
                btnFlipHorizontal => Office.MsoFlipCmd.msoFlipHorizontal,
                btnFlipVertical => Office.MsoFlipCmd.msoFlipVertical,
                _ => null
            };
            ArrangementHelper.Flip(shapeRange, cmd);
        }

        public void EdtAngle_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            RotationHelper.ChangeAngleOfString(shapeRange, text, () => {
                _ribbon?.InvalidateControl(edtAngle, grpRotateAndFlip);
            });
        }

        public string EdtAngle_GetText(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return "";
            }
            return RotationHelper.GetAngleOfString(shapeRange).Item1;
        }

        public void BtnCopyAndPasteAngle_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            RotationHelper.CopyAndPasteCmd? cmd = ribbonControl.Id() switch {
                btnCopyAngle => RotationHelper.CopyAndPasteCmd.Copy,
                btnPasteAngle => RotationHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            RotationHelper.CopyAndPasteAngle(shapeRange, cmd, () => {
                _ribbon?.InvalidateControl(btnCopyAngle, grpRotateAndFlip);
                _ribbon?.InvalidateControl(btnPasteAngle, grpRotateAndFlip);
                _ribbon?.InvalidateControl(edtAngle, grpRotateAndFlip);
            });
        }

        public void BtnResetAngle_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            RotationHelper.ResetObjectAngle(shapeRange, () => {
                _ribbon?.InvalidateControl(edtAngle, grpRotateAndFlip);
            });
        }

        public void BtnMove_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            Office.MsoZOrderCmd? cmd = ribbonControl.Id() switch {
                btnMoveForward => Office.MsoZOrderCmd.msoBringForward,
                btnMoveBackward => Office.MsoZOrderCmd.msoSendBackward,
                btnMoveFront => Office.MsoZOrderCmd.msoBringToFront,
                btnMoveBack => Office.MsoZOrderCmd.msoSendToBack,
                _ => null
            };
            ArrangementHelper.LayerMove(shapeRange, cmd);
        }

        public void BtnGroup_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.GroupCmd? cmd = ribbonControl.Id() switch {
                btnGroup => ArrangementHelper.GroupCmd.Group,
                btnUngroup => ArrangementHelper.GroupCmd.Ungroup,
                _ => null
            };
            ArrangementHelper.Group(shapeRange, cmd, () => _ribbon?.Invalidate());
        }


        public void BtnResetMediaSize_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.ResetMediaSize(shapeRange, _scaleFromFlag, () => {
                _ribbon?.InvalidateControls(edtPositionX, grpObjectSize, _sizeAndPositionGroups);
                _ribbon?.InvalidateControls(edtPositionY, grpObjectSize, _sizeAndPositionGroups);
            });
        }

        public void BtnLockAspectRatio_Click(Office.IRibbonControl ribbonControl, bool pressed) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            var cmd = SizeAndPositionHelper.LockAspectRatioCmd.Toggle;
            SizeAndPositionHelper.ToggleLockAspectRatio(shapeRange, cmd, () => {
                _ribbon?.InvalidateControl(btnLockAspectRatio, grpTextbox);
            });
        }

        public bool BtnLockAspectRatio_GetPressed(Office.IRibbonControl _) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return false;
            }
            return SizeAndPositionHelper.GetAspectRatioIsLocked(shapeRange);
        }

        public void EdtSize_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.SizeKind? kind = ribbonControl.Id() switch {
                edtSizeHeight => SizeAndPositionHelper.SizeKind.Height,
                edtSizeWidth => SizeAndPositionHelper.SizeKind.Width,
                _ => null
            };
            SizeAndPositionHelper.ChangeSizeOfString(shapeRange, kind, _scaleFromFlag, text, () => {
                _ribbon?.InvalidateControl(edtSizeHeight, grpObjectSize);
                _ribbon?.InvalidateControl(edtSizeWidth, grpObjectSize);
            });
        }

        public string EdtSize_GetText(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return "";
            }
            SizeAndPositionHelper.SizeKind? kind = ribbonControl.Id() switch {
                edtSizeHeight => SizeAndPositionHelper.SizeKind.Height,
                edtSizeWidth => SizeAndPositionHelper.SizeKind.Width,
                _ => null
            };
            return SizeAndPositionHelper.GetSizeOfString(shapeRange, kind).Item1;
        }

        public void BtnCopyAndPasteSize_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.CopyAndPasteCmd? cmd = ribbonControl.Id() switch {
                btnCopySize => SizeAndPositionHelper.CopyAndPasteCmd.Copy,
                btnPasteSize => SizeAndPositionHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            SizeAndPositionHelper.CopyAndPasteSize(shapeRange, cmd, _scaleFromFlag, () => {
                _ribbon?.InvalidateControls(btnCopySize, grpObjectSize, _sizeAndPositionGroups);
                _ribbon?.InvalidateControls(btnPasteSize, grpObjectSize, _sizeAndPositionGroups);
                _ribbon?.InvalidateControl(edtSizeHeight, grpObjectSize);
                _ribbon?.InvalidateControl(edtSizeWidth, grpObjectSize);
            });
        }

        public void EdtPosition_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.PositionKind? kind = ribbonControl.Id() switch {
                edtPositionX => SizeAndPositionHelper.PositionKind.X,
                edtPositionY => SizeAndPositionHelper.PositionKind.Y,
                _ => null
            };
            SizeAndPositionHelper.ChangePositionOfString(shapeRange, kind, text, () => {
                _ribbon?.InvalidateControls(edtPositionX, grpObjectPosition, _sizeAndPositionGroups);
                _ribbon?.InvalidateControls(edtPositionY, grpObjectPosition, _sizeAndPositionGroups);
            });
        }

        public string EdtPosition_GetText(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return "";
            }
            SizeAndPositionHelper.PositionKind? kind = ribbonControl.Id() switch {
                edtPositionX => SizeAndPositionHelper.PositionKind.X,
                edtPositionY => SizeAndPositionHelper.PositionKind.Y,
                _ => null
            };
            return SizeAndPositionHelper.GetPositionOfString(shapeRange, kind).Item1;
        }

        public void BtnCopyAndPastePosition_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.CopyAndPasteCmd? cmd = ribbonControl.Id() switch {
                btnCopyPosition => SizeAndPositionHelper.CopyAndPasteCmd.Copy,
                btnPastePosition => SizeAndPositionHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            SizeAndPositionHelper.CopyAndPastePosition(shapeRange, cmd, () => {
                _ribbon?.InvalidateControls(btnCopyPosition, grpObjectPosition, _sizeAndPositionGroups);
                _ribbon?.InvalidateControls(btnPastePosition, grpObjectPosition, _sizeAndPositionGroups);
                _ribbon?.InvalidateControls(edtPositionX, grpObjectPosition, _sizeAndPositionGroups);
                _ribbon?.InvalidateControls(edtPositionY, grpObjectPosition, _sizeAndPositionGroups);
            });
        }

        private SizeAndPositionHelper.DistanceType _distanceType = SizeAndPositionHelper.DistanceType.RightLeft;

        public void BtnDistanceType_Click(Office.IRibbonControl ribbonControl, bool _) {
            switch (ribbonControl.Id()) {
            case btnDistanceType_RightLeft:
                _distanceType = SizeAndPositionHelper.DistanceType.RightLeft;
                break;
            case btnDistanceType_LeftLeft:
                _distanceType = SizeAndPositionHelper.DistanceType.LeftLeft;
                break;
            case btnDistanceType_RightRight:
                _distanceType = SizeAndPositionHelper.DistanceType.RightRight;
                break;
            case btnDistanceType_LeftRight:
                _distanceType = SizeAndPositionHelper.DistanceType.LeftRight;
                break;
            }
            _ribbon?.InvalidateControl(btnDistanceType_RightLeft, grpObjectPosition);
            _ribbon?.InvalidateControl(btnDistanceType_LeftLeft, grpObjectPosition);
            _ribbon?.InvalidateControl(btnDistanceType_RightRight, grpObjectPosition);
            _ribbon?.InvalidateControl(btnDistanceType_LeftRight, grpObjectPosition);
        }

        public bool BtnDistanceType_GetPressed(Office.IRibbonControl ribbonControl) {
            return ribbonControl.Id() switch {
                btnDistanceType_RightLeft => _distanceType == SizeAndPositionHelper.DistanceType.RightLeft,
                btnDistanceType_LeftLeft => _distanceType == SizeAndPositionHelper.DistanceType.LeftLeft,
                btnDistanceType_RightRight => _distanceType == SizeAndPositionHelper.DistanceType.RightRight,
                btnDistanceType_LeftRight => _distanceType == SizeAndPositionHelper.DistanceType.LeftRight,
                _ => false
            };
        }

        public void BtnCopyAndPasteDistance_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.CopyAndPasteCmd? cmd = ribbonControl.Id() switch {
                btnCopyDistanceH or btnCopyDistanceV => SizeAndPositionHelper.CopyAndPasteCmd.Copy,
                btnPasteDistanceH or btnPasteDistanceV => SizeAndPositionHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            if (ribbonControl.Id() == btnCopyDistanceH || ribbonControl.Id() == btnPasteDistanceH) {
                SizeAndPositionHelper.CopyAndPasteDistance(shapeRange, cmd, _distanceType, true, () => {
                    _ribbon?.InvalidateControls(btnCopyDistanceH, grpObjectPosition);
                    _ribbon?.InvalidateControls(btnPasteDistanceH, grpObjectPosition);
                    _ribbon?.InvalidateControls(edtPositionX, grpObjectPosition, _sizeAndPositionGroups);
                    _ribbon?.InvalidateControls(edtPositionY, grpObjectPosition, _sizeAndPositionGroups);
                });
            } else if (ribbonControl.Id() == btnCopyDistanceV || ribbonControl.Id() == btnPasteDistanceV) {
                SizeAndPositionHelper.CopyAndPasteDistance(shapeRange, cmd, _distanceType, false, () => {
                    _ribbon?.InvalidateControls(btnCopyDistanceV, grpObjectPosition);
                    _ribbon?.InvalidateControls(btnPasteDistanceV, grpObjectPosition);
                    _ribbon?.InvalidateControls(edtPositionX, grpObjectPosition, _sizeAndPositionGroups);
                    _ribbon?.InvalidateControls(edtPositionY, grpObjectPosition, _sizeAndPositionGroups);
                });
            }
        }

        public void BtnAutofit_Click(Office.IRibbonControl ribbonControl, bool _) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            TextboxHelper.TextboxStatusCmd? cmd = ribbonControl.Id() switch {
                btnAutofitOff => TextboxHelper.TextboxStatusCmd.AutofitOff,
                btnAutoShrinkText => TextboxHelper.TextboxStatusCmd.AutoShrinkText,
                btnAutoResizeShape => TextboxHelper.TextboxStatusCmd.AutoResizeShape,
                _ => null
            };
            TextboxHelper.ChangeAutofitStatus(textFrame, cmd, () => {
                _ribbon?.InvalidateControl(btnAutofitOff, grpTextbox);
                _ribbon?.InvalidateControl(btnAutoShrinkText, grpTextbox);
                _ribbon?.InvalidateControl(btnAutoResizeShape, grpTextbox);
            });
        }

        public bool BtnAutofit_GetPressed(Office.IRibbonControl ribbonControl) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return false;
            }
            TextboxHelper.TextboxStatusCmd? cmd = ribbonControl.Id() switch {
                btnAutofitOff => TextboxHelper.TextboxStatusCmd.AutofitOff,
                btnAutoShrinkText => TextboxHelper.TextboxStatusCmd.AutoShrinkText,
                btnAutoResizeShape => TextboxHelper.TextboxStatusCmd.AutoResizeShape,
                _ => null
            };
            return TextboxHelper.GetAutofitStatus(textFrame, cmd);
        }

        public void BtnWrapText_Click(Office.IRibbonControl _, bool __) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            var cmd = TextboxHelper.TextboxStatusCmd.WrapTextOnOff;
            TextboxHelper.ChangeAutofitStatus(textFrame, cmd, () => {
                _ribbon?.InvalidateControl(btnWrapText, grpTextbox);
            });
        }

        public bool BtnWrapText_GetPressed(Office.IRibbonControl _) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return false;
            }
            var cmd = TextboxHelper.TextboxStatusCmd.WrapTextOnOff;
            return TextboxHelper.GetAutofitStatus(textFrame, cmd);
        }

        public void BtnResetMargin_Click(Office.IRibbonControl ribbonControl) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            TextboxHelper.ResetMarginCmd? cmd = ribbonControl.Id() switch {
                btnResetHorizontalMargin => TextboxHelper.ResetMarginCmd.Horizontal,
                btnResetVerticalMargin => TextboxHelper.ResetMarginCmd.Vertical,
                _ => null
            };
            TextboxHelper.ResetMargin(textFrame, cmd, () => {
                _ribbon?.InvalidateControl(edtMarginLeft, grpTextbox);
                _ribbon?.InvalidateControl(edtMarginRight, grpTextbox);
                _ribbon?.InvalidateControl(edtMarginTop, grpTextbox);
                _ribbon?.InvalidateControl(edtMarginBottom, grpTextbox);
            });
        }

        public void EdtMargin_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            TextboxHelper.MarginKind? kind = ribbonControl.Id() switch {
                edtMarginLeft => TextboxHelper.MarginKind.Left,
                edtMarginRight => TextboxHelper.MarginKind.Right,
                edtMarginTop => TextboxHelper.MarginKind.Top,
                edtMarginBottom => TextboxHelper.MarginKind.Bottom,
                _ => null
            };
            TextboxHelper.ChangeMarginOfString(textFrame, kind, text, () => {
                _ribbon?.InvalidateControl(edtMarginLeft, grpTextbox);
                _ribbon?.InvalidateControl(edtMarginRight, grpTextbox);
                _ribbon?.InvalidateControl(edtMarginTop, grpTextbox);
                _ribbon?.InvalidateControl(edtMarginBottom, grpTextbox);
            });
        }

        public string EdtMargin_GetText(Office.IRibbonControl ribbonControl) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return "";
            }
            TextboxHelper.MarginKind? kind = ribbonControl.Id() switch {
                edtMarginLeft => TextboxHelper.MarginKind.Left,
                edtMarginRight => TextboxHelper.MarginKind.Right,
                edtMarginTop => TextboxHelper.MarginKind.Top,
                edtMarginBottom => TextboxHelper.MarginKind.Bottom,
                _ => null
            };
            return TextboxHelper.GetMarginOfString(textFrame, kind).Item1;
        }

        // This flag is used for picture replacing related callbacks, that is BtnReplacePicture_Click.
        private ReplacePictureHelper.ReplacePictureFlag _replacePictureFlag =
            ReplacePictureHelper.ReplacePictureFlag.ReplaceToContain | ReplacePictureHelper.ReplacePictureFlag.ReplaceToMiddle;

        public void ChkReplaceMode_Click(Office.IRibbonControl ribbonControl, bool _) {
            switch (ribbonControl.Id()) {
            case chkReplaceToFill:
                if (!ChkReplaceMode_GetPressed(ribbonControl)) {
                    _replacePictureFlag |= ReplacePictureHelper.ReplacePictureFlag.ReplaceToFill;
                    _replacePictureFlag &= ~ReplacePictureHelper.ReplacePictureFlag.ReplaceToContain;
                } else {
                    _replacePictureFlag &= ~ReplacePictureHelper.ReplacePictureFlag.ReplaceToFill;
                }
                break;
            case chkReplaceToContain:
                if (!ChkReplaceMode_GetPressed(ribbonControl)) {
                    _replacePictureFlag |= ReplacePictureHelper.ReplacePictureFlag.ReplaceToContain;
                    _replacePictureFlag &= ~ReplacePictureHelper.ReplacePictureFlag.ReplaceToFill;
                } else {
                    _replacePictureFlag &= ~ReplacePictureHelper.ReplacePictureFlag.ReplaceToContain;
                }
                break;
            case chkReplaceToMiddle:
                if (!ChkReplaceMode_GetPressed(ribbonControl)) {
                    _replacePictureFlag |= ReplacePictureHelper.ReplacePictureFlag.ReplaceToMiddle;
                } else {
                    _replacePictureFlag &= ~ReplacePictureHelper.ReplacePictureFlag.ReplaceToMiddle;
                }
                break;
            }
            _ribbon?.InvalidateControl(chkReplaceToFill, grpReplacePicture);
            _ribbon?.InvalidateControl(chkReplaceToContain, grpReplacePicture);
            _ribbon?.InvalidateControl(chkReplaceToMiddle, grpReplacePicture);
        }

        public bool ChkReplaceMode_GetPressed(Office.IRibbonControl ribbonControl) {
            return ribbonControl.Id() switch {
                chkReplaceToFill => (_replacePictureFlag & ReplacePictureHelper.ReplacePictureFlag.ReplaceToFill) != 0,
                chkReplaceToContain => (_replacePictureFlag & ReplacePictureHelper.ReplacePictureFlag.ReplaceToContain) != 0,
                chkReplaceToMiddle => (_replacePictureFlag & ReplacePictureHelper.ReplacePictureFlag.ReplaceToMiddle) != 0,
                _ => false
            };
        }

        public void BtnReplacePicture_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ReplacePictureHelper.ReplacePictureCmd? cmd = ribbonControl.Id() switch {
                btnReplaceWithClipboard => ReplacePictureHelper.ReplacePictureCmd.WithClipboard,
                btnReplaceWithFile => ReplacePictureHelper.ReplacePictureCmd.WithFile,
                _ => null
            };
            ReplacePictureHelper.ReplacePicture(shapeRange, cmd, _replacePictureFlag, () => _ribbon?.Invalidate());
        }

    }

}
