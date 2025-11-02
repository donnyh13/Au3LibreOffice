# Changelog

All notable changes to ["Au3LibreOffice"](https://github.com/mlipok/Au3LibreOffice/tree/main) SDK/API will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
This project also adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend---types-of-changes) for further information about the types of changes.

## Releases

|    Version       |    Changes                         |    Download                 |     Released   |    Compare on GitHub       |
|:-----------------|:----------------------------------:|:---------------------------:|:--------------:|:---------------------------|
|    **v0.10.0**   | [Change Log](#0100---2024-04-)     | [v0.10.0][v0.10.0]          | _Unreleased_   | [Compare][v0.10.0-Compare] |
|    **v0.9.1**    | [Change Log](#091---2023-10-28)    | [v0.9.1][v0.9.1]            | 2023-10-28     | [Compare][v0.9.1-Compare]  |
|    **v0.9.0**    | [Change Log](#090---2023-10-28)    | [v0.9.0][v0.9.0]            | 2023-10-28     | [Compare][v0.9.0-Compare]  |
|    **v0.0.0.3**  | [Change Log](#0003---2023-08-10)   | [v0.0.0.3][v0.0.0.3]        | 2023-08-10     | [Compare][v0.0.0.3-Compare]|
|    **v0.0.0.2**  | [Change Log](#0002---2023-07-16)   | [v0.0.0.2][v0.0.0.2]        | 2023-07-16     | [Compare][v0.0.0.2-Compare]|
|    **v0.0.0.1**  | [Change Log](#0001---2023-07-02)   | [v0.0.0.1][v0.0.0.1]        | 2023-07-02     |                            |

## [0.10.0] - 2024-04-

### Project

- Added logo to ReadMe. @mLipok
- Filled in ReadMe. @mLipok
- Formatted Changelog

> [!NOTE]
> **LibreOfficeUDF**

### Added

- Central Constants File
	- LibreOffice_Constants.au3
		- $__LOCONST_SLEEP_DIV
		- $LO_COLOR_*
		- $LO_PATHCONV_*
		- $__LO_STATUS_*
		- $__LOCONST_CONVERT_*
	- LibreOffice_Helper.au3
		- _LO_ComError_UserFunction
		- _LO_ConvertColorFromLong
		- _LO_ConvertColorToLong
		- _LO_ConvertFromMicrometer
		- _LO_ConvertToMicrometer
		- _LO_InitializePortable
		- _LO_PathConvert
		- _LO_VersionGet
	- LibreOffice_Internal.au3
		- __LO_AddTo1DArray
		- __LO_ArrayFill
		- __LO_CreateStruct
		- __LO_DeleteTempReg
		- __LO_InternalComErrorHandler
		- __LO_IntIsBetween
		- __LO_NumIsBetween
		- __LO_ServiceManager
		- __LO_SetPortableServiceManager
		- __LO_SetPropertyValue
		- __LO_UnitConvert
		- __LO_VarsAreNull
		- __LO_VersionCheck
- Central UDF File for all components (@mLipok)
	- LibreOffice.au3
- Support for LibreOffice Portable usage. See `_LO_InitializePortable`.

### Changed

- All Internal Error Constants from `$__LOW_STATUS_` or `$__LOC_STATUS_` To `$__LO_STATUS_`
- Attempted to standardize `$__LO_STATUS_INIT_ERROR` and `$__LO_STATUS_PROCESSING_ERROR` usage throughout functions.
- Removed Error returns listed in Function Headers that no longer existed.
- Added missing error values and corrected wrong error values listed in the headers.
- All calls to ObjCreate to create a com.sun.star.ServiceManager Object are routed through an internal Function which stores a reference to the Object rather than creating a new instance each time, this also allows the ability to automate Portable LO. Affected Functions are (and any functions using these functions):
	- _LOCalc_RangeSortAlt
	- __LOCalc_CreateStruct
	- _LOCalc_DocConnect
	- _LOCalc_DocCreate
	- _LOCalc_DocOpen
	- _LOCalc_DocPrintersGetNames
	- _LOCalc_FontExists
	- _LOCalc_FontsGetNames
	- _LOCalc_VersionGet
	- __LOWriter_CreateStruct
	- __LOWriter_NumStyleInitiateDocument
	- _LOWriter_DirFrmtClear
	- _LOWriter_DocConnect
	- _LOWriter_DocConvertTableToText
	- _LOWriter_DocConvertTextToTable
	- _LOWriter_DocCreate
	- _LOWriter_DocExecuteDispatch
	- _LOWriter_DocOpen
	- _LOWriter_DocPrintersGetNames
	- _LOWriter_DocReplaceAllInRange
	- _LOWriter_DocZoom
	- _LOWriter_FontExists
	- _LOWriter_FontsGetNames
	- _LOWriter_FormPropertiesData
	- _LOWriter_ImageInsert
	- _LOWriter_VersionGet

### Fixed

- Align Parameters, Error/Return values, Remarks, and Related, to the same position.

### Removed

- "Note" from Remarks section in Header. (@mLipok)
- Double spaces from Headers.
- Tabs from headers, replaced with spaces.
- Manual line breaks from headers.
- Removed Error returns listed in Function Headers that no longer existed.

> [!NOTE]
> **LibreOfficeCalc**

### Added

- Main Calc File
	- LibreOfficeCalc.au3
- Individual Calc Element Files
	- LibreOfficeCalc_Cell.au3
	- LibreOfficeCalc_CellStyle.au3
	- LibreOfficeCalc_Comments.au3
	- LibreOfficeCalc_Constants.au3
	- LibreOfficeCalc_Cursor.au3
	- LibreOfficeCalc_Doc.au3
	- LibreOfficeCalc_Field.au3
	- LibreOfficeCalc_Helper.au3
	- LibreOfficeCalc_Internal.au3
	- LibreOfficeCalc_Page.au3
	- LibreOfficeCalc_Range.au3
	- LibreOfficeCalc_Sheet.au3
- Cell/Cell Range Formatting Functions and Examples
	- _LOCalc_CellBackColor
	- _LOCalc_CellBorderColor
	- _LOCalc_CellBorderPadding
	- _LOCalc_CellBorderStyle
	- _LOCalc_CellBorderWidth
	- _LOCalc_CellCreateTextCursor
	- _LOCalc_CellEffect
	- _LOCalc_CellFont
	- _LOCalc_CellFontColor
	- _LOCalc_CellFormula
	- _LOCalc_CellGetType
	- _LOCalc_CellNumberFormat
	- _LOCalc_CellOverline
	- _LOCalc_CellProtection
	- _LOCalc_CellShadow
	- _LOCalc_CellStrikeOut
	- _LOCalc_CellString
	- _LOCalc_CellTextAlign
	- _LOCalc_CellTextOrient
	- _LOCalc_CellTextProperties
	- _LOCalc_CellUnderline
	- _LOCalc_CellValue
- Cell/Cell Range Functions and Examples
	- _LOCalc_RangeAutoOutline
	- _LOCalc_RangeClearContents
	- _LOCalc_RangeColumnDelete
	- _LOCalc_RangeColumnGetName
	- _LOCalc_RangeColumnGetObjByName
	- _LOCalc_RangeColumnGetObjByPosition
	- _LOCalc_RangeColumnInsert
	- _LOCalc_RangeColumnPageBreak
	- _LOCalc_RangeColumnsGetCount
	- _LOCalc_RangeColumnVisible
	- _LOCalc_RangeColumnWidth
	- _LOCalc_RangeCompute
	- _LOCalc_RangeCopyMove
	- _LOCalc_RangeCreateCursor
	- _LOCalc_RangeData
	- _LOCalc_RangeDatabaseAdd
	- _LOCalc_RangeDatabaseDelete
	- _LOCalc_RangeDatabaseGetNames
	- _LOCalc_RangeDatabaseGetObjByName
	- _LOCalc_RangeDatabaseHasByName
	- _LOCalc_RangeDatabaseModify
	- _LOCalc_RangeDelete
	- _LOCalc_RangeDetail
	- _LOCalc_RangeFill
	- _LOCalc_RangeFillRandom
	- _LOCalc_RangeFillSeries
	- _LOCalc_RangeFilter
	- _LOCalc_RangeFilterAdvanced
	- _LOCalc_RangeFilterClear
	- _LOCalc_RangeFindAll
	- _LOCalc_RangeFindNext
	- _LOCalc_RangeFormula
	- _LOCalc_RangeGetAddressAsName
	- _LOCalc_RangeGetAddressAsPosition
	- _LOCalc_RangeGetCellByName
	- _LOCalc_RangeGetCellByPosition
	- _LOCalc_RangeGetSheet
	- _LOCalc_RangeGroup
	- _LOCalc_RangeInsert
	- _LOCalc_RangeIsMerged
	- _LOCalc_RangeMerge
	- _LOCalc_RangeNamedAdd
	- _LOCalc_RangeNamedChangeScope
	- _LOCalc_RangeNamedDelete
	- _LOCalc_RangeNamedGetNames
	- _LOCalc_RangeNamedGetObjByName
	- _LOCalc_RangeNamedHasByName
	- _LOCalc_RangeNamedModify
	- _LOCalc_RangeNumbers
	- _LOCalc_RangeOutlineClearAll
	- _LOCalc_RangeOutlineShow
	- _LOCalc_RangePivotDelete
	- _LOCalc_RangePivotDest
	- _LOCalc_RangePivotExists
	- _LOCalc_RangePivotFieldGetObjByName
	- _LOCalc_RangePivotFieldItemsGetNames
	- _LOCalc_RangePivotFieldsColumnsGetNames
	- _LOCalc_RangePivotFieldsDataGetNames
	- _LOCalc_RangePivotFieldSettings
	- _LOCalc_RangePivotFieldsFiltersGetNames
	- _LOCalc_RangePivotFieldsGetNames
	- _LOCalc_RangePivotFieldsRowsGetNames
	- _LOCalc_RangePivotFieldsUnusedGetNames
	- _LOCalc_RangePivotFilter
	- _LOCalc_RangePivotFilterClear
	- _LOCalc_RangePivotGetObjByIndex
	- _LOCalc_RangePivotGetObjByName
	- _LOCalc_RangePivotInsert
	- _LOCalc_RangePivotName
	- _LOCalc_RangePivotRefresh
	- _LOCalc_RangePivotSettings
	- _LOCalc_RangePivotsGetCount
	- _LOCalc_RangePivotsGetNames
	- _LOCalc_RangePivotSource
	- _LOCalc_RangeQueryColumnDiff
	- _LOCalc_RangeQueryContents
	- _LOCalc_RangeQueryDependents
	- _LOCalc_RangeQueryEmpty
	- _LOCalc_RangeQueryFormula
	- _LOCalc_RangeQueryIntersection
	- _LOCalc_RangeQueryPrecedents
	- _LOCalc_RangeQueryRowDiff
	- _LOCalc_RangeQueryVisible
	- _LOCalc_RangeReplace
	- _LOCalc_RangeReplaceAll
	- _LOCalc_RangeRowDelete
	- _LOCalc_RangeRowGetObjByPosition
	- _LOCalc_RangeRowHeight
	- _LOCalc_RangeRowInsert
	- _LOCalc_RangeRowPageBreak
	- _LOCalc_RangeRowsGetCount
	- _LOCalc_RangeRowVisible
	- _LOCalc_RangeSort
	- _LOCalc_RangeSortAlt
	- _LOCalc_RangeValidation
	- _LOCalc_RangeValidationSettings
- Cell Style Formatting Functions and Examples
	- _LOCalc_CellStyleBackColor
	- _LOCalc_CellStyleBorderColor
	- _LOCalc_CellStyleBorderPadding
	- _LOCalc_CellStyleBorderStyle
	- _LOCalc_CellStyleBorderWidth
	- _LOCalc_CellStyleCreate
	- _LOCalc_CellStyleDelete
	- _LOCalc_CellStyleEffect
	- _LOCalc_CellStyleExists
	- _LOCalc_CellStyleFont
	- _LOCalc_CellStyleFontColor
	- _LOCalc_CellStyleGetObj
	- _LOCalc_CellStyleNumberFormat
	- _LOCalc_CellStyleOrganizer
	- _LOCalc_CellStyleOverline
	- _LOCalc_CellStyleProtection
	- _LOCalc_CellStyleSet
	- _LOCalc_CellStylesGetNames
	- _LOCalc_CellStyleShadow
	- _LOCalc_CellStyleStrikeOut
	- _LOCalc_CellStyleTextAlign
	- _LOCalc_CellStyleTextOrient
	- _LOCalc_CellStyleTextProperties
	- _LOCalc_CellStyleUnderline
- Comment Functions and Examples.
	- _LOCalc_CommentAdd
	- _LOCalc_CommentAreaColor
	- _LOCalc_CommentAreaFillStyle
	- _LOCalc_CommentAreaGradient
	- _LOCalc_CommentAreaGradientMulticolor
	- _LOCalc_CommentAreaShadow
	- _LOCalc_CommentAreaTransparency
	- _LOCalc_CommentAreaTransparencyGradient
	- _LOCalc_CommentAreaTransparencyGradientMulti
	- _LOCalc_CommentCallout
	- _LOCalc_CommentCreateTextCursor
	- _LOCalc_CommentDelete
	- _LOCalc_CommentGetCell
	- _LOCalc_CommentGetLastEdit
	- _LOCalc_CommentGetObjByCell
	- _LOCalc_CommentGetObjByIndex
	- _LOCalc_CommentLineArrowStyles
	- _LOCalc_CommentLineProperties
	- _LOCalc_CommentPosition
	- _LOCalc_CommentRotate
	- _LOCalc_CommentsGetCount
	- _LOCalc_CommentsGetList
	- _LOCalc_CommentSize
	- _LOCalc_CommentText
	- _LOCalc_CommentTextAnchor
	- _LOCalc_CommentTextAnimation
	- _LOCalc_CommentTextColumns
	- _LOCalc_CommentTextSettings
	- _LOCalc_CommentVisible
- Cursor Functions and Examples
	- _LOCalc_SheetCursorMove
	- _LOCalc_TextCursorCharPosition
	- _LOCalc_TextCursorCharSpacing
	- _LOCalc_TextCursorEffect
	- _LOCalc_TextCursorFont
	- _LOCalc_TextCursorFontColor
	- _LOCalc_TextCursorGetString
	- _LOCalc_TextCursorGoToRange
	- _LOCalc_TextCursorInsertString
	- _LOCalc_TextCursorIsCollapsed
	- _LOCalc_TextCursorMove
	- _LOCalc_TextCursorOverline
	- _LOCalc_TextCursorParObjCreateList
	- _LOCalc_TextCursorParObjSectionsGet
	- _LOCalc_TextCursorStrikeOut
	- _LOCalc_TextCursorUnderline
- Document Functions and Examples
	- _LOCalc_DocClose
	- _LOCalc_DocColumnsRowsAreFrozen
	- _LOCalc_DocColumnsRowsFreeze
	- _LOCalc_DocConnect
	- _LOCalc_DocCreate
	- _LOCalc_DocExport
	- _LOCalc_DocFormulaBarHeight
	- _LOCalc_DocGetName
	- _LOCalc_DocGetPath
	- _LOCalc_DocHasPath
	- _LOCalc_DocHasSheetName
	- _LOCalc_DocIsActive
	- _LOCalc_DocIsModified
	- _LOCalc_DocIsReadOnly
	- _LOCalc_DocMaximize
	- _LOCalc_DocMinimize
	- _LOCalc_DocOpen
	- _LOCalc_DocPosAndSize
	- _LOCalc_DocPrint
	- _LOCalc_DocPrintersGetNames
	- _LOCalc_DocPrintersAltGetNames
	- _LOCalc_DocRedo
	- _LOCalc_DocRedoClear
	- _LOCalc_DocRedoCurActionTitle
	- _LOCalc_DocRedoGetAllActionTitles
	- _LOCalc_DocRedoIsPossible
	- _LOCalc_DocSave
	- _LOCalc_DocSaveAs
	- _LOCalc_DocSelectionCopy
	- _LOCalc_DocSelectionGet
	- _LOCalc_DocSelectionPaste
	- _LOCalc_DocSelectionSet
	- _LOCalc_DocSelectionSetMulti
	- _LOCalc_DocToFront
	- _LOCalc_DocUndo
	- _LOCalc_DocUndoActionBegin
	- _LOCalc_DocUndoActionEnd
	- _LOCalc_DocUndoClear
	- _LOCalc_DocUndoCurActionTitle
	- _LOCalc_DocUndoGetAllActionTitles
	- _LOCalc_DocUndoIsPossible
	- _LOCalc_DocUndoReset
	- _LOCalc_DocViewDisplaySettings
	- _LOCalc_DocViewWindowSettings
	- _LOCalc_DocVisible
	- _LOCalc_DocWindowFirstColumn
	- _LOCalc_DocWindowFirstRow
	- _LOCalc_DocWindowIsSplit
	- _LOCalc_DocWindowSplit
	- _LOCalc_DocWindowVisibleRange
	- _LOCalc_DocZoom
- Field Functions
	- _LOCalc_FieldCurrentDisplayGet
	- _LOCalc_FieldDateTimeInsert
	- _LOCalc_FieldDelete
	- _LOCalc_FieldFileNameInsert
	- _LOCalc_FieldGetAnchor
	- _LOCalc_FieldHyperlinkInsert
	- _LOCalc_FieldHyperlinkModify
	- _LOCalc_FieldPageCountInsert
	- _LOCalc_FieldPageNumberInsert
	- _LOCalc_FieldsGetList
	- _LOCalc_FieldSheetNameInsert
	- _LOCalc_FieldTitleInsert
- Helper Functions
	- _LOCalc_ComError_UserFunction
	- ~~_LOCalc_ConvertColorFromLong~~
	- ~~_LOCalc_ConvertColorToLong~~
	- ~~_LOCalc_ConvertFromMicrometer~~
	- ~~_LOCalc_ConvertToMicrometer~~
	- _LOCalc_FilterDescriptorCreate
	- _LOCalc_FilterDescriptorModify
	- _LOCalc_FilterFieldCreate
	- _LOCalc_FilterFieldModify
	- _LOCalc_FontExists
	- _LOCalc_FontsGetNames
	- _LOCalc_FormatKeyCreate
	- _LOCalc_FormatKeyDelete
	- _LOCalc_FormatKeyExists
	- _LOCalc_FormatKeyGetStandard
	- _LOCalc_FormatKeyGetString
	- _LOCalc_FormatKeyList
	- _LOCalc_GradientMulticolorAdd
	- _LOCalc_GradientMulticolorDelete
	- _LOCalc_GradientMulticolorModify
	- ~~_LOCalc_PathConvert~~
	- _LOCalc_SearchDescriptorCreate
	- _LOCalc_SearchDescriptorModify
	- _LOCalc_SearchDescriptorSimilarityModify
	- _LOCalc_SortFieldCreate
	- _LOCalc_SortFieldModify
	- _LOCalc_TransparencyGradientMultiAdd
	- _LOCalc_TransparencyGradientMultiDelete
	- _LOCalc_TransparencyGradientMultiModify
	- ~~_LOCalc_VersionGet~~
- Internal Functions
	- ~~__LOCalc_AddTo1DArray~~
	- ~~__LOCalc_ArrayFill~~
	- __LOCalc_CellAddressIsSame
	- __LOCalc_CellBackColor
	- __LOCalc_CellBorder
	- __LOCalc_CellBorderPadding
	- __LOCalc_CellEffect
	- __LOCalc_CellFont
	- __LOCalc_CellFontColor
	- __LOCalc_CellNumberFormat
	- __LOCalc_CellOverLine
	- __LOCalc_CellProtection
	- __LOCalc_CellShadow
	- __LOCalc_CellStrikeOut
	- __LOCalc_CellStyleBorder
	- __LOCalc_CellTextAlign
	- __LOCalc_CellTextOrient
	- __LOCalc_CellTextProperties
	- __LOCalc_CellUnderLine
	- __LOCalc_CharPosition
	- __LOCalc_CharSpacing
	- __LOCalc_CommentAreaShadowModify
	- __LOCalc_CommentArrowStyleName
	- __LOCalc_CommentGetObjByCell
	- __LOCalc_CommentLineStyleName
	- ~~__LOCalc_CreateStruct~~
	- __LOCalc_FieldGetObj
	- __LOCalc_FieldTypeServices
	- __LOCalc_FilterNameGet
	- __LOCalc_GradientNameInsert
	- __LOCalc_GradientPresets
	- __LOCalc_Internal_CursorGetType
	- __LOCalc_InternalComErrorHandler
	- ~~__LOCalc_IntIsBetween~~
	- __LOCalc_NamedRangeGetScopeObj
	- ~~__LOCalc_NumIsBetween~~
	- __LOCalc_PageStyleBorder
	- __LOCalc_PageStyleFooterBorder
	- __LOCalc_PageStyleHeaderBorder
	- __LOCalc_RangeAddressIsSame
	- ~~__LOCalc_SetPropertyValue~~
	- __LOCalc_SheetCursorMove
	- __LOCalc_TextCursorMove
	- __LOCalc_TransparencyGradientConvert
	- __LOCalc_TransparencyGradientNameInsert
	- ~~__LOCalc_UnitConvert~~
	- ~~__LOCalc_VarsAreNull~~
	- ~~__LOCalc_VersionCheck~~
- Page Style Functions and Examples
	- _LOCalc_PageStyleAreaColor
	- _LOCalc_PageStyleBorderColor
	- _LOCalc_PageStyleBorderPadding
	- _LOCalc_PageStyleBorderStyle
	- _LOCalc_PageStyleBorderWidth
	- _LOCalc_PageStyleCreate
	- _LOCalc_PageStyleDelete
	- _LOCalc_PageStyleExists
	- _LOCalc_PageStyleFooter
	- _LOCalc_PageStyleFooterAreaColor
	- _LOCalc_PageStyleFooterBorderColor
	- _LOCalc_PageStyleFooterBorderPadding
	- _LOCalc_PageStyleFooterBorderStyle
	- _LOCalc_PageStyleFooterBorderWidth
	- _LOCalc_PageStyleFooterCreateTextCursor
	- _LOCalc_PageStyleFooterObj
	- _LOCalc_PageStyleFooterShadow
	- _LOCalc_PageStyleGetObj
	- _LOCalc_PageStyleHeader
	- _LOCalc_PageStyleHeaderAreaColor
	- _LOCalc_PageStyleHeaderBorderColor
	- _LOCalc_PageStyleHeaderBorderPadding
	- _LOCalc_PageStyleHeaderBorderStyle
	- _LOCalc_PageStyleHeaderBorderWidth
	- _LOCalc_PageStyleHeaderCreateTextCursor
	- _LOCalc_PageStyleHeaderObj
	- _LOCalc_PageStyleHeaderShadow
	- _LOCalc_PageStyleLayout
	- _LOCalc_PageStyleMargins
	- _LOCalc_PageStyleOrganizer
	- _LOCalc_PageStylePaperFormat
	- _LOCalc_PageStyleSet
	- _LOCalc_PageStylesGetNames
	- _LOCalc_PageStyleShadow
	- _LOCalc_PageStyleSheetPageOrder
	- _LOCalc_PageStyleSheetPrint
	- _LOCalc_PageStyleSheetScale
- Sheet Functions and Examples
	- _LOCalc_SheetActivate
	- _LOCalc_SheetAdd
	- _LOCalc_SheetCopy
	- _LOCalc_SheetCreateCursor
	- _LOCalc_SheetDetectiveClear
	- _LOCalc_SheetDetectiveDependent
	- _LOCalc_SheetDetectiveInvalidData
	- _LOCalc_SheetDetectivePrecedent
	- _LOCalc_SheetDetectiveTraceError
	- _LOCalc_SheetGetActive
	- _LOCalc_SheetGetObjByName
	- _LOCalc_SheetGetObjByPosition
	- _LOCalc_SheetImport
	- _LOCalc_SheetIsActive
	- _LOCalc_SheetIsProtected
	- _LOCalc_SheetLink
	- _LOCalc_SheetLinkModify
	- _LOCalc_SheetMove
	- _LOCalc_SheetName
	- _LOCalc_SheetPrintColumnsRepeat
	- _LOCalc_SheetPrintRangeModify
	- _LOCalc_SheetPrintRowsRepeat
	- _LOCalc_SheetProtect
	- _LOCalc_SheetRemove
	- _LOCalc_SheetsGetCount
	- _LOCalc_SheetsGetNames
	- _LOCalc_SheetTabColor
	- _LOCalc_SheetUnprotect
	- _LOCalc_SheetVisible
- Calc Constants
	- $__LOCCONST_FILL_STYLE_*
	- $LOC_BORDERSTYLE_*
	- $LOC_BORDERWIDTH_*
	- $LOC_CELL_ALIGN_HORI_*
	- $LOC_CELL_ALIGN_VERT_*
	- $LOC_CELL_DELETE_MODE_*
	- $LOC_CELL_FLAG_*
	- $LOC_CELL_INSERT_MODE_*
	- $LOC_CELL_ROTATE_REF_*
	- $LOC_CELL_TYPE_*
	- ~~$LOC_COLOR_*~~
	- $LOC_COMMENT_ANCHOR_*
	- $LOC_COMMENT_ANIMATION_DIR_*
	- $LOC_COMMENT_ANIMATION_KIND_*
	- $LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_*
	- $LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_*
	- $LOC_COMMENT_CALLOUT_EXT_*
	- $LOC_COMMENT_CALLOUT_STYLE_*
	- $LOC_COMMENT_LINE_ARROW_TYPE_*
	- $LOC_COMMENT_LINE_CAP_*
	- $LOC_COMMENT_LINE_JOINT_*
	- $LOC_COMMENT_LINE_STYLE_*
	- $LOC_COMMENT_SHADOW_*
	- $LOC_COMPUTE_FUNC_*
	- $LOC_CURTYPE_*
	- $LOC_DUPLEX_*
	- $LOC_FIELD_TYPE_*
	- $LOC_FILL_DATE_MODE_*
	- $LOC_FILL_DIR_*
	- $LOC_FILL_MODE_*
	- $LOC_FILTER_CONDITION_*
	- $LOC_FILTER_OPERATOR_*
	- $LOC_FORMAT_KEYS_*
	- $LOC_FORMULA_RESULT_TYPE_*
	- $LOC_GRAD_NAME_*
	- $LOC_GRAD_TYPE_*
	- $LOC_GROUP_ORIENT_*
	- $LOC_NAMED_RANGE_OPT_*
	- $LOC_NUM_STYLE_*
	- $LOC_PAGE_LAYOUT_*
	- $LOC_PAPER_HEIGHT_*
	- $LOC_PAPER_WIDTH_*
	- ~~$LOC_PATHCONV_*~~
	- $LOC_PIVOT_TBL_FIELD_BASE_*
	- $LOC_PIVOT_TBL_FIELD_DISP_*
	- $LOC_PIVOT_TBL_FIELD_TYPE_*
	- $LOC_POSTURE_*
	- $LOC_RELIEF_*
	- $LOC_SCALE_*
	- $LOC_SEARCH_IN_*
	- $LOC_SHADOW_*
	- $LOC_SHEET_LINK_MODE_*
	- $LOC_SHEETCUR_*
	- $LOC_SORT_DATA_TYPE_*
	- $LOC_STRIKEOUT_*
	- $LOC_TEXTCUR_*
	- $LOC_TXT_DIR_*
	- $LOC_UNDERLINE_*
	- $LOC_VALIDATION_COND_*
	- $LOC_VALIDATION_ERROR_*
	- $LOC_VALIDATION_LIST_*
	- $LOC_VALIDATION_TYPE_*
	- $LOC_WEIGHT_*
	- $LOC_ZOOMTYPE_*
- Auto size option to Range Data, Formulas, and Numbers fill functions.
- Retrieve Linked Sheet names only to `_LOCalc_SheetsGetNames`.
- Line numbers to Example Error messages.
- Top-Most attribute to Example message boxes.
- LibreOffice SDK/API Constant names to constants.

### Fixed

- Removed unused variables and parameters in some functions. Affected functions are as follows:
	- `_LOCalc_FormatKeyDelete` -- removed internal variable.
- NamedRange names incorrectly reported as invalid in certain functions when the name began with an underscore.
	- _LOCalc_RangeNamedModify
	- _LOCalc_RangeNamedAdd
- `_LOCalc_DocOpen` now uses a different method for connecting to an already open document, as the previous method was causing errors.
- `_LOCalc_DocCreate` would return if there was an error creating a property, instead of increasing the error count.
- `LibreOfficeCalc_Cell.au3` was missing an Include file.
- Several Cell or Cell range functions that should support Column/Rows would not work with them.
- `LibreOfficeCalc_Sheet.au3` was missing an Include file.
- `_LOCalc_DocViewWindowSettings`, return values in wrong order. Thanks to user JALucena. <https://www.autoitscript.com/forum/topic/210514-libreoffice-udf-help-and-support/page/2/#findComment-1543326>
- `_LOCalc_DocCreate` not finding a blank open document to connect to, if available, due to reversed logical operator, and non-existent method.
- Missing error values in the header and wrong error values in `_LOCalc_CommentAreaTransparencyGradient`.

### Changed

- Constant `$__LOCCONST_FILL_STYLE_*` to `$LOC_AREA_FILL_STYLE_*`
- `__LOCalc_IntIsBetween` to accept only a minimum value. Also optimized it.
	- Modified function usage to match changes.
- Renamed Constant `$LOC_COMPUTE_*` to `$LOC_COMPUTE_FUNC_*`
- Renamed Functions to be consistent when retrieving arrays of names or objects:
	- `_LOCalc_DocEnumPrinters` --> `_LOCalc_DocPrintersGetNames`
	- `_LOCalc_DocEnumPrintersAlt` --> `_LOCalc_DocPrintersAltGetNames`
	- `_LOCalc_FontsList` --> `_LOCalc_FontsGetNames`
	- `_LOCalc_RangePivotFieldItemsGetList` --> `_LOCalc_RangePivotFieldItemsGetNames`
	- `_LOCalc_RangePivotFieldsColumnsGetList` --> `_LOCalc_RangePivotFieldsColumnsGetNames`
	- `_LOCalc_RangePivotFieldsDataGetList` --> `_LOCalc_RangePivotFieldsDataGetNames`
	- `_LOCalc_RangePivotFieldsFiltersGetList` --> `_LOCalc_RangePivotFieldsFiltersGetNames`
	- `_LOCalc_RangePivotFieldsGetList` --> `_LOCalc_RangePivotFieldsGetNames`
	- `_LOCalc_RangePivotFieldsRowsGetList` --> `_LOCalc_RangePivotFieldsRowsGetNames`
	- `_LOCalc_RangePivotFieldsUnusedGetList` --> `_LOCalc_RangePivotFieldsUnusedGetNames`
	- `_LOCalc_RangePivotsGetList` --> `_LOCalc_RangePivotsGetNames`
- Renamed Functions to be consistent when testing if a thing exists:
	- `_LOCalc_DocHasSheetName` --> `_LOCalc_SheetExists`
	- `_LOCalc_RangeDatabaseHasByName` --> `_LOCalc_RangeDatabaseExists`
	- `_LOCalc_RangeNamedHasByName` --> `_LOCalc_RangeNamedExists`
- Some functions would return an integer instead of an empty Array when no results were present when retrieving array of names or objects, this has been changed to return an empty array:
	- _LOCalc_CellStylesGetNames
	- _LOCalc_PageStylesGetNames
- Modified `_LOCalc_DocPrintersAltGetNames` @Extended value when retrieving the default printer name, @Extended is now 1, instead of 2.
- `_LOCalc_DocRedoGetAllActionTitles` now returns the number of results in @Extended value.
- `_LOCalc_DocUndoGetAllActionTitles` now returns the number of results in @Extended value.
- Made $oDoc parameter for `_LOCalc_FontExists` optional. This will affect the parameters and error return values of the following functions:
	- __LOCalc_CellFont
	- _LOCalc_CellFont
	- _LOCalc_CellStyleFont
	- _LOCalc_FontExists
	- _LOCalc_TextCursorFont
- Renamed more Functions to be consistent when retrieving arrays of names or objects:
	- `_LOCalc_FormatKeyList` --> `_LOCalc_FormatKeysGetList`
- Made $oDoc Parameter optional for `_LOCalc_FontsGetNames`.
- Added count of number of results for `_LOCalc_DocConnect`, connect-all and partial name search when more than one result is present.
- Removed _ArrayDisplay from most examples.
- Corrected wrong usage of `$__LO_STATUS_INIT_ERROR`, and renumbered `$__LO_STATUS_PROCESSING_ERROR`, the following functions were affected:
	- _LOCalc_CommentAreaGradient
	- _LOCalc_CommentAreaTransparencyGradient
	- _LOCalc_CommentPosition
	- _LOCalc_CommentSize
- `_LOCalc_DocRedoCurActionTitle` to only have one Success return, either with an empty String or the Current Redo Action Title.
- `_LOCalc_DocUndoCurActionTitle` to only have one Success return, either with an empty String or the Current Undo Action Title.

### Removed

- __LOCalc_VarsAreDefault
- `LibreOfficeCalc_Font` file, combined functions into `LibreOfficeCalc_Helper`.
- `$__LO_STATUS_DOC_ERROR` Error Constant and renumber all after errors.
- Centralized some internal functions. Thus removing the following individual Functions:
	- __LOCalc_ArrayFill
	- __LOCalc_AddTo1DArray
	- __LOCalc_CreateStruct
	- __LOCalc_IntIsBetween
	- __LOCalc_NumIsBetween
	- __LOCalc_SetPropertyValue
	- __LOCalc_UnitConvert
	- __LOCalc_VarsAreNull
	- __LOCalc_VersionCheck
- Centralized some Helper functions. Thus removing the following individual Functions:
	- _LOCalc_ConvertColorFromLong
	- _LOCalc_ConvertColorToLong
	- _LOCalc_ConvertFromMicrometer
	- _LOCalc_ConvertToMicrometer
	- _LOCalc_PathConvert
	- _LOCalc_VersionGet
- Centralized some Constants. Thus removing the following individual Constants:
	- $LOC_PATHCONV_*
	- $LOC_COLOR_*

> [!NOTE]
> **LibreOfficeImpress**

### Added

- Main Impress File
	- LibreOfficeImpress.au3
- Individual Impress Element Files
	- LibreOfficeImpress_Constants.au3
	- LibreOfficeImpress_Doc.au3
	- LibreOfficeImpress_DrawShapes.au3
	- LibreOfficeImpress_Helper.au3
	- LibreOfficeImpress_Internal.au3
	- LibreOfficeImpress_Slides.au3
- Constants
	- $LOI_ALIGN_VERT_*
	- $LOI_ANCHOR_*
	- $LOI_ANIMATION_DIR_*
	- $LOI_ANIMATION_TYPE_*
	- $LOI_AREA_FILL_STYLE_*
	- $LOI_DRAWSHAPE_CONNECTOR_TYPE_*
	- $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_*
	- $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_*
	- $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_*
	- $LOI_DRAWSHAPE_LINE_ARROW_TYPE_*
	- $LOI_DRAWSHAPE_LINE_CAP_*
	- $LOI_DRAWSHAPE_LINE_JOINT_*
	- $LOI_DRAWSHAPE_LINE_STYLE_*
	- $LOI_DRAWSHAPE_POINT_TYPE_*
	- $LOI_DRAWSHAPE_TYPE_*
	- $LOI_GRAD_NAME_*
	- $LOI_GRAD_TYPE_*
	- $LOI_ORIENT_HORI_*
	- $LOI_ORIENT_VERT_*
	- $LOI_RELATIVE_*
	- $LOI_SHAPE_TYPE_*
	- $LOI_TEXT_ALIGN_HORI_*
	- $LOI_TEXT_ALIGN_VERT_*
	- $LOI_TEXT_ANCHOR_*
	- $LOI_ZOOMTYPE_*
- Document Functions
	- _LOImpress_DocClose
	- _LOImpress_DocConnect
	- _LOImpress_DocCreate
	- _LOImpress_DocExecuteDispatch
	- _LOImpress_DocExport
	- _LOImpress_DocGetName
	- _LOImpress_DocGetPath
	- _LOImpress_DocHasPath
	- _LOImpress_DocIsActive
	- _LOImpress_DocIsModified
	- _LOImpress_DocIsReadOnly
	- _LOImpress_DocMaximize
	- _LOImpress_DocMinimize
	- _LOImpress_DocOpen
	- _LOImpress_DocPosAndSize
	- _LOImpress_DocRedo
	- _LOImpress_DocRedoClear
	- _LOImpress_DocRedoCurActionTitle
	- _LOImpress_DocRedoGetAllActionTitles
	- _LOImpress_DocRedoIsPossible
	- _LOImpress_DocSave
	- _LOImpress_DocSaveAs
	- _LOImpress_DocToFront
	- _LOImpress_DocUndo
	- _LOImpress_DocUndoActionBegin
	- _LOImpress_DocUndoActionEnd
	- _LOImpress_DocUndoClear
	- _LOImpress_DocUndoCurActionTitle
	- _LOImpress_DocUndoGetAllActionTitles
	- _LOImpress_DocUndoIsPossible
	- _LOImpress_DocUndoReset
	- _LOImpress_DocVisible
	- _LOImpress_DocZoom
- Drawing Shape functions
	- _LOImpress_DrawShapeAreaColor
	- _LOImpress_DrawShapeAreaFillStyle
	- _LOImpress_DrawShapeAreaGradient
	- _LOImpress_DrawShapeAreaGradientMulticolor
	- _LOImpress_DrawShapeAreaTransparency
	- _LOImpress_DrawShapeAreaTransparencyGradient
	- _LOImpress_DrawShapeAreaTransparencyGradientMulti
	- _LOImpress_DrawShapeConnectorModify
	- _LOImpress_DrawShapeConnectorSettings
	- _LOImpress_DrawShapeDelete
	- _LOImpress_DrawShapeDimensionSettings
	- _LOImpress_DrawShapeDimensionTextAnimation
	- _LOImpress_DrawShapeDimensionTextColumns
	- _LOImpress_DrawShapeDimensionTextSettings
	- _LOImpress_DrawShapeExists
	- _LOImpress_DrawShapeGetType
	- _LOImpress_DrawShapeInsert
	- _LOImpress_DrawShapeLineArrowStyles
	- _LOImpress_DrawShapeLineProperties
	- _LOImpress_DrawShapeName
	- _LOImpress_DrawShapePointsAdd
	- _LOImpress_DrawShapePointsGetCount
	- _LOImpress_DrawShapePointsModify
	- _LOImpress_DrawShapePointsRemove
	- _LOImpress_DrawShapePosition
	- _LOImpress_DrawShapeRotateSlant
	- _LOImpress_DrawShapeText
	- _LOImpress_DrawShapeTextboxCreateTextCursor
	- _LOImpress_DrawShapeTypePosition
	- _LOImpress_DrawShapeTypeSize
- Helper Functions
	- _LOImpress_ComError_UserFunction
	- _LOImpress_CursorInsertString
	- _LOImpress_FontExists
	- _LOImpress_FontsGetNames
	- _LOImpress_GradientMulticolorAdd
	- _LOImpress_GradientMulticolorDelete
	- _LOImpress_GradientMulticolorModify
	- _LOImpress_TransparencyGradientMultiAdd
	- _LOImpress_TransparencyGradientMultiDelete
	- _LOImpress_TransparencyGradientMultiModify
- Internal Functions
	- __LOImpress_ColorRemoveAlpha
	- __LOImpress_CreatePoint
	- __LOImpress_DrawShape_CreateArrow
	- __LOImpress_DrawShape_CreateBasic
	- __LOImpress_DrawShape_CreateCallout
	- __LOImpress_DrawShape_CreateFlowchart
	- __LOImpress_DrawShape_CreateLine
	- __LOImpress_DrawShape_CreateStars
	- __LOImpress_DrawShape_CreateSymbol
	- __LOImpress_DrawShape_GetCustomType
	- __LOImpress_DrawShapeArrowStyleName
	- __LOImpress_DrawShapeLineStyleName
	- __LOImpress_DrawShapePointGetSettings
	- __LOImpress_DrawShapePointModify
	- __LOImpress_FilterNameGet
	- __LOImpress_GetShapeName
	- __LOImpress_GradientNameInsert
	- __LOImpress_GradientPresets
	- __LOImpress_InternalComErrorHandler
	- __LOImpress_ShapeGetType
	- __LOImpress_TransparencyGradientConvert
	- __LOImpress_TransparencyGradientNameInsert
- Slide Functions
	- _LOImpress_SlideAdd
	- _LOImpress_SlideAreaColor
	- _LOImpress_SlideAreaFillStyle
	- _LOImpress_SlideAreaGradient
	- _LOImpress_SlideAreaTransparency
	- _LOImpress_SlideAreaTransparencyGradient
	- _LOImpress_SlideCopy
	- _LOImpress_SlideCurrent
	- _LOImpress_SlideDeleteByIndex
	- _LOImpress_SlideDeleteByObj
	- _LOImpress_SlideGetByIndex
	- _LOImpress_SlideMove
	- _LOImpress_SlidesGetCount
	- _LOImpress_SlideShapesGetList

### Removed

- Centralized some internal functions. Thus removing the following individual Functions:
	- __LOImpress_ArrayFill
	- __LOImpress_AddTo1DArray
	- __LOImpress_CreateStruct
	- __LOImpress_IntIsBetween
	- __LOImpress_NumIsBetween
	- __LOImpress_SetPropertyValue
	- __LOImpress_UnitConvert
	- __LOImpress_VarsAreNull
	- __LOImpress_VersionCheck
- Centralized some Helper functions. Thus removing the following individual Functions:
	- _LOImpress_ConvertColorFromLong
	- _LOImpress_ConvertColorToLong
	- _LOImpress_ConvertFromMicrometer
	- _LOImpress_ConvertToMicrometer
	- _LOImpress_PathConvert
	- _LOImpress_VersionGet
- Centralized some Constants. Thus removing the following individual Constants:
	- $LOI_PATHCONV_*
	- $LOI_COLOR_*

> [!NOTE]
> **LibreOfficeWriter**

### Added

- More Undo Functions and Examples
	- _LOWriter_DocRedoClear
	- _LOWriter_DocUndoActionBegin
	- _LOWriter_DocUndoActionEnd
	- _LOWriter_DocUndoClear
	- _LOWriter_DocUndoReset
- Shape Functions and Examples
	- _LOWriter_DocHasShapeName
	- _LOWriter_ShapeAreaColor
	- _LOWriter_ShapeAreaGradient
	- _LOWriter_ShapeDelete
	- _LOWriter_ShapeGetAnchor
	- _LOWriter_ShapeGetObjByName
	- _LOWriter_ShapeGetType
	- _LOWriter_ShapeInsert
	- _LOWriter_ShapeLineArrowStyles
	- _LOWriter_ShapeLineProperties
	- _LOWriter_ShapeName
	- _LOWriter_ShapePointsAdd
	- _LOWriter_ShapePointsGetCount
	- _LOWriter_ShapePointsModify
	- _LOWriter_ShapePointsRemove
	- _LOWriter_ShapePosition
	- _LOWriter_ShapeRotateSlant
	- _LOWriter_ShapesGetNames
	- _LOWriter_ShapeTextBox
	- _LOWriter_ShapeTransparency
	- _LOWriter_ShapeTransparencyGradient
	- _LOWriter_ShapeTypePosition
	- _LOWriter_ShapeTypeSize
	- _LOWriter_ShapeWrap
	- _LOWriter_ShapeWrapOptions
	- __LOWriter_CreatePoint
	- __LOWriter_GetShapeName
	- __LOWriter_Shape_CreateArrow
	- __LOWriter_Shape_CreateBasic
	- __LOWriter_Shape_CreateCallout
	- __LOWriter_Shape_CreateFlowchart
	- __LOWriter_Shape_CreateLine
	- __LOWriter_Shape_CreateStars
	- __LOWriter_Shape_CreateSymbol
	- __LOWriter_Shape_GetCustomType
	- __LOWriter_ShapeArrowStyleName
	- __LOWriter_ShapeLineStyleName
	- __LOWriter_ShapePointGetSettings
	- __LOWriter_ShapePointModify
- Shape Point Constants in `LibreOfficeWriter_Constants`. `$LOW_SHAPE_POINT_TYPE_*`
- Standard Format Key retrieval function `_LOWriter_FormatKeyGetStandard`
- Alpha Removal function `__LOWriter_ColorRemoveAlpha`.
- _LOWriter_FrameAreaFillStyle.
- _LOWriter_FrameStyleAreaFillStyle
- _LOWriter_ImageAreaFillStyle
- _LOWriter_PageStyleAreaFillStyle
- _LOWriter_PageStyleFooterAreaFillStyle
- _LOWriter_PageStyleHeaderAreaFillStyle
- _LOWriter_ShapeAreaFillStyle
- Line numbers to Example Error messages.
- Top-Most attribute to Example message boxes.
- Selection set and get functions
	- _LOWriter_DocSelectionGet
	- _LOWriter_DocSelectionSet
- `__LOWriter_NumRuleCreateMap` for modifying Numbering Rules more efficiently.
- LibreOffice SDK/API Constant names to constants.
- Form/Form Control Constants
	- $LOW_FORM_CON_BORDER_*
	- $LOW_FORM_CON_CHKBX_STATE_*
	- $LOW_FORM_CON_DATE_FRMT_*
	- $LOW_FORM_CON_IMG_ALIGN_*
	- $LOW_FORM_CON_IMG_BTN_SCALE_*
	- $LOW_FORM_CON_MOUSE_SCROLL_*
	- $LOW_FORM_CON_PUSH_CMD_*
	- $LOW_FORM_CON_SCROLL_*
	- $LOW_FORM_CON_SOURCE_TYPE_*
	- $LOW_FORM_CON_TIME_FRMT_*
	- $LOW_FORM_CON_TXT_TYPE_*
	- $LOW_FORM_CON_TYPE_*
	- $LOW_FORM_CONTENT_TYPE_*
	- $LOW_FORM_CYCLE_MODE_*
	- $LOW_FORM_NAV_BAR_MODE_*
	- $LOW_FORM_SUBMIT_ENCODING_*
	- $LOW_FORM_SUBMIT_METHOD_*
- Internal Form Functions
	- __LOWriter_FormConIdentify
	- __LOWriter_FormConGetObj
	- __LOWriter_FormConSetGetFontDesc
- Form functions
	- _LOWriter_DocFormSettings
	- _LOWriter_FontDescCreate
	- _LOWriter_FontDescEdit
	- _LOWriter_FormAdd
	- _LOWriter_FormConCheckBoxData
	- _LOWriter_FormConCheckBoxGeneral
	- _LOWriter_FormConCheckBoxState
	- _LOWriter_FormConComboBoxData
	- _LOWriter_FormConComboBoxGeneral
	- _LOWriter_FormConComboBoxValue
	- _LOWriter_FormConCurrencyFieldData
	- _LOWriter_FormConCurrencyFieldGeneral
	- _LOWriter_FormConCurrencyFieldValue
	- _LOWriter_FormConDateFieldData
	- _LOWriter_FormConDateFieldGeneral
	- _LOWriter_FormConDateFieldValue
	- _LOWriter_FormConDelete
	- _LOWriter_FormConFileSelFieldGeneral
	- _LOWriter_FormConFileSelFieldValue
	- _LOWriter_FormConFormattedFieldData
	- _LOWriter_FormConFormattedFieldGeneral
	- _LOWriter_FormConFormattedFieldValue
	- _LOWriter_FormConGetParent
	- _LOWriter_FormConGroupBoxGeneral
	- _LOWriter_FormConImageButtonGeneral
	- _LOWriter_FormConImageControlData
	- _LOWriter_FormConImageControlGeneral
	- _LOWriter_FormConInsert
	- _LOWriter_FormConLabelGeneral
	- _LOWriter_FormConListBoxData
	- _LOWriter_FormConListBoxGeneral
	- _LOWriter_FormConListBoxGetCount
	- _LOWriter_FormConListBoxSelection
	- _LOWriter_FormConNavBarGeneral
	- _LOWriter_FormConNumericFieldData
	- _LOWriter_FormConNumericFieldGeneral
	- _LOWriter_FormConNumericFieldValue
	- _LOWriter_FormConOptionButtonData
	- _LOWriter_FormConOptionButtonGeneral
	- _LOWriter_FormConOptionButtonState
	- _LOWriter_FormConPatternFieldData
	- _LOWriter_FormConPatternFieldGeneral
	- _LOWriter_FormConPatternFieldValue
	- _LOWriter_FormConPosition
	- _LOWriter_FormConPushButtonGeneral
	- _LOWriter_FormConPushButtonState
	- _LOWriter_FormConsGetList
	- _LOWriter_FormConSize
	- _LOWriter_FormConTableConCheckBoxData
	- _LOWriter_FormConTableConCheckBoxGeneral
	- _LOWriter_FormConTableConColumnAdd
	- _LOWriter_FormConTableConColumnDelete
	- _LOWriter_FormConTableConColumnsGetList
	- _LOWriter_FormConTableConComboBoxData
	- _LOWriter_FormConTableConComboBoxGeneral
	- _LOWriter_FormConTableConCurrencyFieldData
	- _LOWriter_FormConTableConCurrencyFieldGeneral
	- _LOWriter_FormConTableConDateFieldData
	- _LOWriter_FormConTableConDateFieldGeneral
	- _LOWriter_FormConTableConFormattedFieldData
	- _LOWriter_FormConTableConFormattedFieldGeneral
	- _LOWriter_FormConTableConGeneral
	- _LOWriter_FormConTableConListBoxData
	- _LOWriter_FormConTableConListBoxGeneral
	- _LOWriter_FormConTableConNumericFieldData
	- _LOWriter_FormConTableConNumericFieldGeneral
	- _LOWriter_FormConTableConPatternFieldData
	- _LOWriter_FormConTableConPatternFieldGeneral
	- _LOWriter_FormConTableConTextBoxData
	- _LOWriter_FormConTableConTextBoxGeneral
	- _LOWriter_FormConTableConTimeFieldData
	- _LOWriter_FormConTableConTimeFieldGeneral
	- _LOWriter_FormConTextBoxCreateTextCursor
	- _LOWriter_FormConTextBoxData
	- _LOWriter_FormConTextBoxGeneral
	- _LOWriter_FormConTimeFieldData
	- _LOWriter_FormConTimeFieldGeneral
	- _LOWriter_FormConTimeFieldValue
	- _LOWriter_FormDelete
	- _LOWriter_FormGetObjByIndex
	- _LOWriter_FormParent
	- _LOWriter_FormPropertiesData
	- _LOWriter_FormPropertiesGeneral
	- _LOWriter_FormsGetCount
	- _LOWriter_FormsGetList
- Multicolor Gradient support, adding the following functions:
	1. Helper Functions:
		- _LOWriter_GradientMulticolorAdd
		- _LOWriter_GradientMulticolorDelete
		- _LOWriter_GradientMulticolorModify
		- _LOWriter_TransparencyGradientMultiAdd
		- _LOWriter_TransparencyGradientMultiDelete
		- _LOWriter_TransparencyGradientMultiModify
	2. Frame Functions:
		- _LOWriter_FrameAreaGradientMulticolor
		- _LOWriter_FrameAreaTransparencyGradientMulti
		- _LOWriter_FrameStyleAreaGradientMulticolor
		- _LOWriter_FrameStyleAreaTransparencyGradientMulti
	3. Image Functions:
		- _LOWriter_ImageAreaGradientMulticolor
		- _LOWriter_ImageAreaTransparencyGradientMulti
	4. PageStyle Functions:
		- _LOWriter_PageStyleAreaGradientMulticolor
		- _LOWriter_PageStyleAreaTransparencyGradientMulti
		- _LOWriter_PageStyleFooterAreaGradientMulticolor
		- _LOWriter_PageStyleFooterAreaTransparencyGradientMulti
		- _LOWriter_PageStyleHeaderAreaGradientMulticolor
		- _LOWriter_PageStyleHeaderAreaTransparencyGradientMulti
	5. Shape Functions:
		- _LOWriter_ShapeAreaGradientMulticolor
		- _LOWriter_ShapeAreaTransparencyGradientMulti

### Changed

- ReArranged Parameters in PageBreak functions to be more logically ordered. The previous order was:

	> $iBreakType = Null, $iPgNumOffSet = Null, $sPageStyle = Null

	However the functions stated:

	> _Break Type_ must be set **before** _Page Style_ will be able to be set, and _page style_ needs set **before** _$iPgNumOffSet_ can be set.

	Thus the parameters have been rearranged to:

	> $iBreakType = Null, $sPageStyle = Null, $iPgNumOffSet

	Affected functions are:
	- _LOWriter_DirFrmtParPageBreak
	- _LOWriter_ParStylePageBreak
	- __LOWriter_ParPageBreak

- Undo/Redo Action Title retrieval functions to no longer return an error when Undo/Redo Action Titles are unavailable, instead they  return an empty string or array, respectively.
	- _LOWriter_DocRedoCurActionTitle;
	- _LOWriter_DocRedoGetAllActionTitles;
	- _LOWriter_DocUndoGetAllActionTitles;
	- _LOWriter_DocUndoCurActionTitle.
- Minor documentation adjustments.
- Moved Search Descriptor functions from `LibreOfficeWriter.au3` to `LibreOfficeWriter_Helper.au3`.
	- _LOWriter_SearchDescriptorCreate
	- _LOWriter_SearchDescriptorModify
	- _LOWriter_SearchDescriptorSimilarityModify
- Default value of `$__LOWCONST_SLEEP_DIV` from 15 to 0.
- Sorted Constants in LibreOfficeWriter_Constants alphabetically.
- Renamed `$LOW_FIELDADV_TYPE_*` Constants to `$LOW_FIELD_ADV_TYPE_*` to match formatting of other Field Type Constants.
- Renamed `$LibreOfficeWriter_DirectFormating.au3` to `LibreOfficeWriter_DirectFormatting.au3` (misspelling correction)
- Renamed Parameter $iBorder for Gradient Functions to $iTransitionStart for clarity of which setting it corresponds to in the L.O. UI.
	- _LOWriter_FrameAreaGradient
	- _LOWriter_FrameTransparencyGradient
	- _LOWriter_FrameStyleAreaGradient
	- _LOWriter_FrameStyleTransparencyGradient
	- _LOWriter_ImageAreaGradient
	- _LOWriter_ImageAreaTransparencyGradient
	- _LOWriter_PageStyleAreaGradient
	- _LOWriter_PageStyleFooterAreaGradient
	- _LOWriter_PageStyleFooterTransparencyGradient
	- _LOWriter_PageStyleHeaderAreaGradient
	- _LOWriter_PageStyleHeaderTransparencyGradient
	- _LOWriter_PageStyleTransparencyGradient
	- _LOWriter_ShapeAreaGradient
	- _LOWriter_ShapeTransparencyGradient
- `_LOWriter_DocOpen` now uses a different method for connecting to an already open document, as the previous method could potentially cause errors.
- Renamed `$__LOWCONST_FILL_STYLE_*` Constant to `$LOW_AREA_FILL_STYLE_*`.
- `__LOWriter_IntIsBetween` to accept only a minimum value. Also optimized it.
	- Modified function usage to match changes.
- Improved the processing speed for modifying Numbering Rules by using Maps.
- Renamed Functions to be consistent when retrieving arrays of names or objects:
	- `_LOWriter_DocBookmarksList` --> `_LOWriter_DocBookmarksGetNames`
	- `_LOWriter_DocEnumPrinters` --> `_LOWriter_DocPrintersGetNames`
	- `_LOWriter_DocEnumPrintersAlt` --> `_LOWriter_DocPrintersAltGetNames`
	- `_LOWriter_FieldSetVarMasterList` --> `_LOWriter_FieldSetVarMastersGetNames`
	- `_LOWriter_FieldSetVarMasterListFields` --> `_LOWriter_FieldSetVarMasterFieldsGetList`
	- `_LOWriter_FontsList` --> `_LOWriter_FontsGetNames`
	- `_LOWriter_TableGetCellNames` --> `_LOWriter_TableCellsGetNames`
- Renamed Functions to be consistent when testing if a thing exists:
	- `_LOWriter_DocBookmarksHasName` --> `_LOWriter_DocBookmarkExists`
	- `_LOWriter_DocHasFrameName` --> `_LOWriter_FrameExists`
	- `_LOWriter_DocHasImageName` --> `_LOWriter_ImageExists`
	- `_LOWriter_DocHasShapeName` --> `_LOWriter_ShapeExists`
	- `_LOWriter_DocHasTableName` --> `_LOWriter_TableExists`
- Some functions would return an integer instead of an empty Array when no results were present when retrieving array of names or objects, this has been changed to return an empty array:
	- _LOWriter_CharStylesGetNames
	- _LOWriter_DocBookmarksGetNames
	- _LOWriter_EndnotesGetList
	- _LOWriter_FieldRefMarkList
	- _LOWriter_FieldSetVarMasterFieldsGetList
	- _LOWriter_FootnotesGetList
	- _LOWriter_FrameStylesGetNames
	- _LOWriter_ImagesGetNames
	- _LOWriter_NumStylesGetNames
	- _LOWriter_PageStylesGetNames
	- _LOWriter_ParStylesGetNames
	- _LOWriter_TablesGetNames
- Modified `_LOWriter_DocPrintersAltGetNames` @Extended value when retrieving the default printer name, @Extended is now 1, instead of 2.
- `_LOWriter_DocRedoGetAllActionTitles` now returns the number of results in @Extended value.
- `_LOWriter_DocUndoGetAllActionTitles` now returns the number of results in @Extended value.
- Made $oDoc parameter for `_LOWriter_FontExists` optional. This will affect the parameters and error return values of the following functions:
	- __LOWriter_CharFont
	- _LOWriter_CharStyleFont
	- _LOWriter_DirFrmtFont
	- _LOWriter_FindFormatModifyFont
	- _LOWriter_FontExists
	- _LOWriter_FontDescCreate
	- _LOWriter_FontDescEdit
	- _LOWriter_ParStyleFont
- Modified `_LOWriter_ShapesGetNames` to return Shape Type instead of Implementation name.
- Renamed `__LOWriter_ControlSetGetFontDesc` --> `__LOWriter_FormConSetGetFontDesc` to match other Form control function names.
- Renamed all internal and user `_LOWriter_FormControl*` functions to `_LOWriter_FormCon*` for brevity.
- Renamed all `$LOW_FORM_CONTROL_*` constants to `$LOW_FORM_CON_*` for brevity.
- Add $iBorderColor to `_LOWriter_FormTableConGeneral` parameters. Update function and example as necessary.
- Renamed more Functions to be consistent when retrieving arrays of names or objects:
	- `_LOWriter_DateFormatKeyList` --> `_LOWriter_DateFormatKeysGetList`
	- `_LOWriter_DirFrmtParTabStopList` --> `_LOWriter_DirFrmtParTabStopsGetList`
	- `_LOWriter_FieldRefMarkList` --> `_LOWriter_FieldRefMarksGetNames`
	- `_LOWriter_FormatKeyList` --> `_LOWriter_FormatKeysGetList`
	- `_LOWriter_ParStyleTabStopList` --> `_LOWriter_ParStyleTabStopsGetList`
	- `__LOWriter_ParTabStopList` --> `__LOWriter_ParTabStopsGetList`
- Made $oDoc Parameter optional for `_LOWriter_FontsGetNames`.
- Added count of number of results for `_LOWriter_DocConnect`, connect-all and partial name search when more than one result is present.
- Removed _ArrayDisplay from most examples.
- Added checks to some Form functions whether the Document is ReadOnly, which would occur if the Form Document is opened in Viewing mode.
- Modified all DocFind/DocReplace functions to use null to skip the Format Parameters instead of an empty array.
- Rearranged Parameters in `_LOWriter_DocFindAllInRange`. $oRange and $atFormat are now in reverse order.
- Rewrote `_LOWriter_DocReplaceAllInRange` to work better, and also to eliminate buggy `__LOWriter_RegExpConvert` function.
- Made `_LOWriter_DocReplaceAll` and `_LOWriter_DocReplaceAllInRange` return the number of replacements made, instead of setting @Extended.
- Made `_LOWriter_DocFindAll` and `_LOWriter_DocFindAllInRange` always return an Array, instead of "1" when no results were found.
- Renamed Transparency functions to be consistent and kept together with other background functions:
	- `_LOWriter_FrameStyleTransparency` --> `_LOWriter_FrameStyleAreaTransparency`
	- `_LOWriter_FrameStyleTransparencyGradient` --> `_LOWriter_FrameStyleAreaTransparencyGradient`
	- `_LOWriter_FrameTransparency` --> `_LOWriter_FrameAreaTransparency`
	- `_LOWriter_FrameTransparencyGradient` --> `_LOWriter_FrameAreaTransparencyGradient`
	- `_LOWriter_PageStyleFooterTransparency` --> `_LOWriter_PageStyleFooterAreaTransparency`
	- `_LOWriter_PageStyleFooterTransparencyGradient` --> `_LOWriter_PageStyleFooterAreaTransparencyGradient`
	- `_LOWriter_PageStyleHeaderTransparency` --> `_LOWriter_PageStyleHeaderAreaTransparency`
	- `_LOWriter_PageStyleHeaderTransparencyGradient` --> `_LOWriter_PageStyleHeaderAreaTransparencyGradient`
	- `_LOWriter_PageStyleTransparency` --> `_LOWriter_PageStyleAreaTransparency`
	- `_LOWriter_PageStyleTransparencyGradient` --> `_LOWriter_PageStyleAreaTransparencyGradient`
	- `_LOWriter_ShapeTransparency` --> `_LOWriter_ShapeAreaTransparency`
	- `_LOWriter_ShapeTransparencyGradient` --> `_LOWriter_ShapeAreaTransparencyGradient`
- `_LOWriter_DocRedoCurActionTitle` to only have one Success return, either with an empty String or the Current Redo Action Title.
- `_LOWriter_DocUndoCurActionTitle` to only have one Success return, either with an empty String or the Current Undo Action Title.

### Fixed

- Missing error return value descriptions in:
	- _LOWriter_FootnoteGetAnchor
	- _LOWriter_EndnoteGetAnchor
	- _LOWriter_FrameGetAnchor
	- _LOWriter_ImageGetAnchor
- Missing Includes in `LibreOfficeWriter` file.
- Ternary Operators missing Parenthesis in responses.
- Export extension was incorrect, "jjpe" --> "jpe"
- Spell Checked the comments and Headers.
- Spell Checked Examples
- Examples on Error now close any documents opened by the example.
- Missing parameters in Null keyword check for `__LOWriter_TableBorder`, causing errors in setting Table Horizontal and Vertical borders singly.
- Wrong error value in `_LOWriter_CursorGetStatus` for error while determining cursor type.
- Missing Parameters in some headers or Syntax header.
- Missing parameter type checks in some functions.
- Missing ExitLoop in `__LOWriter_FieldsGetList` causing unnecessary looping.
- Removed unused variables and parameters in some functions. Affected functions are as follows:
	- `_LOWriter_DocBookmarkModify` -- removed $oDoc parameter.
	- `_LOWriter_FieldDelete` -- removed $oDoc parameter.
	- `_LOWriter_FieldDocInfoEditTimeModify` -- removed $oDoc parameter.
	- `_LOWriter_FieldSetVarMasterFieldsGetList` -- removed $oDoc parameter.
	- `_LOWriter_DateFormatKeyDelete` -- removed internal variable.
	- `_LOWriter_FormatKeyDelete` -- removed internal variable.
	- `_LOWriter_ImageInsert` -- removed internal variable, now uses ViewCursor directly to insert an Image.
	- `_LOWriter_NumStyleSetLevel` -- removed $oDoc parameter.
- `_LOWriter_DocHyperlinkInsert` -- $bOverwrite parameter was not used in function.
- Transparency causing Color values to be returned that including the Alpha value, causing potentially unexpected results.
- Missing Data Type in `_LOWriter_CursorGetDataType` example.
- `_LOWriter_DocCreateTextCursor` would throw an error when creating a Text Cursor at the ViewCursor position.
- When any `FieldsGetList` functions were supposed to return a single dimension array, a two dimensional array was being returned.
- Incorrect usage of ObjEvent.
- `_LOWriter_DocCreate` would return if there was an error creating a property, instead of increasing the error count.
- `_LOWriter_DocCreate` could potentially return a Base Form document, as they have identical Service names.
- `_LOWriter_DocConnect` could potentially return a Base Form document, as they have identical Service names.
- `__LOWriter_TransparencyGradientConvert` would return a wrong Transparency value for certain percentages.
- LibreOffice version 7.6 introduced a new setting for gradients, which broke all gradient functions I had made. Implemented a fix to work with both the new version and the old.
- `_LOWriter_DocPrintMiscSettings` #2 example no longer worked after a change to how one of the functions worked.
- `__LOWriter_GetShapeName` had an error where a COM Error would be triggered each time it was called.
- Added missing example. `_LOWriter_ImageExists`
- Backwards Parameters in VarsAreNull in the following functions:
	- _LOWriter_FieldDateTimeModify
	- _LOWriter_FieldFileNameModify
- Add missing parameter to VarsAreNull in `_LOWriter_DocFormSettings`.
- `_LOWriter_DocCreate` not finding a blank open document to connect to, if available, due to reversed logical operator.

### Removed

- `$LOW_FIELD_TYPE_URL` Constant. -- "com.sun.star.text.TextField.URL" is a Calc-only Field type.
- `__LOWriter_NumStyleRetrieve` function as it is no longer needed.
- Remove `LibreOfficeWriter_Font` file, and merge functions into `LibreOfficeWriter_Helper`
- `$__LO_STATUS_DOC_ERROR` Error Constant and renumber all after errors.
- __LOWriter_RegExpConvert.
- __LOWriter_AddTo2DArray.
- __LOWriter_VarsAreDefault
- Centralized some internal functions. Thus removing the following individual Functions:
	- __LOWriter_ArrayFill
	- __LOWriter_AddTo1DArray
	- __LOWriter_CreateStruct
	- __LOWriter_IntIsBetween
	- __LOWriter_NumIsBetween
	- __LOWriter_SetPropertyValue
	- __LOWriter_UnitConvert
	- __LOWriter_VarsAreNull
	- __LOWriter_VersionCheck
- Centralized some Helper functions. Thus removing the following individual Functions:
	- _LOWriter_ConvertColorFromLong
	- _LOWriter_ConvertColorToLong
	- _LOWriter_ConvertFromMicrometer
	- _LOWriter_ConvertToMicrometer
	- _LOWriter_PathConvert
	- _LOWriter_VersionGet
- Centralized some Constants. Thus removing the following individual Constants:
	- $LOW_PATHCONV_*
	- $LOW_COLOR_*

[To Top](#releases)

## [0.9.1] - 2023-10-28

> [!NOTE]
> **LibreOfficeWriter**

### Changed

- Minor documentation adjustments.

[To Top](#releases)

## [0.9.0] - 2023-10-28

> [!NOTE]
> **LibreOfficeWriter**

### Added

- directive for Au3Check to each UDF file branch. (@mLipok)
- Image functions. (@donnyh13)
	- _ImageAreaColor,
	- _ImageAreaGradient,
	- _ImageAreaTransparency,
	- _ImageAreaTransparencyGradient,
	- _ImageBorderColor,
	- _ImageBorderPadding,
	- _ImageBorderStyle,
	- _ImageBorderWidth,
	- _ImageColorAdjust,
	- _ImageCrop,
	- _ImageDelete,
	- _ImageGetAnchor,
	- _ImageGetObjByName,
	- _ImageHyperlink,
	- _ImageInsert,
	- _ImageModify,
	- _ImageOptions,
	- _ImageOptionsName,
	- _ImageReplace,
	- _ImagesGetNames,
	- _ImageShadow,
	- _ImageSize,
	- _ImageTransparency,
	- _ImageTypePosition,
	- _ImageTypeSize,
	- _ImageWrap,
	- _ImageWrapOptions

### Changed

- Original LibreOffice UDF file split into individual elements, per specific usages. (@mLipok.)
	- LibreOfficeWriter_Cell,
	- LibreOfficeWriter_Char,
	- LibreOfficeWriter_Constants,
	- LibreOfficeWriter_Cursor,
	- LibreOfficeWriter_DirectFormatting,
	- LibreOfficeWriter_Doc,
	- LibreOfficeWriter_Field,
	- LibreOfficeWriter_Font,
	- LibreOfficeWriter_FootEndNotes,
	- LibreOfficeWriter_Frame,
	- LibreOfficeWriter_Helper,
	- LibreOfficeWriter_Images,
	- LibreOfficeWriter_Internal,
	- LibreOfficeWriter_Num,
	- LibreOfficeWriter_Page,
	- LibreOfficeWriter_Par,
	- LibreOfficeWriter_Shapes,
	- LibreOfficeWriter_Table.
- Renamed Constants:
	- `$LOW_PAPER_PORTRAIT` --> `$LOW_PAPER_ORIENT_PORTRAIT`
	- `$LOW_PAPER_LANDSCAPE` --> `$LOW_PAPER_ORIENT_LANDSCAPE`
- Examples layout and error checking cleaned up. (@mLipok.)
- Examples now work from their separate folder. (@mLipok.)
- Major editing of Header layout for every function. As well as several typo corrections and wordiness. (@mLipok) & (@donnyh13.)
- Constants are now all located in the separate Constants files. (@mLipok.)
- All files re-processed with TIDY. (@mLipok)

### Fixed

- Errors caused by residual function calls from filling in "Related" section of the header. (@mLipok.)
	- _LOWriter_DocFindAllInRange,
	- _LOWriter_DocGenPropTemplate.

### Project

- @mLipok and @donnyh13 began jointly working on this project. — Thanks to @mLipok for his tireless work cleaning up many things in this UDF.
- All Constants descriptions are moved to the Constants file. (@donnyh13)
- Added documentation and improved CHANGELOG.md. (@donnyh13)
- Added bug_report.md. (@donnyh13)
- Added feature_request.md. (@donnyh13)
- Added PULL_REQUEST_TEMPLATE.md. (@donnyh13)
- Added CODE_OF_CONDUCT.md (@donnyh13)
- Added CONTRIBUTING.md. (@donnyh13)

Thanks @danp2 and @Sven-Seyfert. All above mentioned MD documents were based on adequate documents from <https://github.com/Danp2/au3WebDriver>.

[To Top](#releases)

## [0.0.0.3] - 2023-08-10

> [!NOTE]
> **LibreOfficeWriter**

### Added

- Paragraph Object functions which allows the ability to copy and paste content without using the clipboard quickly. Thanks to user @Heiko for inspiration.
	- _ParObjSelect,
	- _ParObjCopy,
	- _ParObjPaste,
	- _ParObjDelete.
- _DocExecuteDispatch function, adding some shortcuts to certain commands, such as select all, copy/paste content to/from clipboard.
- _DocConvertTextToTable.
- _DocConvertTableToText.
- examples for the new functions.
- "Related" functions section for most function headers.
- Warning to `_ShapesGetNames`, about Images inserted in a document also being called "TextFrames".
- processing error check.
	- _DocInsertString,
	- _DocInsertControlChar.
- a missing Datatype for possible Cursor data position types, `$LOW_CURDATA_HEADER_FOOTER`, previously attempting to insert a content while the insertion point was located in a Header/Footer would have failed. — Thanks to user @Heiko for helping me locate this error. The affected functions are:
	- _CursorGetDataType.
	- _DocCreateTextCursor,
	- _EndnoteInsert,
	- _FootnoteInsert,
	- _TableInsert,

### Changed

- Renamed Name functions and examples for consistency.
	- `_DocListTableNames` --> `_TablesGetNames`
	- `_FramesListNames` --> `_FramesGetNames`
	- `_ShapesListNames` --> `_ShapesGetNames` 	
	- `_ParGetObjects` --> `_ParObjCreateList`.
	- `_ParSectionsGet` --> `_ParObjSectionsGet`.
	- `_TableGetByCursor` --> `_TableGetObjByCursor`.
	- `_TableGetByName` --> `_TableGetObjByName`.
	- `_TableGetCellByCursor` --> `_TableGetCellObjByCursor`.
	- `_TableGetCellByName` --> `_TableGetCellObjByName`.
	- `_TableGetCellByPosition` --> `_TableGetCellObjByPosition`.
- Removed "IsCollapsed" check and error from _DocGetString.
- `_FramesListNames` to have an option to search for Frames listed under shapes.
- `_ShapesGetNames`, Corrected an error that could occur if images are present.

### Fixed

- An error where a COM error would be produced when attempting to insert a string or control character in certain data types. — Thanks to user @Heiko for helping me locate this error.
	- _DocInsertControlChar,
	- _DocInsertString.

[To Top](#releases)

## [0.0.0.2] - 2023-07-16

> [!NOTE]
> **LibreOfficeWriter**

### Added

- UDF version number in the UDF Header.

### Changed

- `_DocReplaceAllInRange` to have two methods of performing a Regular Expression find and replace.
- Removed the if/else block in $atFindFormat parameter checking.
	- _DocReplaceAll,
	- _DocReplaceAllInRange,
	- _DocFindNext,
	- _DocFindAll,
	- _DocFindAllInRange.

### Fixed

- Method for skipping $atFindFormat and $atReplaceFormat, now uses an empty array called in each parameter to skip.
	- _DocReplaceAll,
	- _DocReplaceAllInRange.

### Project

- Updated function documentation to reflect the changes.

[To Top](#releases)

## [0.0.0.1] - 2023-07-02

> [!NOTE]
> **LibreOfficeWriter**

### Added

- Initial UDF Release.

[To Top](#releases)

---

### Legend - Types of changes

- `Added` for new features.
- `Changed` for changes in existing functionality.
- `Deprecated` for soon-to-be removed features.
- `Fixed` for any bug fixes.
- `Removed` for now removed features.
- `Security` in case of vulnerabilities.
- `Project` for documentation or overall project improvements.

[To the top](#changelog)

---

[v0.10.0-Compare]:	https://github.com/mlipok/Au3LibreOffice/compare/0.9.1...main
[v0.9.1-Compare]:	https://github.com/mlipok/Au3LibreOffice/compare/v0.9.0...0.9.1
[v0.9.0-Compare]:	https://github.com/mlipok/Au3LibreOffice/compare/v0.0.0.3...v0.9.0
[v0.0.0.3-Compare]:	https://github.com/donnyh13/Au3LibreOffice/compare/v0.0.0.2...v0.0.0.3
[v0.0.0.2-Compare]:	https://github.com/donnyh13/Au3LibreOffice/compare/v0.0.0.1...v0.0.0.2

[v0.10.0]:	https://github.com/mlipok/Au3LibreOffice
[v0.9.1]:	https://github.com/mlipok/Au3LibreOffice/releases/tag/0.9.1
[v0.9.0]:	https://github.com/mlipok/Au3LibreOffice/releases/tag/v0.9.0
[v0.0.0.3]:	https://github.com/mlipok/Au3LibreOffice/releases/tag/v0.0.0.3
[v0.0.0.2]:	https://github.com/donnyh13/Au3LibreOffice/releases/tag/v0.0.0.2
[v0.0.0.1]:	https://github.com/donnyh13/Au3LibreOffice/releases/tag/v0.0.0.1
