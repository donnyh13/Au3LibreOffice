#####

# Changelog

All notable changes to "Au3LibreOffice" SDK/API will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
This project also adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend---types-of-changes) for further information about the types of changes.

## [?.?.?] - 2024-

**LibreOfficeCalc**

### Added 

- Initial Document Functions and Examples
	- _LOCalc_DocClose
	- _LOCalc_DocConnect
	- _LOCalc_DocCreate
	- _LOCalc_DocExport
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
	- _LOCalc_DocRedo
	- _LOCalc_DocRedoCurActionTitle
	- _LOCalc_DocRedoGetAllActionTitles
	- _LOCalc_DocRedoIsPossible
	- _LOCalc_DocSave
	- _LOCalc_DocSaveAs
	- _LOCalc_DocToFront
	- _LOCalc_DocUndo
	- _LOCalc_DocUndoCurActionTitle
	- _LOCalc_DocUndoGetAllActionTitles
	- _LOCalc_DocUndoIsPossible
	- _LOCalc_DocVisible
	- _LOCalc_DocZoom
- Sheet Functions and Examples
	- _LOCalc_SheetActivate
	- _LOCalc_SheetAdd
	- _LOCalc_SheetCopy
	- _LOCalc_SheetGetActive
	- _LOCalc_SheetGetCellByName
	- _LOCalc_SheetGetCellByPosition
	- _LOCalc_SheetGetObjByName
	- _LOCalc_SheetIsActive
	- _LOCalc_SheetMove
	- _LOCalc_SheetName
	- _LOCalc_SheetRemove
	- _LOCalc_SheetsGetCount
	- _LOCalc_SheetsGetNames
	- _LOCalc_SheetVisible
- Cell Functions and Examples
	- _LOCalc_CellFormula
	- _LOCalc_CellGetType
	- _LOCalc_CellText
	- _LOCalc_CellValue

**LibreOfficeWriter**

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
- Shape Point Constants in LibreOfficeWriter_Constants. $LOW_SHAPE_POINT_TYPE_*

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
- Filled in ReadMe.
- Moved Search Descriptor functions from LibreOfficeWriter.au3 to LibreOfficeWriter_Helper.au3.
	- _LOWriter_SearchDescriptorCreate
	- _LOWriter_SearchDescriptorModify
	- _LOWriter_SearchDescriptorSimilarityModify
- Default value of $__LOWCONST_SLEEP_DIV from 15 to 0.
- Sorted Constants in LibreOfficeWriter_Constants alphabetically.
- Renamed $LOW_FIELDADV_TYPE_* Constants to $LOW_FIELD_ADV_TYPE_* to match formatting of other Field Type Constants.

### Fixed

- Missing error return value descriptions in:
	- _LOWriter_FootnoteGetAnchor
	- _LOWriter_EndnoteGetAnchor
	- _LOWriter_FrameGetAnchor
	- _LOWriter_ImageGetAnchor
- Missing Includes in LibreOfficeWriter file.
- Ternary Operators missing Parenthesis in responses.
- Export extension was incorrect, "jjpe" --> "jpe"
- Spell Checked the comments and Headers.

## [0.9.1] - 2023-10-28

**LibreOfficeWriter**

### Changed

- Minor documentation adjustments.

## [0.9.0] - 2023-10-28

**LibreOfficeWriter**

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
	- LibreOfficeWriter_DirectFormating, 
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
	- $LOW_PAPER_PORTRAIT --> $LOW_PAPER_ORIENT_PORTRAIT
	- $LOW_PAPER_LANDSCAPE --> $LOW_PAPER_ORIENT_LANDSCAPE
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

Thanks @danp2 and @Sven-Seyfert. All above mentioned MD documents was based on adequate documents from https://github.com/Danp2/au3WebDriver.

## [0.0.0.3] - 2023-08-10

**LibreOfficeWriter**

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
- Warning to _ShapesGetNames, about Images inserted in a document also being called "TextFrames".
- processing error check.
	- _DocInsertString, 
	- _DocInsertControlChar.
- a missing Datatype for possible Cursor data position types, $LOW_CURDATA_HEADER_FOOTER, previously attempting to insert a content while the insertion point was located in a Header/Footer would have failed. — Thanks to user @Heiko for helping me locate this error. The affected functions are: 
	- _CursorGetDataType.
	- _DocCreateTextCursor, 
	- _EndnoteInsert, 
	- _FootnoteInsert,	
	- _TableInsert, 

### Changed

- Renamed Name functions and examples for consistency.
	- _DocListTableNames --> _TablesGetNames 
	- _FramesListNames --> _FramesGetNames 
	- _ShapesListNames --> _ShapesGetNames 	
	- _ParGetObjects --> _ParObjCreateList. 
	- _ParSectionsGet --> _ParObjSectionsGet. 
	- _TableGetByCursor --> _TableGetObjByCursor. 
	- _TableGetByName --> _TableGetObjByName. 
	- _TableGetCellByCursor --> _TableGetCellObjByCursor. 
	- _TableGetCellByName --> _TableGetCellObjByName. 
	- _TableGetCellByPosition --> _TableGetCellObjByPosition. 
- Removed "IsCollpased" check and error from _DocGetString.
- _FramesListNames to have an option to search for Frames listed under shapes.
- _ShapesGetNames, Corrected an error that could occur if images are present.

### Fixed

-  An error where a COM error would be produced when attempting to insert a string or control character in certain data types. — Thanks to user @Heiko for helping me locate this error.
	- _DocInsertControlChar,
	- _DocInsertString.
	
## [0.0.0.2] - 2023-07-16

**LibreOfficeWriter**

### Added

- UDF version number in the UDF Header.

### Changed

- _DocReplaceAllInRange to have two methods of performing a Regular Expression find and replace.
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

## [0.0.0.1] - 2023-07-02

**LibreOfficeWriter**

### Added

- Initial UDF Release.

---

[Unreleased]: https://github.com/mlipok/Au3LibreOffice/compare/v1.0.0...v0.0.0.3
[0.0.0.3]:    https://github.com/mlipok/Au3LibreOffice/releases/tag/v0.0.0.3

### Legend - Types of changes

- `Added` for new features.
- `Changed` for changes in existing functionality.
- `Deprecated` for soon-to-be removed features.
- `Fixed` for any bug fixes.
- `Removed` for now removed features.
- `Security` in case of vulnerabilities.
- `Project` for documentation or overall project improvements.

##

[To the top](#)
