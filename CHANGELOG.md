#####

# Changelog

All notable changes to "Au3LibreOffice" SDK/API will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
This project also adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend---types-of-changes) for further information about the types of changes.

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
