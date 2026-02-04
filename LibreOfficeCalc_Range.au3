#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, or applying settings to L.O. Calc Cell Ranges.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_RangeAutoOutline
; _LOCalc_RangeClearContents
; _LOCalc_RangeColumnDelete
; _LOCalc_RangeColumnGetName
; _LOCalc_RangeColumnGetObjByName
; _LOCalc_RangeColumnGetObjByPosition
; _LOCalc_RangeColumnInsert
; _LOCalc_RangeColumnPageBreak
; _LOCalc_RangeColumnsGetCount
; _LOCalc_RangeColumnVisible
; _LOCalc_RangeColumnWidth
; _LOCalc_RangeCompute
; _LOCalc_RangeCopyMove
; _LOCalc_RangeCreateCursor
; _LOCalc_RangeData
; _LOCalc_RangeDatabaseAdd
; _LOCalc_RangeDatabaseDelete
; _LOCalc_RangeDatabaseExists
; _LOCalc_RangeDatabaseGetNames
; _LOCalc_RangeDatabaseGetObjByName
; _LOCalc_RangeDatabaseModify
; _LOCalc_RangeDelete
; _LOCalc_RangeDetail
; _LOCalc_RangeFill
; _LOCalc_RangeFillRandom
; _LOCalc_RangeFillSeries
; _LOCalc_RangeFilter
; _LOCalc_RangeFilterAdvanced
; _LOCalc_RangeFilterClear
; _LOCalc_RangeFindAll
; _LOCalc_RangeFindNext
; _LOCalc_RangeFormula
; _LOCalc_RangeGetAddressAsName
; _LOCalc_RangeGetAddressAsPosition
; _LOCalc_RangeGetCellByName
; _LOCalc_RangeGetCellByPosition
; _LOCalc_RangeGetSheet
; _LOCalc_RangeGroup
; _LOCalc_RangeInsert
; _LOCalc_RangeIsMerged
; _LOCalc_RangeMerge
; _LOCalc_RangeNamedAdd
; _LOCalc_RangeNamedChangeScope
; _LOCalc_RangeNamedDelete
; _LOCalc_RangeNamedExists
; _LOCalc_RangeNamedGetNames
; _LOCalc_RangeNamedGetObjByName
; _LOCalc_RangeNamedModify
; _LOCalc_RangeNumbers
; _LOCalc_RangeOutlineClearAll
; _LOCalc_RangeOutlineShow
; _LOCalc_RangePivotDelete
; _LOCalc_RangePivotDest
; _LOCalc_RangePivotExists
; _LOCalc_RangePivotFieldGetObjByName
; _LOCalc_RangePivotFieldItemsGetNames
; _LOCalc_RangePivotFieldsColumnsGetNames
; _LOCalc_RangePivotFieldsDataGetNames
; _LOCalc_RangePivotFieldSettings
; _LOCalc_RangePivotFieldsFiltersGetNames
; _LOCalc_RangePivotFieldsGetNames
; _LOCalc_RangePivotFieldsRowsGetNames
; _LOCalc_RangePivotFieldsUnusedGetNames
; _LOCalc_RangePivotFilter
; _LOCalc_RangePivotFilterClear
; _LOCalc_RangePivotGetObjByIndex
; _LOCalc_RangePivotGetObjByName
; _LOCalc_RangePivotInsert
; _LOCalc_RangePivotName
; _LOCalc_RangePivotRefresh
; _LOCalc_RangePivotSettings
; _LOCalc_RangePivotsGetCount
; _LOCalc_RangePivotsGetNames
; _LOCalc_RangePivotSource
; _LOCalc_RangeQueryColumnDiff
; _LOCalc_RangeQueryContents
; _LOCalc_RangeQueryDependents
; _LOCalc_RangeQueryEmpty
; _LOCalc_RangeQueryFormula
; _LOCalc_RangeQueryIntersection
; _LOCalc_RangeQueryPrecedents
; _LOCalc_RangeQueryRowDiff
; _LOCalc_RangeQueryVisible
; _LOCalc_RangeReplace
; _LOCalc_RangeReplaceAll
; _LOCalc_RangeRowDelete
; _LOCalc_RangeRowGetObjByPosition
; _LOCalc_RangeRowHeight
; _LOCalc_RangeRowInsert
; _LOCalc_RangeRowPageBreak
; _LOCalc_RangeRowsGetCount
; _LOCalc_RangeRowVisible
; _LOCalc_RangeSort
; _LOCalc_RangeSortAlt
; _LOCalc_RangeValidation
; _LOCalc_RangeValidationSettings
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeAutoOutline
; Description ...: Set up AutoOutline for a Range of Cells.
; Syntax ........: _LOCalc_RangeAutoOutline(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Range Address Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. AutoOutline was successfully processed for range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeAutoOutline(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tRangeAddress

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tRangeAddress = $oRange.RangeAddress()
	If Not IsObj($tRangeAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oRange.Spreadsheet.autoOutline($tRangeAddress)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeAutoOutline

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeClearContents
; Description ...: Clear specific cell contents in a range.
; Syntax ........: _LOCalc_RangeClearContents(ByRef $oRange, $iFlags)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell to clear the contents of. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFlags              - an integer value (1-1023). The Cell Content type to delete. Can be BitOR'd together. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFlags not an Integer, less than 1 or greater than 1023. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Contents specified was successfully cleared from the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeClearContents(ByRef $oRange, $iFlags)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iFlags, $LOC_CELL_FLAG_VALUE, $LOC_CELL_FLAG_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRange.clearContents($iFlags)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeClearContents

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnDelete
; Description ...: Delete Columns from a Range.
; Syntax ........: _LOCalc_RangeColumnDelete(ByRef $oRange, $iColumn[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iColumn             - an integer value. The column to begin deleting at. The Column called will be deleted. See remarks.
;                  $iCount              - [optional] an integer value. Default is 1. The number of columns to delete after the called column.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumns not an Integer, less than 0 or greater than number of Columns contained in the Range.
;                  @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested columns.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Columns in L.O. Calc are 0 based, to Delete Column "A" in the LibreOffice UI, you would call $iColumn with 0.
;                  Deleting Columns does not decrease the Column count, it simply erases the Column's contents in a specific area and shifts all after content left.
; Related .......: _LOCalc_RangeColumnInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnDelete(ByRef $oRange, $iColumn, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iColumn, 0, $oColumns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iCount, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumns.removeByIndex($iColumn, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeColumnDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnGetName
; Description ...: Retrieve the Column's name.
; Syntax ........: _LOCalc_RangeColumnGetName(ByRef $oColumn)
; Parameters ....: $oColumn             - [in/out] an object. A Column object returned by a previous _LOCalc_RangeColumnGetObjByPosition, or _LOCalc_RangeColumnGetObjByName function.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the Column's name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Success, returning Column's name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnGetName(ByRef $oColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sName = $oColumn.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOCalc_RangeColumnGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnGetObjByName
; Description ...: Retrieve a Column's Object by name.
; Syntax ........: _LOCalc_RangeColumnGetObjByName(ByRef $oRange, $sName)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sName               - a string value. The Column name to retrieve the Object for, such as "A".
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Range does not contain a column with name called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Column Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Success, returning Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeColumnGetObjByPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnGetObjByName(ByRef $oRange, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $oColumn

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oColumns.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumn = $oColumns.getByName($sName)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOCalc_RangeColumnGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnGetObjByPosition
; Description ...: Retrieve the Column's Object by its position.
; Syntax ........: _LOCalc_RangeColumnGetObjByPosition(ByRef $oRange, $iColumn)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iColumn             - an integer value. The Column number to retrieve the Object for. See remarks.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, less than 0 or greater than number of columns contained in the Range.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Column Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Success, returning Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Columns in L.O. Calc are 0 based, to retrieve Column "A" in the LibreOffice UI, you would call $iColumn with 0.
; Related .......: _LOCalc_RangeColumnGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnGetObjByPosition(ByRef $oRange, $iColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $oColumn

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iColumn, 0, $oColumns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumn = $oColumns.getByIndex($iColumn)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOCalc_RangeColumnGetObjByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnInsert
; Description ...: Insert blank columns into a Range at a specific column.
; Syntax ........: _LOCalc_RangeColumnInsert(ByRef $oRange, $iColumn[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iColumn             - an integer value. The Column to insert the new column(s) at. See remarks. New columns will be inserted starting at this column and all content will be shifted right.
;                  $iCount              - [optional] an integer value. Default is 1. The number of blank columns to insert after the Column called.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, less than 0 or greater than number of Columns contained in the Range.
;                  @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully inserted blank columns.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Columns in L.O. Calc are 0 based, to add columns in Column "A" in the LibreOffice UI, you would call $iColumn with 0.
;                  Inserting Columns does not increase the Column count, it simply adds blanks in a specific area and shifts all after content further right.
; Related .......: _LOCalc_RangeColumnDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnInsert(ByRef $oRange, $iColumn, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iColumn, 0, $oColumns.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iCount, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumns.insertByIndex($iColumn, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeColumnInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnPageBreak
; Description ...: Set or retrieve Page Break settings for a Column.
; Syntax ........: _LOCalc_RangeColumnPageBreak(ByRef $oColumn[, $bManualPageBreak = Null[, $bStartOfPageBreak = Null]])
; Parameters ....: $oColumn             - [in/out] an object. A Column object returned by a previous _LOCalc_RangeColumnGetObjByPosition, or _LOCalc_RangeColumnGetObjByName function.
;                  $bManualPageBreak    - [optional] a boolean value. Default is Null. If True, this column is the beginning of a manual Page Break.
;                  $bStartOfPageBreak   - [optional] a boolean value. Default is Null. If True, this column is the beginning of a start of Page Break. See Remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bManualPageBreak not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bStartOfPageBreak not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bManualPageBreak
;                  |                               2 = Error setting $bStartOfPageBreak
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Setting $bStartOfPageBreak to True will insert a Manual Page Break, the same as setting $bManualPageBreak to True would.
;                  $bStartOfPageBreak setting is available more for indicating where Calc is inserting Page Breaks rather than for applying a setting. You can retrieve the settings for each Column, and check if this value is True or not. If the Page break is an automatically inserted one, the value for $bManualPageBreak would be False.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnPageBreak(ByRef $oColumn, $bManualPageBreak = Null, $bStartOfPageBreak = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abBreak[2]

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bManualPageBreak, $bStartOfPageBreak) Then
		__LO_ArrayFill($abBreak, $oColumn.IsManualPageBreak(), $oColumn.IsStartOfNewPage())

		Return SetError($__LO_STATUS_SUCCESS, 1, $abBreak)
	EndIf

	If ($bManualPageBreak <> Null) Then
		If Not IsBool($bManualPageBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oColumn.IsManualPageBreak = $bManualPageBreak
		$iError = ($oColumn.IsManualPageBreak() = $bManualPageBreak) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bStartOfPageBreak <> Null) Then
		If Not IsBool($bStartOfPageBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oColumn.IsStartOfNewPage = $bStartOfPageBreak
		$iError = ($oColumn.IsStartOfNewPage() = $bStartOfPageBreak) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeColumnPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnsGetCount
; Description ...: Retrieve the total count of Columns contained in a Range.
; Syntax ........: _LOCalc_RangeColumnsGetCount(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning number of Columns contained in the Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: There is a fixed number of Columns per sheet, but different L.O. versions contain different amounts of Columns. But this also helps to determine how many columns are contained in a Cell Range.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnsGetCount(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumns.Count())
EndFunc   ;==>_LOCalc_RangeColumnsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnVisible
; Description ...: Set or Retrieve the Column's visibility setting.
; Syntax ........: _LOCalc_RangeColumnVisible(ByRef $oColumn[, $bVisible = Null])
; Parameters ....: $oColumn             - an object. A Column object returned by a previous _LOCalc_RangeColumnGetObjByPosition, or _LOCalc_RangeColumnGetObjByName function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the Column is Visible.
; Return values .: Success: 1 or Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were called with Null, returning Column's current visibility setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnVisible(ByRef $oColumn, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bVisible) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oColumn.IsVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumn.IsVisible = $bVisible
	$iError = ($oColumn.IsVisible() = $bVisible) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeColumnVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnWidth
; Description ...: Set or Retrieve the Column's Width settings.
; Syntax ........: _LOCalc_RangeColumnWidth(ByRef $oColumn[, $bOptimal = Null[, $iWidth = Null]])
; Parameters ....: $oColumn             - an object. A Column object returned by a previous _LOCalc_RangeColumnGetObjByPosition, or _LOCalc_RangeColumnGetObjByName function.
;                  $bOptimal            - [optional] a boolean value. Default is Null. If True, the Optimal width is automatically chosen. See Remarks.
;                  $iWidth              - [optional] an integer value (0-34464). Default is Null. The Width of the Column, set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bOptimal not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer, less than 0 or greater than 34464.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bOptimal
;                  |                               2 = Error setting $iWidth
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bOptimal only accepts True. False will return an error. Calling True again returns the cell to optimal width, setting a custom width essentially disables it.
;                  I am presently unable to find a setting for Optimal Width "Add" Value.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnWidth(ByRef $oColumn, $bOptimal = Null, $iWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avWidth[2]
	Local $iError = 0

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bOptimal, $iWidth) Then
		__LO_ArrayFill($avWidth, $oColumn.OptimalWidth(), $oColumn.Width())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avWidth)
	EndIf

	If ($bOptimal <> Null) Then
		If Not IsBool($bOptimal) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oColumn.OptimalWidth = $bOptimal
		$iError = ($oColumn.OptimalWidth() = $bOptimal) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 0, 34464) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oColumn.Width = $iWidth
		$iError = (__LO_IntIsBetween($oColumn.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeColumnWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeCompute
; Description ...: Perform a Computation function on a Range. See Remarks.
; Syntax ........: _LOCalc_RangeCompute(ByRef $oRange, $iFunction)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFunction           - an integer value (0-12). The Computation Function to perform. See Constants $LOC_COMPUTE_FUNC_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Number
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFunction not an Integer, less than 0 or greater than 12. See Constants $LOC_COMPUTE_FUNC_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to perform computation.
;                  --Success--
;                  @Error 0 @Extended 0 Return Number = Success. Successfully performed the requested computation, returning the result as a Numerical value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This makes no changes in the document itself, it only returns the result of the computation.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeCompute(ByRef $oRange, $iFunction)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $nResult

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iFunction, $LOC_COMPUTE_FUNC_NONE, $LOC_COMPUTE_FUNC_VARP) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$nResult = $oRange.computeFunction($iFunction)
	If Not IsNumber($nResult) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $nResult)
EndFunc   ;==>_LOCalc_RangeCompute

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeCopyMove
; Description ...: Copy or Move a Cell or Cell Range to another range.
; Syntax ........: _LOCalc_RangeCopyMove(ByRef $oSheet, ByRef $oRangeSrc, ByRef $oRangeDest[, $bMove = False])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRangeSrc           - [in/out] an object. The Cell or Cell Range to copy or move from. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oRangeDest          - [in/out] an object. The Cell or Cell Range to copy or move to. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bMove               - [optional] a boolean value. Default is False. If True, the cell range is moved to the destination. If False, the Cell Range is only copied.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRangeSrc not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oRangeDest not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bMove not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.table.CellAddress" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Source Cell Range Address.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Cell Range Address.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Cell or Cell range was successfully copied to destination.
;                  @Error 0 @Extended 1 Return 1 = Success. Cell or Cell range was successfully moved to destination.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Destination Range can be on different Sheet from Source.
;                  $oSheet is the source sheet where $oRangeSrc is located.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeCopyMove(ByRef $oSheet, ByRef $oRangeSrc, ByRef $oRangeDest, $bMove = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tCellSrcAddr, $tInputCellDestAddr, $tCellDestAddr

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRangeSrc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oRangeDest) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bMove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$tCellSrcAddr = $oRangeSrc.RangeAddress()
	If Not IsObj($tCellSrcAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tInputCellDestAddr = $oRangeDest.RangeAddress()
	If Not IsObj($tInputCellDestAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tCellDestAddr = __LO_CreateStruct("com.sun.star.table.CellAddress")
	If Not IsObj($tCellDestAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tCellDestAddr.Sheet = $tInputCellDestAddr.Sheet()
	$tCellDestAddr.Column = $tInputCellDestAddr.StartColumn()
	$tCellDestAddr.Row = $tInputCellDestAddr.StartRow()

	If $bMove Then
		$oSheet.MoveRange($tCellDestAddr, $tCellSrcAddr)

		Return SetError($__LO_STATUS_SUCCESS, 1, 1)

	Else
		$oSheet.CopyRange($tCellDestAddr, $tCellSrcAddr)

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndIf
EndFunc   ;==>_LOCalc_RangeCopyMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeCreateCursor
; Description ...: Create a Sheet Cursor for a particular range.
; Syntax ........: _LOCalc_RangeCreateCursor(ByRef $oSheet, ByRef $oRange)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Sheet Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully created a Sheet Cursor for the specified Range, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Sheet Cursor can be used in functions accepting a range. When created, the Cursor will have the called range selected.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeCreateCursor(ByRef $oSheet, ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheetCursor

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheetCursor = $oSheet.createCursorByRange($oRange)
	If Not IsObj($oSheetCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheetCursor)
EndFunc   ;==>_LOCalc_RangeCreateCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeData
; Description ...: Set or Retrieve Data in a Range.
; Syntax ........: _LOCalc_RangeData(ByRef $oRange[, $aavData = Null[, $bStrictSize = False]])
; Parameters ....: $oRange              - [in/out] an object. The Cell or Cell Range to set or retrieve data . A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $aavData             - [optional] an array of Arrays containing variants. Default is Null. An Array of Arrays containing data, strings or numbers, to fill the range with. See remarks.
;                  $bStrictSize         - [optional] a boolean value. Default is False. If True, The Range size must explicitly match the array sizing. If False, The Range will be resized right or down to fit the Array sizing.
; Return values .: Success: 1 or Array
;                  Failure: 0 or ? and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $aavData not an Array.
;                  @Error 1 @Extended 3 Return 0 = $bStrictSize not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bStrictSize called with True, and $aavData array contains less or more elements than number of rows contained in the cell range.
;                  @Error 1 @Extended 5 Return ? = Element of $aavData does not contain an array. Returning array element number of $aavData containing error.
;                  @Error 1 @Extended 6 Return ? = $bStrictSize called with True, and Array contained in $aavData has less or more elements than number of columns in the cell range. Returning array element number of $aavData containing faulty array.
;                  @Error 1 @Extended 7 Return ? = $bStrictSize called with False, and Array contained in $aavData has less or more elements than first Array contained in $aavData. Returning array element number of $aavData containing faulty array.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve array of Formula Data contained in the Cell Range.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Start of Row from Cell Range.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve End of Row from Cell Range.
;                  @Error 3 @Extended 4 Return 0 = Expanding Range would exceed number of Rows contained in Sheet.
;                  @Error 3 @Extended 5 Return 0 = Failed to re-size Cell Range Rows.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Start of Column from Cell Range.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve End of Column from Cell Range.
;                  @Error 3 @Extended 8 Return 0 = Expanding Range would exceed number of Columns contained in Sheet.
;                  @Error 3 @Extended 9 Return 0 = Failed to re-size Cell Range Columns.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Data was successfully set for the cell range.
;                  @Error 0 @Extended 1 Return Array of Arrays = Success. $aavData called with Null, returning an array containing arrays, which contain any data content contained in the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return Strings and Numbers contained in the cell range when $aavData is called with Null keyword. Array will be an array of arrays. The internal arrays will contain numerical or string data, depending on cell content.
;                  $aavData must be an array containing arrays. If $bStrictSize is called with True, the main Array's element count must match the row count contained in the Cell Range, and each internal Array's element count must match the column count of the Cell Range it is to fill. All internal arrays must be the same size.
;                  Any data previously contained in the Cell Range will be overwritten.
;                  All array elements must contain appropriate data, strings or numbers.
;                  Formulas will be inserted as strings only, and will not be valid.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeData(ByRef $oRange, $aavData = Null, $bStrictSize = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iStart, $iEnd

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($aavData) Then
		$aavData = $oRange.getDataArray()
		If Not IsArray($aavData) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $aavData)
	EndIf

	If Not IsArray($aavData) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bStrictSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	; Determine if the Array is sized appropriately
	$iStart = $oRange.RangeAddress.StartRow()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iEnd = $oRange.RangeAddress.EndRow()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $bStrictSize Then ; If Array is wrongly sized, return an error.
		If (UBound($aavData) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Else ; Expand the Range to fit the Array
		If (UBound($aavData) <> ($iEnd - $iStart + 1)) Then
			If (($oRange.RangeAddress.StartRow() + UBound($aavData)) > $oRange.Spreadsheet.getRows.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; Check if resizing range is possible.

			$oRange = $oRange.Spreadsheet.getCellRangeByPosition($oRange.RangeAddress.StartColumn(), $oRange.RangeAddress.StartRow(), $oRange.RangeAddress.EndColumn(), ($oRange.RangeAddress.StartRow() + UBound($aavData) - 1))
			If Not IsObj($oRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
		EndIf
	EndIf

	$iStart = $oRange.RangeAddress.StartColumn()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	$iEnd = $oRange.RangeAddress.EndColumn()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

	If $bStrictSize Then ; Check if the internal arrays are sized correctly, return error if not.

		For $i = 0 To UBound($aavData) - 1
			If Not IsArray($aavData[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
			If (UBound($aavData[$i]) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)
		Next

	Else ; Check if the internal arrays are sized correctly, resize range if not.
		For $i = 0 To UBound($aavData) - 1
			If Not IsArray($aavData[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
			If (UBound($aavData[$i]) <> UBound($aavData[0])) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $i) ; If all arrays aren't same size as first array, then error.
		Next

		If (UBound($aavData[0]) <> ($iEnd - $iStart + 1)) Then ; Resize the Range.
			If (($oRange.RangeAddress.StartColumn() + UBound($aavData[0])) > $oRange.Spreadsheet.getColumns.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

			$oRange = $oRange.Spreadsheet.getCellRangeByPosition($oRange.RangeAddress.StartColumn(), $oRange.RangeAddress.StartRow(), ($oRange.RangeAddress.StartColumn() + UBound($aavData[0]) - 1), $oRange.RangeAddress.EndRow())
			If Not IsObj($oRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 9, 0)
		EndIf
	EndIf

	$oRange.setDataArray($aavData)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeDatabaseAdd
; Description ...: Add a Database Range to a document.
; Syntax ........: _LOCalc_RangeDatabaseAdd(ByRef $oDoc, $oRange, $sName[, $bColumnHeaders = True[, $bTotalsRow = False[, $bAddDeleteCells = True[, $bKeepFormatting = True[, $bDontSaveImport = False[, $bAutoFilter = False]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oRange              - an object. The Range to designate as a Database range. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sName               - a string value. The unique name of the Database Range to create.
;                  $bColumnHeaders      - [optional] a boolean value. Default is True. If True, the top row is considered a Header/label.
;                  $bTotalsRow          - [optional] a boolean value. Default is False. If True, the bottom row will be considered a totals row.
;                  $bAddDeleteCells     - [optional] a boolean value. Default is True. If True, columns or rows are inserted or deleted when the size of the range is changed by an update operation.
;                  $bKeepFormatting     - [optional] a boolean value. Default is True. If True, cell formats are extended when the size of the range is changed by an update operation.
;                  $bDontSaveImport     - [optional] a boolean value. Default is False. If True, cell contents within the database range are left out when the document is saved.
;                  $bAutoFilter         - [optional] a boolean value. Default is False. If True, the Auto Filter option is enabled.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $bColumnHeaders not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bTotalsRow not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bAddDeleteCells not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bKeepFormatting not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bDontSaveImport not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bAutoFilter not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = Document called in $oDoc already contains a Database Range named the same as called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Database Ranges Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve new Database Range's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully added a new Database Range, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeDatabaseExists, _LOCalc_RangeDatabaseDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeDatabaseAdd(ByRef $oDoc, $oRange, $sName, $bColumnHeaders = True, $bTotalsRow = False, $bAddDeleteCells = True, $bKeepFormatting = True, $bDontSaveImport = False, $bAutoFilter = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDatabaseRanges, $oDatabaseRange

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bColumnHeaders) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bTotalsRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bAddDeleteCells) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bKeepFormatting) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsBool($bDontSaveImport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not IsBool($bAutoFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$oDatabaseRanges = $oDoc.DatabaseRanges()
	If Not IsObj($oDatabaseRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oDatabaseRanges.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

	$oDatabaseRanges.addNewByName($sName, $oRange.RangeAddress())

	$oDatabaseRange = $oDatabaseRanges.getByName($sName)
	If Not IsObj($oDatabaseRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	With $oDatabaseRange
		.ContainsHeader = $bColumnHeaders
		.TotalsRow = $bTotalsRow
		.MoveCells = $bAddDeleteCells
		.KeepFormats = $bKeepFormatting
		.StripData = $bDontSaveImport
		.AutoFilter = $bAutoFilter
	EndWith

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDatabaseRange)
EndFunc   ;==>_LOCalc_RangeDatabaseAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeDatabaseDelete
; Description ...: Delete a Database Range from the document.
; Syntax ........: _LOCalc_RangeDatabaseDelete(ByRef $oDoc, $oDatabaseRange)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oDatabaseRange      - an object. A Database Range Object as returned from _LOCalc_RangeDatabaseAdd or _LOCalc_RangeDatabaseGetObjByName.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oDatabaseRange not an Object and not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Database Ranges Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Database Range name.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete requested Database Range.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully deleted the requested Database Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeDatabaseAdd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeDatabaseDelete(ByRef $oDoc, $oDatabaseRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDatabaseRanges
	Local $sDatabaseRange

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDatabaseRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDatabaseRanges = $oDoc.DatabaseRanges()
	If Not IsObj($oDatabaseRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sDatabaseRange = $oDatabaseRange.Name()
	If Not IsString($sDatabaseRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDatabaseRanges.removeByName($sDatabaseRange)

	If $oDatabaseRanges.hasByName($sDatabaseRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeDatabaseDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeDatabaseExists
; Description ...: Check if a Database Range exists in a document.
; Syntax ........: _LOCalc_RangeDatabaseExists(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object.
;                  $sName               - a string value.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Database Ranges Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to query whether document contains the called name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the document contains a Database Range by the called name. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeDatabaseExists(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDatabaseRanges
	Local $bExists

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDatabaseRanges = $oDoc.DatabaseRanges()
	If Not IsObj($oDatabaseRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$bExists = $oDatabaseRanges.hasByName($sName)
	If Not IsBool($bExists) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bExists)
EndFunc   ;==>_LOCalc_RangeDatabaseExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeDatabaseGetNames
; Description ...: Retrieve an array of Database Range names for the document.
; Syntax ........: _LOCalc_RangeDatabaseGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Database Ranges Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Database Ranges contained in the document. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeDatabaseGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeDatabaseGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDatabaseRanges
	Local $asNames[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDatabaseRanges = $oDoc.DatabaseRanges()
	If Not IsObj($oDatabaseRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asNames[$oDatabaseRanges.Count()]

	For $i = 0 To $oDatabaseRanges.Count() - 1
		$asNames[$i] = $oDatabaseRanges.getByIndex($i).Name()

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOCalc_RangeDatabaseGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeDatabaseGetObjByName
; Description ...: Retrieve a Database Range Object by Name.
; Syntax ........: _LOCalc_RangeDatabaseGetObjByName(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - a string value. The name of the Database Range to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Document called in $oDoc does not contain a Database Range by the name called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Database Ranges Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve requested Database Range Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Database Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeDatabaseExists, _LOCalc_RangeDatabaseGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeDatabaseGetObjByName(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDatabaseRanges, $oDatabaseRange

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDatabaseRanges = $oDoc.DatabaseRanges()
	If Not IsObj($oDatabaseRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oDatabaseRanges.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oDatabaseRange = $oDatabaseRanges.getByName($sName)
	If Not IsObj($oDatabaseRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDatabaseRange)
EndFunc   ;==>_LOCalc_RangeDatabaseGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeDatabaseModify
; Description ...: Set or Retrieve the settings for a Database Range.
; Syntax ........: _LOCalc_RangeDatabaseModify(ByRef $oDoc, ByRef $oDatabaseRange[, $oRange = Null[, $sName = Null[, $bColumnHeaders = Null[, $bTotalsRow = Null[, $bAddDeleteCells = Null[, $bKeepFormatting = Null[, $bDontSaveImport = Null[, $bAutoFilter = Null]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oDatabaseRange      - [in/out] an object. A Database Range Object as returned from _LOCalc_RangeDatabaseAdd or _LOCalc_RangeDatabaseGetObjByName.
;                  $oRange              - [optional] an object. Default is Null. The Range to designate as a Database range. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sName               - [optional] a string value. Default is Null. The new unique name to rename the Database Range to.
;                  $bColumnHeaders      - [optional] a boolean value. Default is Null. If True, the top row is considered a Header/label.
;                  $bTotalsRow          - [optional] a boolean value. Default is Null. If True, the bottom row will be considered a totals row.
;                  $bAddDeleteCells     - [optional] a boolean value. Default is Null. If True, columns or rows are inserted or deleted when the size of the range is changed by an update operation.
;                  $bKeepFormatting     - [optional] a boolean value. Default is Null. If True, cell formats are extended when the size of the range is changed by an update operation.
;                  $bDontSaveImport     - [optional] a boolean value. Default is Null. If True, cell contents within the database range are left out when the document is saved.
;                  $bAutoFilter         - [optional] a boolean value. Default is Null. If True, the Auto Filter option is enabled.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oDatabaseRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 4 Return 0 = $sName not a String.
;                  @Error 1 @Extended 5 Return 0 = Document already contains a Database Range with the name as called in $sName.
;                  @Error 1 @Extended 6 Return 0 = $bColumnHeaders not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bTotalsRow not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bAddDeleteCells not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bKeepFormatting not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bDontSaveImport not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bAutoFilter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Cell Object referenced by this Named Range.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oRange
;                  |                               2 = Error setting $sName
;                  |                               4 = Error setting $bColumnHeaders
;                  |                               8 = Error setting $bTotalsRow
;                  |                               16 = Error setting $bAddDeleteCells
;                  |                               32 = Error setting $bKeepFormatting
;                  |                               64 = Error setting $bDontSaveImport
;                  |                               128 = Error setting $bAutoFilter
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  When retrieving the settings, $oRange will be a Range Object.
; Related .......: _LOCalc_RangeDatabaseGetObjByName, _LOCalc_RangeDatabaseAdd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeDatabaseModify(ByRef $oDoc, ByRef $oDatabaseRange, $oRange = Null, $sName = Null, $bColumnHeaders = Null, $bTotalsRow = Null, $bAddDeleteCells = Null, $bKeepFormatting = Null, $bDontSaveImport = Null, $bAutoFilter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avDatabaseRange[8]
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDatabaseRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($oRange, $sName, $bColumnHeaders, $bTotalsRow, $bAddDeleteCells, $bKeepFormatting, $bDontSaveImport, $bAutoFilter) Then
		$oRange = $oDoc.Sheets.getByIndex($oDatabaseRange.DataArea.Sheet()).getCellRangeByPosition( _
				$oDatabaseRange.DataArea.StartColumn(), $oDatabaseRange.DataArea.StartRow(), _
				$oDatabaseRange.DataArea.EndColumn(), $oDatabaseRange.DataArea.EndRow())
		If Not IsObj($oRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LO_ArrayFill($avDatabaseRange, $oRange, $oDatabaseRange.Name(), $oDatabaseRange.ContainsHeader(), $oDatabaseRange.TotalsRow(), $oDatabaseRange.MoveCells(), _
				$oDatabaseRange.KeepFormats(), $oDatabaseRange.StripData(), $oDatabaseRange.AutoFilter())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avDatabaseRange)
	EndIf

	If ($oRange <> Null) Then
		If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDatabaseRange.DataArea = $oRange.RangeAddress()
		$iError = (__LOCalc_RangeAddressIsSame($oDatabaseRange.DataArea(), $oRange.RangeAddress())) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If $oDoc.DatabaseRanges.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oDatabaseRange.Name = $sName
		$iError = ($oDatabaseRange.Name() = $sName) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bColumnHeaders <> Null) Then
		If Not IsBool($bColumnHeaders) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oDatabaseRange.ContainsHeader = $bColumnHeaders
		$iError = ($oDatabaseRange.ContainsHeader() = $bColumnHeaders) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bTotalsRow <> Null) Then
		If Not IsBool($bTotalsRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oDatabaseRange.TotalsRow = $bTotalsRow
		$iError = ($oDatabaseRange.TotalsRow() = $bTotalsRow) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bAddDeleteCells <> Null) Then
		If Not IsBool($bAddDeleteCells) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oDatabaseRange.MoveCells = $bAddDeleteCells
		$iError = ($oDatabaseRange.MoveCells() = $bAddDeleteCells) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bKeepFormatting <> Null) Then
		If Not IsBool($bKeepFormatting) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oDatabaseRange.KeepFormats = $bKeepFormatting
		$iError = ($oDatabaseRange.KeepFormats() = $bKeepFormatting) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bDontSaveImport <> Null) Then
		If Not IsBool($bDontSaveImport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oDatabaseRange.StripData = $bDontSaveImport
		$iError = ($oDatabaseRange.StripData() = $bDontSaveImport) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bAutoFilter <> Null) Then
		If Not IsBool($bAutoFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oDatabaseRange.AutoFilter = $bAutoFilter
		$iError = ($oDatabaseRange.AutoFilter() = $bAutoFilter) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeDatabaseModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeDelete
; Description ...: Delete a Range of cell contents and reposition surrounding cells.
; Syntax ........: _LOCalc_RangeDelete(ByRef $oSheet, $oRange, $iMode)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRange              - an object. A Cell or Cell range to delete. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iMode               - an integer value (0-4). The Cell Deletion Mode. See Constants $LOC_CELL_DELETE_MODE_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iMode not an Integer, less than 0 or greater than 4. See Constants $LOC_CELL_DELETE_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Range Address Struct from $oRange.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Cell range was successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will silently fail if the deletion will cause an array formula to be split -- OOME. 4.1., Page 509.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeDelete(ByRef $oSheet, $oRange, $iMode)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tCellAddr

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iMode, $LOC_CELL_DELETE_MODE_NONE, $LOC_CELL_DELETE_MODE_COLUMNS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oRange.RangeAddress()
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSheet.removeRange($tCellAddr, $iMode)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeDetail
; Description ...: Expand or Hide Grouped cells in a Range.
; Syntax ........: _LOCalc_RangeDetail(ByRef $oRange[, $bShow = True])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bShow               - [optional] a boolean value. Default is True. If True, grouped cells are expanded, If False, grouped cells are collapsed.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bShow not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Range Address Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Expand or Hiding of Grouped cells was successfully processed for Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeDetail(ByRef $oRange, $bShow = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tRangeAddress

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bShow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tRangeAddress = $oRange.RangeAddress()
	If Not IsObj($tRangeAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $bShow Then
		$oRange.Spreadsheet.showDetail($tRangeAddress)

	Else
		$oRange.Spreadsheet.hideDetail($tRangeAddress)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeDetail

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFill
; Description ...: Automatically fill cells with a value. See Remarks.
; Syntax ........: _LOCalc_RangeFill(ByRef $oRange, $iDirection[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iDirection          - an integer value (0-3). The Direction to perform the Fill operation. See Constants $LOC_FILL_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iCount              - [optional] an integer value. Default is 1. The number of Cells to take into account at the beginning of the range to constitute the fill algorithm.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iDirection not an Integer, less than 0 or greater than 3. See Constants $LOC_FILL_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Fill operation was successfully processed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Fill value is calculated based on the first value(s) in the Range, the first value location depends on the Fill direction. If Fill direction is set to Right, the initial value must be in the first cell(s) on the left, and vice versa.
; Related .......: _LOCalc_RangeFillSeries
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFill(ByRef $oRange, $iDirection, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iDirection, $LOC_FILL_DIR_DOWN, $LOC_FILL_DIR_LEFT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iCount, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRange.fillAuto($iDirection, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFill

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFillRandom
; Description ...: Fill a range with random numbers.
; Syntax ........: _LOCalc_RangeFillRandom(ByRef $oRange[, $nMin = 0.0000[, $nMax = 1.0000[, $iDecPlc = 15[, $nSeed = Null[, $bFillByRows = True]]]]])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $nMin                - [optional] a general number value (-2^31-2^31). Default is 0.0000. The minimum number value. Max is -2^31-2^31.
;                  $nMax                - [optional] a general number value (-2^31-2^31). Default is 1.0000. The maximum number value. Max is -2^31-2^31.
;                  $iDecPlc             - [optional] an integer value (0-255). Default is 15. The decimal place to round the value to. Call with 0 to fill with Integers only.
;                  $nSeed               - [optional] a general number value. Default is Null. A seed to use for generating the Random number. Null means no seed is used.
;                  $bFillByRows         - [optional] a boolean value. Default is True. If True, the range is filled top to bottom, left to right. If False, the range is filled left to right, top to bottom.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $nMin not a number, less than -2^31 or greater then 2^31.
;                  @Error 1 @Extended 3 Return 0 = $nMax not a number, less than -2^31 or greater then 2^31.
;                  @Error 1 @Extended 4 Return 0 = $iDecPlc not an Integer, less than 0 or greater than 255.
;                  @Error 1 @Extended 5 Return 0 = $nSeed not a number, less than -2^31 or greater then 2^31.
;                  @Error 1 @Extended 6 Return 0 = $bFillByRows not a boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Range successfully filled with random values.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is a homemade version of Calc's Fill Random, as there is no built in method for calling Libre Office's built-in one. The results of this function may not be similar to the results of Libre Office's random number generator.
;                  Any values in the range will be overwritten.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFillRandom(ByRef $oRange, $nMin = 0.0000, $nMax = 1.0000, $iDecPlc = 15, $nSeed = Null, $bFillByRows = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_NumIsBetween($nMin, -2 ^ 31, 2 ^ 31) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_NumIsBetween($nMax, -2 ^ 31, 2 ^ 31) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_NumIsBetween($iDecPlc, 0, 255) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($nSeed <> Null) And Not __LO_NumIsBetween($nSeed, -2 ^ 31, 2 ^ 31) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bFillByRows) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If ($nSeed <> Null) Then SRandom($nSeed)

	If $bFillByRows Then ; Fill all rows first per column.
		For $iC = $oRange.RangeAddress.StartColumn() To $oRange.RangeAddress.EndColumn()
			For $iR = $oRange.RangeAddress.StartRow() To $oRange.RangeAddress.EndRow()
				$oRange.Spreadsheet.getCellByPosition($iC, $iR).Value = Round(Random($nMin, $nMax), $iDecPlc)
				Sleep((IsInt($iR / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
			Next
		Next

	Else ; Fill all columns first per row.
		For $iR = $oRange.RangeAddress.StartRow() To $oRange.RangeAddress.EndRow()
			For $iC = $oRange.RangeAddress.StartColumn() To $oRange.RangeAddress.EndColumn()
				$oRange.Spreadsheet.getCellByPosition($iC, $iR).Value = Round(Random($nMin, $nMax), $iDecPlc)
				Sleep((IsInt($iR / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
			Next
		Next
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFillRandom

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFillSeries
; Description ...: Fill a Range of Cells with Data.
; Syntax ........: _LOCalc_RangeFillSeries(ByRef $oRange, $iDirection, $iMode, $nStep, $nEnd[, $iDateMode = $LOC_FILL_DATE_MODE_DAY])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iDirection          - an integer value (0-3). The Direction of the Series Fill. See Constants $LOC_FILL_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iMode               - an integer value (0-4). The Fill Type. See Constants $LOC_FILL_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $nStep               - a general number value. The amount the beginning value increments per step.
;                  $nEnd                - a general number value. The maximum Value the Fill series can insert.
;                  $iDateMode           - [optional] an integer value (0-3). Default is $LOC_FILL_DATE_MODE_DAY. The mode to calculate dates if $iMode is set to $LOC_FILL_MODE_DATE. See Constants $LOC_FILL_DATE_MODE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iDirection not an Integer, less than 0 or greater than 3. See Constants $LOC_FILL_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iMode not an Integer, less than 0 or greater than 4. See Constants $LOC_FILL_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $nStep not a Number value.
;                  @Error 1 @Extended 5 Return 0 = $nEnd not a Number value.
;                  @Error 1 @Extended 6 Return 0 = $iDateMode not an Integer, less than 0 or greater than 3. See Constants $LOC_FILL_DATE_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Fill series was successfully processed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeFill
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFillSeries(ByRef $oRange, $iDirection, $iMode, $nStep, $nEnd, $iDateMode = $LOC_FILL_DATE_MODE_DAY)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iDirection, $LOC_FILL_DIR_DOWN, $LOC_FILL_DIR_LEFT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iMode, $LOC_FILL_MODE_SIMPLE, $LOC_FILL_MODE_AUTO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsNumber($nStep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsNumber($nEnd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not __LO_IntIsBetween($iDateMode, $LOC_FILL_DATE_MODE_DAY, $LOC_FILL_DATE_MODE_YEAR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oRange.fillSeries($iDirection, $iMode, $iDateMode, $nStep, $nEnd)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFillSeries

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFilter
; Description ...: Apply a Filter to a Range.
; Syntax ........: _LOCalc_RangeFilter(ByRef $oRange, ByRef $oFilterDesc)
; Parameters ....: $oRange              - [in/out] an object. The Range to Filter. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oFilterDesc         - [in/out] an object. A Filter Descriptor created by a previous _LOCalc_FilterDescriptorCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFilterDesc not an Object.
;                  @Error 1 @Extended 3 Return 0 = Object called in $oFilterDesc not a Filter Descriptor.
;                  @Error 1 @Extended 4 Return ? = Column called in one Filter Field is greater than number of columns in the Range. Returning FilterFields Array element containing bad Filter Field, as an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Filter Fields array from Filter Descriptor.
;                  @Error 3 @Extended 2 Return 0 = Failed to get count of columns contained in Range.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully processed Filter operation.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeFilterClear, _LOCalc_FilterDescriptorCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFilter(ByRef $oRange, ByRef $oFilterDesc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iColumns
	Local $atFilterField[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFilterDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oFilterDesc.supportsService("com.sun.star.sheet.SheetFilterDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$atFilterField = $oFilterDesc.getFilterFields2()
	If Not IsArray($atFilterField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iColumns = $oRange.Columns.Count()
	If Not IsInt($iColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($atFilterField) - 1
		If Not __LO_IntIsBetween($atFilterField[$i].Field(), 0, $iColumns - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
	Next

	$oRange.filter($oFilterDesc)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFilter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFilterAdvanced
; Description ...: Apply an advanced filter to a Range.
; Syntax ........: _LOCalc_RangeFilterAdvanced(ByRef $oRange, ByRef $oFilterDescRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oFilterDescRange    - [in/out] an object. The Range containing the Filter Criteria. See remarks. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFilterDescRange not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Filter Descriptor.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Range was successfully filtered.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oFilterDescRange will be a range containing the filter criteria as described in the L.O. help file for Advanced Filters. It can be from anywhere in the same Calc Document, the same Sheet, or a completely different sheet. Named Ranges can also be used.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFilterAdvanced(ByRef $oRange, ByRef $oFilterDescRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFilterDesc

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFilterDescRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oFilterDesc = $oFilterDescRange.createFilterDescriptorByObject($oRange)
	If Not IsObj($oFilterDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oRange.filter($oFilterDesc)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFilterAdvanced

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFilterClear
; Description ...: Clear any previous filters for a Range.
; Syntax ........: _LOCalc_RangeFilterClear(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. The Range to clear filtering for. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a new, blank, Filter Descriptor.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared any old Filters for the Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeFilter
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFilterClear(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFilterDesc

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oFilterDesc = $oRange.createFilterDescriptor(True)
	If Not IsObj($oFilterDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oRange.filter($oFilterDesc) ; Calling filter with a blank Filter Desc clears any old filters applied.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFilterClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFindAll
; Description ...: Find all matches contained in a document of a specified Search String.
; Syntax ........: _LOCalc_RangeFindAll(ByRef $oRange, ByRef $oSrchDescript, $sSearchString)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or regular expression to search for.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSrchDescriptObject not a Search Descriptor Object.
;                  @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Search did not return an Object, something went wrong.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Search was Successful, but found no results.
;                  @Error 0 @Extended ? Return Array = Success. Search was Successful, returning 1 dimensional array containing the objects to each match, @Extended is set to the number of matches.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Objects returned are Ranges and can be used in any of the functions accepting a Range Object etc., to modify their properties or even the text itself.
;                  Only the Sheet that contains the Range is searched, to search all Sheets you will have to cycle through and perform a search for each.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_RangeFindNext, _LOCalc_RangeReplaceAll, _LOCalc_RangeReplace
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFindAll(ByRef $oRange, ByRef $oSrchDescript, $sSearchString)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResults
	Local $aoResults[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSrchDescript.SearchString = $sSearchString

	$oResults = $oRange.findAll($oSrchDescript)
	If Not IsObj($oResults) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oResults.getCount() > 0) Then
		ReDim $aoResults[$oResults.getCount]
		For $i = 0 To $oResults.getCount() - 1
			$aoResults[$i] = $oResults.getByIndex($i)
			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	Return (UBound($aoResults) > 0) ? (SetError($__LO_STATUS_SUCCESS, UBound($aoResults), $aoResults)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeFindAll

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFindNext
; Description ...: Find a Search String in a Document once or one at a time.
; Syntax ........: _LOCalc_RangeFindNext(ByRef $oRange, ByRef $oSrchDescript, $sSearchString[, $oLastFind = Null])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $oLastFind           - [optional] an object. Default is Null. The last returned Object by a previous call to this function to begin the search from, if called with Null, the search begins at the start of the Range.
; Return values .: Success: Object or 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;                  @Error 1 @Extended 5 Return 0 = $oLastFind not an Object, or failed to retrieve starting position from $oRange.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Search was successful but found no matches.
;                  @Error 0 @Extended 1 Return Object = Success. Search was successful, returning the resulting Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Object returned is a Range and can be used in any of the functions accepting a Range Object etc., to modify their properties or even the text itself.
;                  Only the Sheet that contains the Range is searched, to search all Sheets you will have to cycle through and perform a search for each.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_RangeFindAll, _LOCalc_RangeReplaceAll, _LOCalc_RangeReplace
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFindNext(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $oLastFind = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResult, $oFindRange

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($oLastFind <> Null) Then ; If Last find is set, set search start for beginning or end of last result, depending SearchBackwards value.
		If Not IsObj($oLastFind) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		; If Search Backwards is False, then retrieve the end of the last result's range, else get the Start.
		$oFindRange = ($oSrchDescript.SearchBackwards() = False) ? ($oRange.getCellByPosition($oLastFind.RangeAddress.EndColumn(), $oLastFind.RangeAddress.EndRow())) : ($oRange.getCellByPosition($oLastFind.RangeAddress.StartColumn(), $oLastFind.RangeAddress.StartRow()))
	EndIf

	$oSrchDescript.SearchString = $sSearchString

	If IsObj($oLastFind) Then
		$oResult = $oRange.findNext($oFindRange, $oSrchDescript)

	Else ; If a search hasn't been done before, FindFirst must be used or a result could be missed in the first cell.
		$oResult = $oRange.findFirst($oSrchDescript)
	EndIf

	Return (IsObj($oResult)) ? (SetError($__LO_STATUS_SUCCESS, 1, $oResult)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeFindNext

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFormula
; Description ...: Set or Retrieve Formulas in a Range.
; Syntax ........: _LOCalc_RangeFormula(ByRef $oRange[, $aasFormulas = Null[, $bStrictSize = False]])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $aasFormulas         - [optional] an array or arrays containing strings. Default is Null. An Array of Arrays containing formula strings to fill the range with. See remarks.
;                  $bStrictSize         - [optional] a boolean value. Default is False. If True, The Range size must explicitly match the array sizing. If False, The Range will be resized right or down to fit the Array sizing.
; Return values .: Success: 1 or Array
;                  Failure: 0 or ? and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $aasFormulas not an Array.
;                  @Error 1 @Extended 3 Return 0 = $bStrictSize not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bStrictSize called with True, and $aasFormulas array contains less or more elements than number of rows contained in the cell range.
;                  @Error 1 @Extended 5 Return ? = Element of $aasFormulas does not contain an array. Returning array element number of $aasFormulas containing error.
;                  @Error 1 @Extended 6 Return ? = $bStrictSize called with True, and Array contained in $aasFormulas has less or more elements than number of columns in the cell range. Returning array element number of $aasFormulas containing faulty array.
;                  @Error 1 @Extended 7 Return ? = $bStrictSize called with False, and Array contained in $aasFormulas has less or more elements than first Array contained in $aasFormulas. Returning array element number of $aasFormulas containing faulty array.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve array of Formula Data contained in the Cell Range.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Start of Row from Cell Range.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve End of Row from Cell Range.
;                  @Error 3 @Extended 4 Return 0 = Expanding Range would exceed number of Rows contained in Sheet.
;                  @Error 3 @Extended 5 Return 0 = Failed to re-size Cell Range Rows.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Start of Column from Cell Range.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve End of Column from Cell Range.
;                  @Error 3 @Extended 8 Return 0 = Expanding Range would exceed number of Columns contained in Sheet.
;                  @Error 3 @Extended 9 Return 0 = Failed to re-size Cell Range Columns.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Formulas were successfully set for the cell range.
;                  @Error 0 @Extended 1 Return Array of Arrays = Success. $aasFormulas called with Null, returning an array containing arrays, which contain any Formula content contained in the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return only formulas contained in the cell range when $aasFormulas is called with Null keyword. Array will be an array of arrays. The internal arrays will contain blank cells or formula strings, depending on cell content.
;                  $aasFormulas must be an array containing arrays. If $bStrictSize is called with True, the main Array's element count must match the row count contained in the Cell Range, and each internal Array's element count must match the column count of the Cell Range it is to fill. All internal arrays must be the same size.
;                  Any data previously contained in the Cell Range will be overwritten.
;                  All array elements must contain strings, blank or otherwise.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFormula(ByRef $oRange, $aasFormulas = Null, $bStrictSize = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iStart, $iEnd

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($aasFormulas) Then
		$aasFormulas = $oRange.getFormulaArray()
		If Not IsArray($aasFormulas) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $aasFormulas)
	EndIf

	If Not IsArray($aasFormulas) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bStrictSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	; Determine if the Array is sized appropriately
	$iStart = $oRange.RangeAddress.StartRow()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iEnd = $oRange.RangeAddress.EndRow()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $bStrictSize Then ; If Array is wrongly sized, return an error.
		If (UBound($aasFormulas) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Else ; Expand the Range to fit the Array
		If (UBound($aasFormulas) <> ($iEnd - $iStart + 1)) Then
			If (($oRange.RangeAddress.StartRow() + UBound($aasFormulas)) > $oRange.Spreadsheet.getRows.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; Check if resizing range is possible.

			$oRange = $oRange.Spreadsheet.getCellRangeByPosition($oRange.RangeAddress.StartColumn(), $oRange.RangeAddress.StartRow(), $oRange.RangeAddress.EndColumn(), ($oRange.RangeAddress.StartRow() + UBound($aasFormulas) - 1))
			If Not IsObj($oRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
		EndIf
	EndIf

	$iStart = $oRange.RangeAddress.StartColumn()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	$iEnd = $oRange.RangeAddress.EndColumn()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

	If $bStrictSize Then ; Check if the internal arrays are sized correctly, return error if not.

		For $i = 0 To UBound($aasFormulas) - 1
			If Not IsArray($aasFormulas[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
			If (UBound($aasFormulas[$i]) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)
		Next

	Else ; Check if the internal arrays are sized correctly, resize range if not.
		For $i = 0 To UBound($aasFormulas) - 1
			If Not IsArray($aasFormulas[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
			If (UBound($aasFormulas[$i]) <> UBound($aasFormulas[0])) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $i) ; If all arrays aren't same size as first array, then error.
		Next

		If (UBound($aasFormulas[0]) <> ($iEnd - $iStart + 1)) Then ; Resize the Range.
			If (($oRange.RangeAddress.StartColumn() + UBound($aasFormulas[0])) > $oRange.Spreadsheet.getColumns.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

			$oRange = $oRange.Spreadsheet.getCellRangeByPosition($oRange.RangeAddress.StartColumn(), $oRange.RangeAddress.StartRow(), ($oRange.RangeAddress.StartColumn() + UBound($aasFormulas[0]) - 1), $oRange.RangeAddress.EndRow())
			If Not IsObj($oRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 9, 0)
		EndIf
	EndIf

	$oRange.setFormulaArray($aasFormulas)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetAddressAsName
; Description ...: Retrieve the Name of the beginning and ending cells of the range.
; Syntax ........: _LOCalc_RangeGetAddressAsName(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Retrieve Range Address.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Successfully retrieved Range's address, returning it as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Return will be like the following, including the dollar signs. "$Sheet1.$A$1:$F$18"
; Related .......: _LOCalc_RangeGetAddressAsPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetAddressAsName(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sName = $oRange.AbsoluteName()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOCalc_RangeGetAddressAsName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetAddressAsPosition
; Description ...: Retrieve the Position of the beginning and ending cells of the range.
; Syntax ........: _LOCalc_RangeGetAddressAsPosition(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Retrieve Range Address.
;                  --Success--
;                  @Error 0 @Extended 0 Return Array = Success. Successfully retrieved Range's address, returning it as a 5 element Array. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The return will be a 5 element array giving the Range's address in the following order: Sheet index number, Range's first Cell Column, First Cell Row, Last Cell Column, Last Cell Row.
; Related .......: _LOCalc_RangeGetAddressAsName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetAddressAsPosition(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tRangeAddr
	Local $aiAddress[5]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tRangeAddr = $oRange.RangeAddress()
	If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aiAddress[0] = $tRangeAddr.Sheet()
	$aiAddress[1] = $tRangeAddr.StartColumn()
	$aiAddress[2] = $tRangeAddr.StartRow()
	$aiAddress[3] = $tRangeAddr.EndColumn()
	$aiAddress[4] = $tRangeAddr.EndRow()

	Return SetError($__LO_STATUS_SUCCESS, 0, $aiAddress)
EndFunc   ;==>_LOCalc_RangeGetAddressAsPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetCellByName
; Description ...: Retrieve a Cell or Cell Range Object by Cell name.
; Syntax ........: _LOCalc_RangeGetCellByName(ByRef $oRange, $sFromCellName[, $sToCellName = Null])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sFromCellName       - a string value. The cell to retrieve the Object for, or to begin the Cell Range. See remarks.
;                  $sToCellName         - [optional] a string value. Default is Null. The cell to end the Cell Range at.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFromCellName not a String.
;                  @Error 1 @Extended 3 Return 0 = $sToCellName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve requested Cell or Cell Range Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully retrieved and returning requested Cell or Cell Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFromCellName can be a Cell Name or a defined Cell Range name.
; Related .......: _LOCalc_RangeGetCellByPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetCellByName(ByRef $oRange, $sFromCellName, $sToCellName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCellRange
	Local $sCellRange = $sFromCellName

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFromCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($sToCellName <> Null) And Not IsString($sToCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($sToCellName <> Null) Then $sCellRange &= ":" & $sToCellName

	$oCellRange = $oRange.getCellRangeByName($sCellRange)
	If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCellRange)
EndFunc   ;==>_LOCalc_RangeGetCellByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetCellByPosition
; Description ...: Retrieve a Cell or Cell Range Object by position.
; Syntax ........: _LOCalc_RangeGetCellByPosition(ByRef $oRange, $iColumn, $iRow[, $iToColumn = Null[, $iToRow = Null]])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iColumn             - an integer value. The Column of the desired cell, or of the beginning of the Cell range. 0 Based. See remarks.
;                  $iRow                - an integer value. The Row of the desired cell, or of the beginning of the Cell range. 0 Based. See remarks.
;                  $iToColumn           - [optional] an integer value. Default is Null. The Column of the end of the Cell range. 0 Based. Must be greater or equal to $iColumn.
;                  $iToRow              - [optional] an integer value. Default is Null. The Row of the end of the Cell range. 0 Based. Must be greater or equal to $iRow.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, less than 0 or greater than number of Columns contained in the Range.
;                  @Error 1 @Extended 3 Return 0 = $iRow not an Integer, less than 0 or greater than number of Rows contained in the Range.
;                  @Error 1 @Extended 4 Return 0 = $iToColumn not an Integer, less than 0 or greater than number of Columns contained in the Range.
;                  @Error 1 @Extended 5 Return 0 = $iToRow not an Integer, less than 0 or greater than number of Rows contained in the Range.
;                  @Error 1 @Extended 6 Return 0 = $iToColumn less than $iColumn.
;                  @Error 1 @Extended 7 Return 0 = $iToRow less than $iRow.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve an individual Cell's Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve a Cell Range's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully retrieved and returned an Individual Cell's Object.
;                  @Error 0 @Extended 1 Return Object = Success. Successfully retrieved and returned a Cell Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: According to the wiki (https://wiki.documentfoundation.org/Faq/Calc/022), the maximum Columns contained in a sheet is 1024 until version 7.3, or 16384 from 7.3. and up..
;                  According to Andrew Pitonyak, (OOME. 4.1 Page 492), the maximum number of rows contained in a sheet is 65,536 as of OOo Calc 3.0, but according to the wiki (https://wiki.documentfoundation.org/Faq/Calc/022), the maximum or Rows for Libre Office Calc is 1,048,576.
; Related .......: _LOCalc_RangeGetCellByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetCellByPosition(ByRef $oRange, $iColumn, $iRow, $iToColumn = Null, $iToRow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell, $oCellRange

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iColumn, 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iRow, 0, $oRange.Rows.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($iToColumn <> Null) Or ($iToRow <> Null) Then
		If Not __LO_IntIsBetween($iToColumn, 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If Not __LO_IntIsBetween($iToRow, 0, $oRange.Rows.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If ($iToColumn < $iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If ($iToRow < $iRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	EndIf

	If __LO_VarsAreNull($iToColumn, $iToRow) Then
		$oCell = $oRange.getCellByPosition($iColumn, $iRow)
		If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $oCell)

	Else
		$oCellRange = $oRange.getCellRangeByPosition($iColumn, $iRow, $iToColumn, $iToRow)
		If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $oCellRange)
	EndIf
EndFunc   ;==>_LOCalc_RangeGetCellByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetSheet
; Description ...: Return the Sheet Object that contains the Range.
; Syntax ........: _LOCalc_RangeGetSheet(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Retrieve Sheet Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully retrieved Range's parent Sheet, returning the Sheet's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetSheet(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheet = $oRange.Spreadsheet()
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheet)
EndFunc   ;==>_LOCalc_RangeGetSheet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGroup
; Description ...: Group or Ungroup cells in a Range.
; Syntax ........: _LOCalc_RangeGroup(ByRef $oRange[, $iOrientation = $LOC_GROUP_ORIENT_ROWS[, $bGroup = True]])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iOrientation        - [optional] an integer value (0-1). Default is $LOC_GROUP_ORIENT_ROWS. Whether to Group Rows or Columns. See Constants $LOC_GROUP_ORIENT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bGroup              - [optional] a boolean value. Default is True. If True Cells are Grouped, if False, cells are Ungrouped.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants $LOC_GROUP_ORIENT_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $bGroup not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Range Address Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Group or Ungroup was successfully processed for range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGroup(ByRef $oRange, $iOrientation = $LOC_GROUP_ORIENT_ROWS, $bGroup = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tRangeAddress

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iOrientation, $LOC_GROUP_ORIENT_COLUMNS, $LOC_GROUP_ORIENT_ROWS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bGroup) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tRangeAddress = $oRange.RangeAddress()
	If Not IsObj($tRangeAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $bGroup Then
		$oRange.Spreadsheet.Group($tRangeAddress, $iOrientation)

	Else
		$oRange.Spreadsheet.Ungroup($tRangeAddress, $iOrientation)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeGroup

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeInsert
; Description ...: Insert blank cells at a Cell Range.
; Syntax ........: _LOCalc_RangeInsert(ByRef $oSheet, $oRange, $iMode)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRange              - an object. A Cell or Cell Range to insert new blank cells at. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iMode               - an integer value (0-4). The Cell Insertion Mode. See Constants $LOC_CELL_INSERT_MODE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iMode not an Integer, less than 0 or greater than 4. See Constants $LOC_CELL_INSERT_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Range Address Struct from $oRange.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Blank cells were successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The new range of cells inserted will be the same size as the range called in $oRange.
;                  Non-Empty cells cannot be moved off of the sheet.
;                  This function will silently fail if the insertion will cause an array formula to be split -- OOME. 4.1., Page 509.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeInsert(ByRef $oSheet, $oRange, $iMode)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tCellAddr

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iMode, $LOC_CELL_INSERT_MODE_NONE, $LOC_CELL_INSERT_MODE_COLUMNS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oRange.RangeAddress()
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSheet.insertCells($tCellAddr, $iMode)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeIsMerged
; Description ...: Check if any part of a range contains merged cells.
; Syntax ........: _LOCalc_RangeIsMerged(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to test if Range is Merged.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if Range is merged, else False. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return True only in the following cases: If the called Range covers the entire area of a merged range of cells, OR if the top-left most cell of a merged range of cells is called alone, or included in the Range.
; Related .......: _LOCalc_RangeMerge
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeIsMerged(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bMerged

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bMerged = $oRange.getIsMerged()
	If Not IsBool($bMerged) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bMerged)
EndFunc   ;==>_LOCalc_RangeIsMerged

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeMerge
; Description ...: Merge or Unmerge a Range of cells.
; Syntax ........: _LOCalc_RangeMerge(ByRef $oRange, $bMerge)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bMerge              - a boolean value. If True, the Cells within the range are merged. If False, any merged cells intercepting the Range will be unmurged. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bMerge not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Range was successfully merged or unmerged.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Any merged cells that are part of the original merge will be unmerged, even if they aren't contained in the called range, as long as the top-left most cell of the merged range is contained in the called range, i.e., I merge Range A1:C5, if I then attempt to unmerge A1:A5, the entire range of A1:C5 will be unmerged, but if I attempt to unmerge B1:C3, nothing will be unmerged.
; Related .......: _LOCalc_RangeIsMerged
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeMerge(ByRef $oRange, $bMerge)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bMerge) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRange.Merge($bMerge)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeMerge

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeNamedAdd
; Description ...: Add a Named Range to a specific Scope.
; Syntax ........: _LOCalc_RangeNamedAdd(ByRef $oObj, $vRange, $sName[, $iOptions = $LOC_NAMED_RANGE_OPT_NONE[, $oRefCell = Null]])
; Parameters ....: $oObj                - [in/out] an object. See remarks. A Document or Sheet object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, _LOCalc_DocCreate, _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $vRange              - a variant value. See remarks. May be a String or a Cell Range object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sName               - a string value. The unique name of the Named Range to create. Must start with a letter or an Underscore, and ONLY contain Letters, Numbers and Underscores, no Spaces.
;                  $iOptions            - [optional] an integer value (0-15). Default is $LOC_NAMED_RANGE_OPT_NONE. Any options to set for the Named Range, can be BitOR'd together. See Constants $LOC_NAMED_RANGE_OPT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $oRefCell            - [optional] an object. Default is Null. The reference cell for the Range or Formula set in $vRange.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $vRange not an Object and not a String.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sName contains invalid characters.
;                  @Error 1 @Extended 5 Return 0 = $iOptions not an Integer, less than 0 or greater than 15 (all constants added together). See Constants $LOC_NAMED_RANGE_OPT_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $vRange is a String and $oRefCell is not an Object.
;                  @Error 1 @Extended 7 Return 0 = Scope called in $oObj already contains a Named Range named the same as called in $sName.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.table.CellAddress" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Named Ranges Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Absolute Name of Range called in $vRange.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve new Named Range's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully added a new Named Range, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Object called in $oObj determines the scope you are inserting the new Named Range in, either Globally (Document Object), or locally (Sheet Object).
;                  $vRange can be a string representation of the Range covered by the NamedRange, i.e., $Sheet1.$A$1:$C$14, or a Formula, such as A1+A2, or a Cell Range Object.
;                  If $vRange is a String, $oRefCell must be called with the Cell Object of either the first cell of the desired Range, or the reference cell for the formula. See explanation below.
;                  $oRefCell "acts as the base address for cells referenced in a relative way. If the cell range is not specified as an absolute address, the referenced range will be different based on where in the spreadsheet the range is used."
;                  Or in the case of a formula, an example would if we created a "named range 'AddLeft', which refers to the equation A3+B3 with C3 as the reference cell. The cells A3 and B3 are the two cells directly to the left of C3, so, the equation =AddLeft calculates the sum of the two cells directly to the left of the cell that contains the equation. Changing the reference cell to C4, which is below A3 and B3, causes the AddLeft equation to calculate the sum of the two cells that are to the left on the previous row."
;                  [Both quotations above are adapted from Andrew Pitonyak's book OOME 4.1, pdf Page 523, book page 519]
; Related .......: _LOCalc_RangeNamedDelete, _LOCalc_RangeNamedExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeNamedAdd(ByRef $oObj, $vRange, $sName, $iOptions = $LOC_NAMED_RANGE_OPT_NONE, $oRefCell = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPLEADING = 1, $__STR_STRIPTRAILING = 2
	Local $oNamedRanges, $oNamedRange
	Local $sRange
	Local $tCellAddr

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($vRange) And Not IsString($vRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$sName = StringStripWS($sName, ($__STR_STRIPLEADING + $__STR_STRIPTRAILING))
	If StringRegExp($sName, "[^a-zA-Z0-9_]") Or StringRegExp($sName, "^[^a-zA-Z_]") Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iOptions, $LOC_NAMED_RANGE_OPT_NONE, 15) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; 15 = all flags added together.
	If IsString($vRange) And Not IsObj($oRefCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oNamedRanges = $oObj.NamedRanges()
	If Not IsObj($oNamedRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oNamedRanges.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$tCellAddr = __LO_CreateStruct("com.sun.star.table.CellAddress")
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If IsObj($vRange) Then
		If IsObj($oRefCell) Then
			$tCellAddr.Sheet = $oRefCell.RangeAddress.Sheet()
			$tCellAddr.Column = $oRefCell.RangeAddress.StartColumn()
			$tCellAddr.Row = $oRefCell.RangeAddress.StartRow()

		Else
			$tCellAddr.Sheet = $vRange.RangeAddress.Sheet()
			$tCellAddr.Column = $vRange.RangeAddress.StartColumn()
			$tCellAddr.Row = $vRange.RangeAddress.StartRow()
		EndIf

		$sRange = $vRange.AbsoluteName()
		If Not IsString($sRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Else
		$tCellAddr.Sheet = $oRefCell.RangeAddress.Sheet()
		$tCellAddr.Column = $oRefCell.RangeAddress.StartColumn()
		$tCellAddr.Row = $oRefCell.RangeAddress.StartRow()

		$sRange = $vRange
	EndIf

	$oNamedRanges.addNewByName($sName, $sRange, $tCellAddr, $iOptions)

	$oNamedRange = $oNamedRanges.getByName($sName)
	If Not IsObj($oNamedRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNamedRange)
EndFunc   ;==>_LOCalc_RangeNamedAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeNamedChangeScope
; Description ...: Change the scope a Named Range is located in.
; Syntax ........: _LOCalc_RangeNamedChangeScope(ByRef $oDoc, ByRef $oNamedRange, ByRef $oNewScope[, $sNewName = ""])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oNamedRange         - [in/out] an object. A Named Range Object returned by a previous _LOCalc_RangeNamedGetObjByName, or _LOCalc_RangeNamedAdd function.
;                  $oNewScope           - [in/out] an object. The new Scope to place the Named Range in. A Document or Sheet object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, _LOCalc_DocCreate, _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sNewName            - [optional] a string value. Default is "". A new name for the Range. Empty String means the name is reused.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oNamedRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oNewScope not an Object.
;                  @Error 1 @Extended 4 Return 0 = $sNewName not a String.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sNewName already exists in $oNewScope.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = $oNewScope already contains a Named Range with the same name as Range called in $oNamedRange.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Name of $oNamedRange.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Content of $oNamedRange.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Scope Object of $oNamedRange.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Reference Position of $oNamedRange.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Options applied to $oNamedRange.
;                  @Error 3 @Extended 7 Return 0 = Failed to remove Named Range from old Scope.
;                  @Error 3 @Extended 8 Return 0 = Failed to retrieve new Named Range Object in new scope.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully changed the scope of the Named Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeNamedModify, _LOCalc_RangeNamedExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeNamedChangeScope(ByRef $oDoc, ByRef $oNamedRange, ByRef $oNewScope, $sNewName = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTempNamedRange, $oObj
	Local $sNamedRange, $sContent
	Local $tRefPos
	Local $iType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNamedRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oNewScope) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($sNewName = "") Then
		If $oNewScope.NamedRanges.hasByName($oNamedRange.Name()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$sNewName = $oNamedRange.Name()

	Else
		If $oNewScope.NamedRanges.hasByName($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	EndIf

	$sNamedRange = $oNamedRange.Name()
	If Not IsString($sNamedRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$sContent = $oNamedRange.Content()
	If Not IsString($sContent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oObj = __LOCalc_NamedRangeGetScopeObj($oDoc, $sNamedRange, $oNamedRange.TokenIndex(), $oNamedRange.Content())
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tRefPos = $oNamedRange.ReferencePosition()
	If Not IsObj($tRefPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$iType = $oNamedRange.Type()
	If Not IsInt($iType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	$oObj.NamedRanges.removeByName($sNamedRange)
	If $oObj.NamedRanges.hasByName($sNamedRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

	$oNewScope.NamedRanges.addNewByName($sNewName, $sContent, $tRefPos, $iType)

	$oTempNamedRange = $oNewScope.NamedRanges.getByName($sNewName)
	If Not IsObj($oTempNamedRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

	$oNamedRange = $oTempNamedRange

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeNamedChangeScope

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeNamedDelete
; Description ...: Delete a Named Range from a particular scope.
; Syntax ........: _LOCalc_RangeNamedDelete(ByRef $oObj, $vNamedRange)
; Parameters ....: $oObj                - [in/out] an object. See remarks. A Document or Sheet object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, _LOCalc_DocCreate, _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $vNamedRange         - a variant value. The name of the Named Range to delete, as a string, or the NamedRange Object as returned from _LOCalc_RangeNamedAdd or _LOCalc_RangeNamedGetObjByName.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $vNamedRange not an Object and not a String.
;                  @Error 1 @Extended 3 Return 0 = Scope called in $oObj does not contain a Named Range as called in $vNamedRange.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Named Ranges Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Named Range's name.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete requested Named Range.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully deleted the requested Named Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Object called in $oObj must be the scope the Named Range is present in, either Globally (Document Object), or locally (Sheet Object).
; Related .......: _LOCalc_RangeNamedAdd, _LOCalc_RangeNamedExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeNamedDelete(ByRef $oObj, $vNamedRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNamedRanges
	Local $sNamedRange

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($vNamedRange) And Not IsObj($vNamedRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oNamedRanges = $oObj.NamedRanges()
	If Not IsObj($oNamedRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If IsObj($vNamedRange) Then
		$sNamedRange = $vNamedRange.Name()
		If Not IsString($sNamedRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Else
		$sNamedRange = $vNamedRange
	EndIf

	If Not $oNamedRanges.hasByName($sNamedRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oNamedRanges.removeByName($sNamedRange)

	If $oNamedRanges.hasByName($sNamedRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeNamedDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeNamedExists
; Description ...: Check if a Named Range exists in a particular scope.
; Syntax ........: _LOCalc_RangeNamedExists(ByRef $oObj, $sName)
; Parameters ....: $oObj                - [in/out] an object. See remarks. A Document or Sheet object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, _LOCalc_DocCreate, _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sName               - a string value. The Named Range name to look for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Named Ranges Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to query whether Scope contains the called name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the Scope contains a Named Range by the called name. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Object called in $oObj determines the scope you are searching in for the Named Range specified, either Globally (Document Object), or locally (Sheet Object).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeNamedExists(ByRef $oObj, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNamedRanges
	Local $bExists

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oNamedRanges = $oObj.NamedRanges()
	If Not IsObj($oNamedRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$bExists = $oNamedRanges.hasByName($sName)
	If Not IsBool($bExists) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bExists)
EndFunc   ;==>_LOCalc_RangeNamedExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeNamedGetNames
; Description ...: Retrieve an array of Named Range names for either the document or sheet.
; Syntax ........: _LOCalc_RangeNamedGetNames(ByRef $oObj)
; Parameters ....: $oObj                - [in/out] an object. See remarks. A Document or Sheet object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, _LOCalc_DocCreate, _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Named Ranges Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Named Ranges contained in the called scope. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Object called in $oObj determines the scope you are retrieving the array of names for, either Globally (Document Object), or locally (Sheet Object).
; Related .......: _LOCalc_RangeNamedGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeNamedGetNames(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNamedRanges
	Local $asNames[0]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oNamedRanges = $oObj.NamedRanges()
	If Not IsObj($oNamedRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asNames[$oNamedRanges.Count()]

	For $i = 0 To $oNamedRanges.Count() - 1
		$asNames[$i] = $oNamedRanges.getByIndex($i).Name()

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOCalc_RangeNamedGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeNamedGetObjByName
; Description ...: Retrieve a Named Range Object by Name.
; Syntax ........: _LOCalc_RangeNamedGetObjByName(ByRef $oObj, $sName)
; Parameters ....: $oObj                - [in/out] an object. See remarks. A Document or Sheet object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, _LOCalc_DocCreate, _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sName               - a string value. The name of the Named Range to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Scope called in $oObj does not contain a Named Range by the name called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Named Ranges Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve requested Named Range Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Named Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Object called in $oObj must be the scope the Named Range is present in, either Globally (Document Object), or locally (Sheet Object).
; Related .......: _LOCalc_RangeNamedGetNames, _LOCalc_RangeNamedExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeNamedGetObjByName(ByRef $oObj, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNamedRanges, $oNamedRange

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oNamedRanges = $oObj.NamedRanges()
	If Not IsObj($oNamedRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oNamedRanges.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oNamedRange = $oNamedRanges.getByName($sName)
	If Not IsObj($oNamedRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNamedRange)
EndFunc   ;==>_LOCalc_RangeNamedGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeNamedModify
; Description ...: Set or Retrieve the settings for a Named Range.
; Syntax ........: _LOCalc_RangeNamedModify(ByRef $oDoc, ByRef $oNamedRange[, $vRange = Null[, $sName = Null[, $iOptions = Null[, $oRefCell = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oNamedRange         - [in/out] an object. A Named Range Object returned by a previous _LOCalc_RangeNamedGetObjByName, or _LOCalc_RangeNamedAdd function.
;                  $vRange              - [optional] a variant value. Default is Null. See remarks. May be a String or a Cell Range object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sName               - [optional] a string value. Default is Null. The unique name of the Named Range to create. Must start with a letter or Underscore, and ONLY contain Letters, Numbers and Underscores, no Spaces.
;                  $iOptions            - [optional] an integer value (0-15). Default is Null. Any options to set for the Named Range, can be BitOR'd together. See Constants $LOC_NAMED_RANGE_OPT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $oRefCell            - [optional] an object. Default is Null. The reference cell for the Range or Formula set in $vRange.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oNamedRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $vRange not an Object and not a String.
;                  @Error 1 @Extended 4 Return 0 = $sName not a String.
;                  @Error 1 @Extended 5 Return 0 = $sName contains invalid characters.
;                  @Error 1 @Extended 6 Return 0 = Scope containing Named Range already has a Named Range with the name as called in $sName.
;                  @Error 1 @Extended 7 Return 0 = $iOptions not an Integer, less than 0 or greater than 15 (all constants added together). See Constants $LOC_NAMED_RANGE_OPT_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $oRefCell not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Cell Object referenced by this Named Range.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Reference Position of Named Range.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Named Range's Scope Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $vRange
;                  |                               2 = Error setting $sName
;                  |                               4 = Error setting $iOptions
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $vRange can be a string representation of the Range covered by the NamedRange, i.e., $Sheet1.$A$1:$C$14, or a Formula, such as A1+A2, or a Cell Range Object.
;                  If $vRange is a String, $oRefCell must be called with the Cell Object of either the first cell of the desired Range, or the reference cell for the formula. See explanation below.
;                  $oRefCell "acts as the base address for cells referenced in a relative way. If the cell range is not specified as an absolute address, the referenced range will be different based on where in the spreadsheet the range is used."
;                  Or in the case of a formula, an example would if we created a "named range 'AddLeft', which refers to the equation A3+B3 with C3 as the reference cell. The cells A3 and B3 are the two cells directly to the left of C3, so, the equation =AddLeft calculates the sum of the two cells directly to the left of the cell that contains the equation. Changing the reference cell to C4, which is below A3 and B3, causes the AddLeft equation to calculate the sum of the two cells that are to the left on the previous row."
;                  [Both quotations above are adapted from Andrew Pitonyak's book OOME 4.1, pdf Page 523, book page 519.]
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  When retrieving the settings, $vRange will be in a String format, either being a formula or Range Address String, i.e. $Sheet1.$A$1:$C$14.
;                  When retrieving the settings, $oRefCell will be a Cell Object.
; Related .......: _LOCalc_RangeNamedGetObjByName, _LOCalc_RangeNamedAdd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeNamedModify(ByRef $oDoc, ByRef $oNamedRange, $vRange = Null, $sName = Null, $iOptions = Null, $oRefCell = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPLEADING = 1, $__STR_STRIPTRAILING = 2
	Local $avNamedRange[4]
	Local $oObj
	Local $iError = 0
	Local $tCellAddr

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNamedRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($vRange, $sName, $iOptions, $oRefCell) Then
		$oRefCell = $oDoc.Sheets.getByIndex($oNamedRange.ReferencePosition.Sheet()).getCellByPosition($oNamedRange.ReferencePosition.Column(), $oNamedRange.ReferencePosition.Row())
		If Not IsObj($oRefCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LO_ArrayFill($avNamedRange, $oNamedRange.Content(), $oNamedRange.Name(), $oNamedRange.Type(), $oRefCell)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avNamedRange)
	EndIf

	If ($vRange <> Null) Then
		If Not IsObj($vRange) And Not IsString($vRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If IsObj($vRange) Then
			$tCellAddr = $oNamedRange.ReferencePosition()
			If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tCellAddr.Sheet = $vRange.RangeAddress.Sheet()
			$tCellAddr.Column = $vRange.RangeAddress.StartColumn()
			$tCellAddr.Row = $vRange.RangeAddress.StartRow()

			$oNamedRange.ReferencePosition = $tCellAddr
			$oNamedRange.Content = $vRange.AbsoluteName()
			$iError = ($oNamedRange.Content() = $vRange.AbsoluteName()) ? ($iError) : (BitOR($iError, 1))

		Else ; $vRange is String
			$oNamedRange.Content = $vRange
			$iError = ($oNamedRange.Content() = $vRange) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$sName = StringStripWS($sName, ($__STR_STRIPLEADING + $__STR_STRIPTRAILING))
		If StringRegExp($sName, "[^a-zA-Z0-9_]") Or StringRegExp($sName, "^[^a-zA-Z_]") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj = __LOCalc_NamedRangeGetScopeObj($oDoc, $oNamedRange.Name(), $oNamedRange.TokenIndex(), $oNamedRange.Content())
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If $oObj.NamedRanges.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oNamedRange.Name = $sName
		$iError = ($oNamedRange.Name() = $sName) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iOptions <> Null) Then
		If Not __LO_IntIsBetween($iOptions, $LOC_NAMED_RANGE_OPT_NONE, 15) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0) ; 15 = all flags added together.

		$oNamedRange.Type = $iOptions
		$iError = ($oNamedRange.Type() = $iOptions) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($oRefCell <> Null) Then
		If Not IsObj($oRefCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tCellAddr = $oNamedRange.ReferencePosition()
		If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$tCellAddr.Sheet = $oRefCell.RangeAddress.Sheet()
		$tCellAddr.Column = $oRefCell.RangeAddress.StartColumn()
		$tCellAddr.Row = $oRefCell.RangeAddress.StartRow()

		$oNamedRange.ReferencePosition = $tCellAddr
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeNamedModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeNumbers
; Description ...: Set or Retrieve Numbers in a Range.
; Syntax ........: _LOCalc_RangeNumbers(ByRef $oRange[, $aanNumbers = Null[, $bStrictSize = False]])
; Parameters ....: $oRange              - [in/out] an object. A cell or cell range to set or retrieve number values for. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $aanNumbers          - [optional] an array of arrays containing general numbers. Default is Null. An Array of Arrays containing numbers to fill the range with. See remarks.
;                  $bStrictSize         - [optional] a boolean value. Default is False. If True, The Range size must explicitly match the array sizing. If False, The Range will be resized right or down to fit the Array sizing.
; Return values .: Success: 1 or Array
;                  Failure: 0 or ? and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $aanNumbers not an Array.
;                  @Error 1 @Extended 3 Return 0 = $bStrictSize not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bStrictSize called with True, and $aanNumbers array contains less or more elements than number of rows contained in the cell range.
;                  @Error 1 @Extended 5 Return ? = Element of $aanNumbers does not contain an array. Returning array element number of $aanNumbers containing error.
;                  @Error 1 @Extended 6 Return ? = $bStrictSize called with True, and Array contained in $aanNumbers has less or more elements than number of columns in the cell range. Returning array element number of $aanNumbers containing faulty array.
;                  @Error 1 @Extended 7 Return ? = $bStrictSize called with False, and Array contained in $aanNumbers has less or more elements than first Array contained in $aanNumbers. Returning array element number of $aanNumbers containing faulty array.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve array of Formula Data contained in the Cell Range.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Start of Row from Cell Range.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve End of Row from Cell Range.
;                  @Error 3 @Extended 4 Return 0 = Expanding Range would exceed number of Rows contained in Sheet.
;                  @Error 3 @Extended 5 Return 0 = Failed to re-size Cell Range Rows.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Start of Column from Cell Range.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve End of Column from Cell Range.
;                  @Error 3 @Extended 8 Return 0 = Expanding Range would exceed number of Columns contained in Sheet.
;                  @Error 3 @Extended 9 Return 0 = Failed to re-size Cell Range Columns.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Values were successfully set for the cell range.
;                  @Error 0 @Extended 1 Return Array of Arrays = Success. $aanNumbers called with Null, returning an array containing arrays, which contain any numerical content contained in the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return only numbers contained in the cell range when $aanNumbers is called with Null keyword. Array will be an array of arrays. The internal arrays will contain blank cells or numbers, depending on cell content.
;                  $aanNumbers must be an array containing arrays. If $bStrictSize is called with True, the main Array's element count must match the row count contained in the Cell Range, and each internal Array's element count must match the column count of the Cell Range it is to fill. All internal arrays must be the same size.
;                  Any data previously contained in the Cell Range will be overwritten.
;                  All array elements must contain numbers.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeNumbers(ByRef $oRange, $aanNumbers = Null, $bStrictSize = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iStart, $iEnd

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($aanNumbers) Then
		$aanNumbers = $oRange.getData()
		If Not IsArray($aanNumbers) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $aanNumbers)
	EndIf

	If Not IsArray($aanNumbers) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bStrictSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	; Determine if the Array is sized appropriately
	$iStart = $oRange.RangeAddress.StartRow()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iEnd = $oRange.RangeAddress.EndRow()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $bStrictSize Then ; If Array is wrongly sized, return an error.
		If (UBound($aanNumbers) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Else ; Expand the Range to fit the Array
		If (UBound($aanNumbers) <> ($iEnd - $iStart + 1)) Then
			If (($oRange.RangeAddress.StartRow() + UBound($aanNumbers)) > $oRange.Spreadsheet.getRows.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; Check if resizing range is possible.

			$oRange = $oRange.Spreadsheet.getCellRangeByPosition($oRange.RangeAddress.StartColumn(), $oRange.RangeAddress.StartRow(), $oRange.RangeAddress.EndColumn(), ($oRange.RangeAddress.StartRow() + UBound($aanNumbers) - 1))
			If Not IsObj($oRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
		EndIf
	EndIf

	$iStart = $oRange.RangeAddress.StartColumn()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	$iEnd = $oRange.RangeAddress.EndColumn()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

	If $bStrictSize Then ; Check if the internal arrays are sized correctly, return error if not.

		For $i = 0 To UBound($aanNumbers) - 1
			If Not IsArray($aanNumbers[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
			If (UBound($aanNumbers[$i]) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)
		Next

	Else ; Check if the internal arrays are sized correctly, resize range if not.
		For $i = 0 To UBound($aanNumbers) - 1
			If Not IsArray($aanNumbers[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
			If (UBound($aanNumbers[$i]) <> UBound($aanNumbers[0])) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $i) ; If all arrays aren't same size as first array, then error.
		Next

		If (UBound($aanNumbers[0]) <> ($iEnd - $iStart + 1)) Then ; Resize the Range.
			If (($oRange.RangeAddress.StartColumn() + UBound($aanNumbers[0])) > $oRange.Spreadsheet.getColumns.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

			$oRange = $oRange.Spreadsheet.getCellRangeByPosition($oRange.RangeAddress.StartColumn(), $oRange.RangeAddress.StartRow(), ($oRange.RangeAddress.StartColumn() + UBound($aanNumbers[0]) - 1), $oRange.RangeAddress.EndRow())
			If Not IsObj($oRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 9, 0)
		EndIf
	EndIf

	$oRange.setData($aanNumbers)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeNumbers

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeOutlineClearAll
; Description ...: Clear all Outline groups for a Sheet.
; Syntax ........: _LOCalc_RangeOutlineClearAll(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Outlining successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeOutlineClearAll(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheet.clearOutline()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeOutlineClearAll

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeOutlineShow
; Description ...: Show Outlined groups of cells up a specific level in a Sheet.
; Syntax ........: _LOCalc_RangeOutlineShow(ByRef $oSheet, $iLevel[, $iOrientation = $LOC_GROUP_ORIENT_ROWS])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iLevel              - an integer value. The level of Outlines to show, beginning at 1 and continuing to the level input. Call 0 to collapse them all.
;                  $iOrientation        - [optional] an integer value (0-1). Default is $LOC_GROUP_ORIENT_ROWS. The orientation of the Outlines. See Constants $LOC_GROUP_ORIENT_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iLevel not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants $LOC_GROUP_ORIENT_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. command was successfully processed for the Sheet.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeOutlineShow(ByRef $oSheet, $iLevel, $iOrientation = $LOC_GROUP_ORIENT_ROWS)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iLevel, 0, $iLevel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iOrientation, $LOC_GROUP_ORIENT_COLUMNS, $LOC_GROUP_ORIENT_ROWS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSheet.showLevel($iLevel, $iOrientation)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeOutlineShow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotDelete
; Description ...: Delete a Pivot Table.
; Syntax ........: _LOCalc_RangePivotDelete(ByRef $oDoc, ByRef $oPivotTable)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an object.
;                  @Error 1 @Extended 2 Return 0 = $oPivotTable not an Object.
;                  @Error 1 @Extended 3 Return 0 = Document called in $oDoc does not contain the Pivot Table called in $oPivotTable.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the Pivot Table's parent Sheet.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Pivot Table's name.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete the Pivot Table.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Pivot Table was deleted successfully.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotDelete(ByRef $oDoc, ByRef $oPivotTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName
	Local $oSheet

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; Get the sheet that contains the Pivot Table.
	$oSheet = $oDoc.Sheets.getByIndex($oPivotTable.OutputRange.Sheet())
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sName = $oPivotTable.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oSheet.DataPilotTables.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Wrong Doc object called?

	$oSheet.DataPilotTables.removeByName($sName)
	If $oSheet.DataPilotTables.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangePivotDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotDest
; Description ...: Set or Retrieve the Pivot Table's Destination Range.
; Syntax ........: _LOCalc_RangePivotDest(ByRef $oDoc, ByRef $oPivotTable[, $oDestRange = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
;                  $oDestRange          - [optional] an object. Default is Null. The Range to output the Pivot Table to. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPivotTable not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oDestRange not an Object.
;                  @Error 1 @Extended 4 Return 0 = Range called in $oDestRange is within the source range.
;                  @Error 1 @Extended 5 Return 0 = Document called in $oDoc does not contain the Pivot Table called in $oPivotTable.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Data Pilot Descriptor Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create com.sun.star.table.CellAddress Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Source Range Parent Sheet.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Source Range Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Pivot Table's name.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Pivot Table Field Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to delete the original Pivot Table.
;                  @Error 3 @Extended 6 Return 0 = Failed to insert new Pivot Table.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve new Pivot Table Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oDestRange
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Object = Success. All optional parameters were called with Null, returning current destination Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I was unable to find a setting for "Show Expand/Collapse buttons", therefore the current setting will be lost, because to change the output range, the entire Pivot Table needs to be copied over and re-inserted.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Any existing data within the Destination range will be overwritten.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotDest(ByRef $oDoc, ByRef $oPivotTable, $oDestRange = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet, $oOrigRange, $oNewPivotTable, $oPivotDesc, $oPivotField
	Local $tCellAddr
	Local $sName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheet = $oDoc.Sheets.getByIndex($oPivotTable.OutputRange.Sheet())
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oOrigRange = $oSheet.getCellRangeByPosition($oPivotTable.OutputRange.StartColumn(), $oPivotTable.OutputRange.StartRow(), $oPivotTable.OutputRange.EndColumn(), $oPivotTable.OutputRange.EndRow())
	If Not IsObj($oOrigRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($oDestRange) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oOrigRange)

	If Not IsObj($oDestRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPivotDesc = $oDestRange.Spreadsheet.DataPilotTables.createDataPilotDescriptor()
	If Not IsObj($oPivotDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($oPivotTable.SourceRange.Sheet() = $oDestRange.RangeAddress.Sheet()) And _
			__LO_IntIsBetween($oDestRange.RangeAddress.StartColumn(), $oPivotTable.SourceRange.StartColumn(), $oPivotTable.SourceRange.EndColumn()) And _
			__LO_IntIsBetween($oDestRange.RangeAddress.StartRow(), $oPivotTable.SourceRange.StartRow(), $oPivotTable.SourceRange.EndRow()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$tCellAddr = __LO_CreateStruct("com.sun.star.table.CellAddress")
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$tCellAddr.Sheet = $oDestRange.RangeAddress.Sheet()
	$tCellAddr.Column = $oDestRange.RangeAddress.StartColumn()
	$tCellAddr.Row = $oDestRange.RangeAddress.StartRow()

	$sName = $oPivotTable.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If Not $oOrigRange.Spreadsheet.DataPilotTables.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	; Transfer properties to Pivot descriptor.
	$oPivotDesc.SourceRange = $oPivotTable.SourceRange()

	With $oPivotDesc
		.ColumnGrand = $oPivotTable.ColumnGrand()
		.RowGrand = $oPivotTable.RowGrand()
		.IgnoreEmptyRows = $oPivotTable.IgnoreEmptyRows()
		.RepeatIfEmpty = $oPivotTable.RepeatIfEmpty()
		.ShowFilterButton = $oPivotTable.ShowFilterButton()
		.DrillDownOnDoubleClick = $oPivotTable.DrillDownOnDoubleClick()
	EndWith

	For $i = 0 To $oPivotTable.ColumnFields.Count() - 1
		$oPivotField = $oPivotTable.ColumnFields.getByIndex($i)
		If Not IsObj($oPivotField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).Orientation = $oPivotField.Orientation()
		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).Function = $oPivotField.Function()
		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).ShowEmpty = $oPivotField.ShowEmpty()

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	For $i = 0 To $oPivotTable.RowFields.Count() - 1
		$oPivotField = $oPivotTable.RowFields.getByIndex($i)
		If Not IsObj($oPivotField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).Orientation = $oPivotField.Orientation()
		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).Function = $oPivotField.Function()
		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).ShowEmpty = $oPivotField.ShowEmpty()

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	For $i = 0 To $oPivotTable.PageFields.Count() - 1
		$oPivotField = $oPivotTable.PageFields.getByIndex($i)
		If Not IsObj($oPivotField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).Orientation = $oPivotField.Orientation()
		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).Function = $oPivotField.Function()
		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).ShowEmpty = $oPivotField.ShowEmpty()

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	For $i = 0 To $oPivotTable.DataFields.Count() - 1
		$oPivotField = $oPivotTable.DataFields.getByIndex($i)
		If Not IsObj($oPivotField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).Orientation = $oPivotField.Orientation()
		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).Function = $oPivotField.Function()
		$oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).ShowEmpty = $oPivotField.ShowEmpty()
		If IsObj($oPivotField.Reference()) Then $oPivotDesc.DataPilotFields.getByName($oPivotField.Name()).Reference = $oPivotField.Reference()

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	With $oPivotDesc.FilterDescriptor
		.ContainsHeader = $oPivotTable.FilterDescriptor.ContainsHeader()
		.CopyOutputData = $oPivotTable.FilterDescriptor.CopyOutputData()
		.FilterFields = $oPivotTable.FilterDescriptor.FilterFields()
		.FilterFields2 = $oPivotTable.FilterDescriptor.FilterFields2()
		.FilterFields3 = $oPivotTable.FilterDescriptor.FilterFields3()
		.IsCaseSensitive = $oPivotTable.FilterDescriptor.IsCaseSensitive()
		.Orientation = $oPivotTable.FilterDescriptor.Orientation()
		.OutputPosition = $oPivotTable.FilterDescriptor.OutputPosition()
		.SaveOutputPosition = $oPivotTable.FilterDescriptor.SaveOutputPosition()
		.SkipDuplicates = $oPivotTable.FilterDescriptor.SkipDuplicates()
		.UseRegularExpressions = $oPivotTable.FilterDescriptor.UseRegularExpressions()
	EndWith

	$oOrigRange.Spreadsheet.DataPilotTables.removeByName($sName)
	If $oOrigRange.Spreadsheet.DataPilotTables.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$oDestRange.Spreadsheet.DataPilotTables.insertNewByName($sName, $tCellAddr, $oPivotDesc)
	If Not $oDestRange.Spreadsheet.DataPilotTables.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	$oNewPivotTable = $oDestRange.Spreadsheet.DataPilotTables.getByName($sName)
	If Not IsObj($oNewPivotTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

	$oPivotTable = $oNewPivotTable

	If ($oPivotTable.Name() <> $sName) Then $oPivotTable.Name = $sName

	If Not (($oPivotTable.OutputRange.Sheet() = $oDestRange.RangeAddress.Sheet()) And _
			($oPivotTable.OutputRange.StartColumn() = $oDestRange.RangeAddress.StartColumn()) And _
			($oPivotTable.OutputRange.StartRow() = $oDestRange.RangeAddress.StartRow())) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangePivotDest

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotExists
; Description ...: Query if a Pivot Table with a specific name exists in a Sheet.
; Syntax ........: _LOCalc_RangePivotExists(ByRef $oSheet, $sName)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sName               - a string value. The Pivot Table name to look for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query Sheet for Pivot Table name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning Boolean whether the Sheet contains a Pivot Table with the called name (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotExists(ByRef $oSheet, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$bReturn = $oSheet.DataPilotTables.hasByName($sName)
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOCalc_RangePivotExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFieldGetObjByName
; Description ...: Retrieve an Object for one of the Pivot Table Fields by Name.
; Syntax ........: _LOCalc_RangePivotFieldGetObjByName(ByRef $oPivotTable, $sName)
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
;                  $sName               - a string value. The Pivot Field name to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Pivot Table called in $oPivotTable does not contain a Field with name called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Pivot Table Field Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Pivot Table Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFieldGetObjByName(ByRef $oPivotTable, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPivotField

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPivotTable.DataPilotFields.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPivotField = $oPivotTable.DataPilotFields.getByName($sName)
	If Not IsObj($oPivotField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPivotField)
EndFunc   ;==>_LOCalc_RangePivotFieldGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFieldItemsGetNames
; Description ...: Retrieve an array of Item names contained in a Field.
; Syntax ........: _LOCalc_RangePivotFieldItemsGetNames(ByRef $oPivotField)
; Parameters ....: $oPivotField         - [in/out] an object. A Pivot Table Field object returned by a previous _LOCalc_RangePivotFieldGetObjByName function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotField not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve array of Item names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning array of Item names contained in the Column/Field. @Extended is set to the number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The element names are the items contained in each row for a specific column/field.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFieldItemsGetNames(ByRef $oPivotField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asPivotFieldItems[0]

	If Not IsObj($oPivotField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asPivotFieldItems = $oPivotField.Items.ElementNames()
	If Not IsArray($asPivotFieldItems) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asPivotFieldItems), $asPivotFieldItems)
EndFunc   ;==>_LOCalc_RangePivotFieldItemsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFieldsColumnsGetNames
; Description ...: Retrieve an array of Field Names set as Column Fields.
; Syntax ........: _LOCalc_RangePivotFieldsColumnsGetNames(ByRef $oPivotTable)
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of Pivot Table Fields.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Pivot Table Field Name.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Pivot Table Field Names currently set as Column Fields, contained in the Pivot Table. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFieldsColumnsGetNames(ByRef $oPivotTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]
	Local $iCount

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oPivotTable.ColumnFields.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asNames[$iCount]

	For $i = 0 To $iCount - 1
		$asNames[$i] = $oPivotTable.ColumnFields.getByIndex($i).Name()
		If Not IsString($asNames[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asNames)
EndFunc   ;==>_LOCalc_RangePivotFieldsColumnsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFieldsDataGetNames
; Description ...: Retrieve an array of Field Names set as Data Fields.
; Syntax ........: _LOCalc_RangePivotFieldsDataGetNames(ByRef $oPivotTable)
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of Pivot Table Fields.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Pivot Table Field Name.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Pivot Table Field Names currently set as Data Fields, contained in the Pivot Table. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFieldsDataGetNames(ByRef $oPivotTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]
	Local $iCount

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oPivotTable.DataFields.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asNames[$iCount]

	For $i = 0 To $iCount - 1
		$asNames[$i] = $oPivotTable.DataFields.getByIndex($i).Name()
		If Not IsString($asNames[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asNames)
EndFunc   ;==>_LOCalc_RangePivotFieldsDataGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFieldSettings
; Description ...: Set or Retrieve Pivot Field settings.
; Syntax ........: _LOCalc_RangePivotFieldSettings(ByRef $oPivotField[, $iFieldType = Null[, $iFunc = Null[, $bShowEmpty = Null[, $iDisplayType = Null[, $sBaseField = Null[, $iBaseItem = Null[, $sBaseItem = Null]]]]]]])
; Parameters ....: $oPivotField         - [in/out] an object. A Pivot Table Field object returned by a previous _LOCalc_RangePivotFieldGetObjByName function.
;                  $iFieldType          - [optional] an integer value (0-4). Default is Null. The type of the Field, or field layout, either a Column, Row, Filter or Data Field or not used at all. See Constants $LOC_PIVOT_TBL_FIELD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iFunc               - [optional] an integer value (0-12). Default is Null. The Function used by the field to calculate the subtotal. See Constants $LOC_COMPUTE_FUNC_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bShowEmpty          - [optional] a boolean value. Default is Null. If True, empty Columns and Rows will be included in the results table.
;                  $iDisplayType        - [optional] an integer value (0-8). Default is Null. The type of calculation to be done to the results. See Constants $LOC_PIVOT_TBL_FIELD_DISP_* as defined in LibreOfficeCalc_Constants.au3.
;                  $sBaseField          - [optional] a string value. Default is Null. The Field to base the calculation upon.
;                  $iBaseItem           - [optional] an integer value (0-2). Default is Null. The type of Base Item to base the calculation on. See remarks. See Constants $LOC_PIVOT_TBL_FIELD_BASE_ITEM_* as defined in LibreOfficeCalc_Constants.au3.
;                  $sBaseItem           - [optional] a string value. Default is Null. The base item's name to base the calculation on, if $iBaseItem is set to $LOC_PIVOT_TBL_FIELD_BASE_ITEM_NAMED.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotField not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFieldType not an Integer, less than 0 or greater than 4. See Constants $LOC_PIVOT_TBL_FIELD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iFunc not an Integer, less than 0 or greater than 12. See Constants $LOC_COMPUTE_FUNC_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $bShowEmpty not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iDisplayType not an Integer, less than 0 or greater than 8. See Constants $LOC_PIVOT_TBL_FIELD_DISP_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $sBaseField not a String.
;                  @Error 1 @Extended 7 Return 0 = $iBaseItem not an Integer, less than 0 or greater than 2. See Constants $LOC_PIVOT_TBL_FIELD_BASE_ITEM_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iBaseItem set to $LOC_PIVOT_TBL_FIELD_BASE_ITEM_NAMED, and $sBaseItem is not called and no previous value is set.
;                  @Error 1 @Extended 9 Return 0 = $sBaseItem not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.sheet.DataPilotFieldReference" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Reference Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iFieldType
;                  |                               2 = Error setting $iFunc
;                  |                               4 = Error setting $bShowEmpty
;                  |                               8 = Error setting $iDisplayType
;                  |                               16 = Error setting $sBaseField
;                  |                               32 = Error setting $iBaseItem
;                  |                               64 = Error setting $sBaseItem
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: It is the user's responsibility to ensure the a Base Item's name is correct, and exists, also field names etc.
;                  If $iBaseItem is set to $LOC_PIVOT_TBL_FIELD_BASE_ITEM_NAMED, you must fill in $sBaseItem also.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFieldSettings(ByRef $oPivotField, $iFieldType = Null, $iFunc = Null, $bShowEmpty = Null, $iDisplayType = Null, $sBaseField = Null, $iBaseItem = Null, $sBaseItem = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avPivotField[7]
	Local $tReference
	Local $iError = 0

	If Not IsObj($oPivotField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iFieldType, $iFunc, $bShowEmpty, $iDisplayType, $sBaseField, $iBaseItem, $sBaseItem) Then
		If IsObj($oPivotField.Reference()) Then
			__LO_ArrayFill($avPivotField, $oPivotField.Orientation(), $oPivotField.Function(), $oPivotField.ShowEmpty(), $oPivotField.Reference.ReferenceType(), _
					$oPivotField.Reference.ReferenceField(), $oPivotField.Reference.ReferenceItemType(), $oPivotField.Reference.ReferenceItemName())

		Else
			__LO_ArrayFill($avPivotField, $oPivotField.Orientation(), $oPivotField.Function(), $oPivotField.ShowEmpty(), $LOC_PIVOT_TBL_FIELD_DISP_NONE, "", _
					$LOC_PIVOT_TBL_FIELD_BASE_ITEM_NAMED, "")
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPivotField)
	EndIf

	If ($iFieldType <> Null) Then
		If Not __LO_IntIsBetween($iFieldType, $LOC_PIVOT_TBL_FIELD_TYPE_HIDDEN, $LOC_PIVOT_TBL_FIELD_TYPE_DATA) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oPivotField.Orientation = $iFieldType
		$iError = ($oPivotField.Orientation() = $iFieldType) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iFunc <> Null) Then
		If Not __LO_IntIsBetween($iFunc, $LOC_COMPUTE_FUNC_NONE, $LOC_COMPUTE_FUNC_VARP) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPivotField.Function = $iFunc
		$iError = ($oPivotField.Function() = $iFunc) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bShowEmpty <> Null) Then
		If Not IsBool($bShowEmpty) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPivotField.ShowEmpty = $bShowEmpty
		$iError = ($oPivotField.ShowEmpty() = $bShowEmpty) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($oPivotField.Orientation() = $LOC_PIVOT_TBL_FIELD_TYPE_DATA) Then
		If IsObj($oPivotField.Reference()) Then
			$tReference = $oPivotField.Reference()
			If Not IsObj($tReference) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Else
			$tReference = __LO_CreateStruct("com.sun.star.sheet.DataPilotFieldReference")
			If Not IsObj($tReference) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		EndIf

		If ($iDisplayType <> Null) Then
			If Not __LO_IntIsBetween($iDisplayType, $LOC_PIVOT_TBL_FIELD_DISP_NONE, $LOC_PIVOT_TBL_FIELD_DISP_INDEX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$tReference.ReferenceType = $iDisplayType
		EndIf

		If ($sBaseField <> Null) Then
			If Not IsString($sBaseField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$tReference.ReferenceField = $sBaseField
		EndIf

		If ($iBaseItem <> Null) Then
			If Not __LO_IntIsBetween($iBaseItem, $LOC_PIVOT_TBL_FIELD_BASE_ITEM_NAMED, $LOC_PIVOT_TBL_FIELD_BASE_ITEM_NEXT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
			If ($iBaseItem = $LOC_PIVOT_TBL_FIELD_BASE_ITEM_NAMED) And ($sBaseItem = Null) And ($tReference.ReferenceItemName() = "") Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$tReference.ReferenceItemType = $iBaseItem
		EndIf

		If ($sBaseItem <> Null) Then
			If Not IsString($sBaseItem) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$tReference.ReferenceItemName = $sBaseItem
		EndIf

		$oPivotField.Reference = $tReference

		$iError = (__LO_VarsAreNull($iDisplayType)) ? ($iError) : (($oPivotField.Reference.ReferenceType() = $iDisplayType) ? ($iError) : (BitOR($iError, 8)))
		$iError = (__LO_VarsAreNull($sBaseField)) ? ($iError) : (($oPivotField.Reference.ReferenceField() = $sBaseField) ? ($iError) : (BitOR($iError, 16)))
		$iError = (__LO_VarsAreNull($iBaseItem)) ? ($iError) : (($oPivotField.Reference.ReferenceItemType() = $iBaseItem) ? ($iError) : (BitOR($iError, 32)))
		$iError = (__LO_VarsAreNull($sBaseItem)) ? ($iError) : (($oPivotField.Reference.ReferenceItemName() = $sBaseItem) ? ($iError) : (BitOR($iError, 64)))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangePivotFieldSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFieldsFiltersGetNames
; Description ...: Retrieve an array of Field Names set as Filter Fields.
; Syntax ........: _LOCalc_RangePivotFieldsFiltersGetNames(ByRef $oPivotTable)
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of Pivot Table Fields.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Pivot Table Field Name.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Pivot Table Field Names currently set as Filter Fields, contained in the Pivot Table. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFieldsFiltersGetNames(ByRef $oPivotTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]
	Local $iCount

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oPivotTable.PageFields.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asNames[$iCount]

	For $i = 0 To $iCount - 1
		$asNames[$i] = $oPivotTable.PageFields.getByIndex($i).Name()
		If Not IsString($asNames[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asNames)
EndFunc   ;==>_LOCalc_RangePivotFieldsFiltersGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFieldsGetNames
; Description ...: Retrieve an array of Fields available in the Pivot Table Source.
; Syntax ........: _LOCalc_RangePivotFieldsGetNames(ByRef $oPivotTable)
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of Pivot Table Fields.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Pivot Table Field Name.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Pivot Table Field Names contained in the Pivot Table. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: There is always a "Data" field present.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFieldsGetNames(ByRef $oPivotTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]
	Local $iCount

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oPivotTable.DataPilotFields.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asNames[$iCount]

	For $i = 0 To $iCount - 1
		$asNames[$i] = $oPivotTable.DataPilotFields.getByIndex($i).Name()
		If Not IsString($asNames[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asNames)
EndFunc   ;==>_LOCalc_RangePivotFieldsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFieldsRowsGetNames
; Description ...: Retrieve an array of Field Names set as Row Fields.
; Syntax ........: _LOCalc_RangePivotFieldsRowsGetNames(ByRef $oPivotTable)
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of Pivot Table Fields.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Pivot Table Field Name.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Pivot Table Field Names currently set as Row Fields, contained in the Pivot Table. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFieldsRowsGetNames(ByRef $oPivotTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]
	Local $iCount

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oPivotTable.RowFields.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asNames[$iCount]

	For $i = 0 To $iCount - 1
		$asNames[$i] = $oPivotTable.RowFields.getByIndex($i).Name()
		If Not IsString($asNames[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asNames)
EndFunc   ;==>_LOCalc_RangePivotFieldsRowsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFieldsUnusedGetNames
; Description ...: Retrieve an array of Field Names not current used in any of the Fields.
; Syntax ........: _LOCalc_RangePivotFieldsUnusedGetNames(ByRef $oPivotTable)
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of Pivot Table Fields.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Pivot Table Field Name.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Pivot Table Field Names currently not used in any field types, contained in the Pivot Table. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: There is always a "Data" field present.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFieldsUnusedGetNames(ByRef $oPivotTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]
	Local $iCount

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oPivotTable.HiddenFields.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asNames[$iCount]

	For $i = 0 To $iCount - 1
		$asNames[$i] = $oPivotTable.HiddenFields.getByIndex($i).Name()
		If Not IsString($asNames[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asNames)
EndFunc   ;==>_LOCalc_RangePivotFieldsUnusedGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFilter
; Description ...: Apply a Filter to a Pivot Table.
; Syntax ........: _LOCalc_RangePivotFilter(ByRef $oPivotTable[, $atFilterField = Null[, $bCaseSensitive = Null[, $bSkipDupl = Null[, $bUseRegExp = Null]]]])
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
;                  $atFilterField       - [optional] an array of dll structs. Default is Null. A single column Array of Filter Fields previously created by _LOCalc_FilterFieldCreate. Maximum of 3 Fields allowed.
;                  $bCaseSensitive      - [optional] a boolean value. Default is Null. If True, the Filtering operation will be case sensitive.
;                  $bSkipDupl           - [optional] a boolean value. Default is Null. If True, Duplicate values will be skipped in the list of filtered data.
;                  $bUseRegExp          - [optional] a boolean value. Default is Null. If True, the String Value set will be considered as using Regular expressions.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $atFilterField not an array or has more than 3 elements.
;                  @Error 1 @Extended 3 Return ? = $atFilterField contains an element that is not an Object. Returning the element number containing the error.
;                  @Error 1 @Extended 4 Return 0 = $bCaseSensitive not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bSkipDupl not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bUseRegExp not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $atFilterField
;                  |                               2 = Error setting $bCaseSensitive
;                  |                               4 = Error setting $bSkipDupl
;                  |                               8 = Error setting $bUseRegExp
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_RangePivotFilterClear, _LOCalc_FilterFieldCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFilter(ByRef $oPivotTable, $atFilterField = Null, $bCaseSensitive = Null, $bSkipDupl = Null, $bUseRegExp = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFilter[4]

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($atFilterField, $bCaseSensitive, $bSkipDupl, $bUseRegExp) Then
		__LO_ArrayFill($avFilter, $oPivotTable.FilterDescriptor.getFilterFields2(), $oPivotTable.FilterDescriptor.IsCaseSensitive(), _
				$oPivotTable.FilterDescriptor.SkipDuplicates(), $oPivotTable.FilterDescriptor.UseRegularExpressions())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFilter)
	EndIf

	If ($atFilterField <> Null) Then
		If Not IsArray($atFilterField) Or (UBound($atFilterField) > 3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To UBound($atFilterField) - 1
			If Not IsObj($atFilterField[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, $i)
		Next

		$oPivotTable.FilterDescriptor.setFilterFields2($atFilterField)
		$iError = (UBound($oPivotTable.FilterDescriptor.getFilterFields2()) = UBound($atFilterField)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bCaseSensitive <> Null) Then
		If Not IsBool($bCaseSensitive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPivotTable.FilterDescriptor.IsCaseSensitive = $bCaseSensitive
		$iError = ($oPivotTable.FilterDescriptor.IsCaseSensitive() = $bCaseSensitive) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bSkipDupl <> Null) Then
		If Not IsBool($bSkipDupl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPivotTable.FilterDescriptor.SkipDuplicates = $bSkipDupl
		$iError = ($oPivotTable.FilterDescriptor.SkipDuplicates() = $bSkipDupl) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bUseRegExp <> Null) Then
		If Not IsBool($bUseRegExp) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPivotTable.FilterDescriptor.UseRegularExpressions = $bUseRegExp
		$iError = ($oPivotTable.FilterDescriptor.UseRegularExpressions() = $bUseRegExp) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangePivotFilter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotFilterClear
; Description ...: Clear any previous filters for a Pivot Table.
; Syntax ........: _LOCalc_RangePivotFilterClear(ByRef $oPivotTable)
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to clear previous filter.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared the Pivot Table Filter.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangePivotFilter
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotFilterClear(ByRef $oPivotTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aEmpty[0]

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPivotTable.FilterDescriptor.setFilterFields($aEmpty)
	$oPivotTable.FilterDescriptor.setFilterFields2($aEmpty)
	$oPivotTable.FilterDescriptor.setFilterFields3($aEmpty)

	If Not (UBound($oPivotTable.FilterDescriptor.getFilterFields()) = 0) Or _
			Not (UBound($oPivotTable.FilterDescriptor.getFilterFields2()) = 0) Or _
			Not (UBound($oPivotTable.FilterDescriptor.getFilterFields3()) = 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangePivotFilterClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotGetObjByIndex
; Description ...: Retrieve the Object for a Pivot table by Index.
; Syntax ........: _LOCalc_RangePivotGetObjByIndex(ByRef $oSheet, $iIndex)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iIndex              - an integer value. The Index number of the Pivot Table to retrieve the Object for. 0 Based.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iIndex not an Integer, less than 0 or greater than number of Pivot Tables contained in Sheet.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Pivot Table Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Pivot Table's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotGetObjByIndex(ByRef $oSheet, $iIndex)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPivotTable

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iIndex, 0, $oSheet.DataPilotTables.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oPivotTable = $oSheet.DataPilotTables.getByIndex($iIndex)
	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPivotTable)
EndFunc   ;==>_LOCalc_RangePivotGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotGetObjByName
; Description ...: Retrieve the object for a Pivot Table by name.
; Syntax ........: _LOCalc_RangePivotGetObjByName(ByRef $oSheet, $sName)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sName               - a string value. The name of the Pivot Table to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Sheet called in $oSheet does not contain a Pivot Table with name called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Pivot Table Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Pivot Table's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotGetObjByName(ByRef $oSheet, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPivotTable

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSheet.DataPilotTables.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPivotTable = $oSheet.DataPilotTables.getByName($sName)
	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPivotTable)
EndFunc   ;==>_LOCalc_RangePivotGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotInsert
; Description ...: Insert a new Pivot Table.
; Syntax ........: _LOCalc_RangePivotInsert(ByRef $oSourceRange, ByRef $oDestRange[, $sName = ""[, $sField = ""[, $iFieldType = $LOC_PIVOT_TBL_FIELD_TYPE_COLUMN[, $iFunc = $LOC_COMPUTE_FUNC_NONE]]]])
; Parameters ....: $oSourceRange        - [in/out] an object. The Range containing the Data to use in the Pivot Table. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oDestRange          - [in/out] an object. The Range to output the Pivot Table to. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sName               - [optional] a string value. Default is "". The name of the new Pivot Table. If blank, an automatic name is generated.
;                  $sField              - [optional] a string value. Default is "". The name of one of the available fields in the source range to use in the Table. See remarks.
;                  $iFieldType          - [optional] an integer value (0-4). Default is $LOC_PIVOT_TBL_FIELD_TYPE_COLUMN. The type to set the field called in $sField to. See Constants $LOC_PIVOT_TBL_FIELD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iFunc               - [optional] an integer value (0-12). Default is $LOC_COMPUTE_FUNC_NONE. The function to set for the Field. See Constants $LOC_COMPUTE_FUNC_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSourceRange not an object.
;                  @Error 1 @Extended 2 Return 0 = $oDestRange not an object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sField not a String.
;                  @Error 1 @Extended 5 Return 0 = $iFieldType not an Integer, less than 0 or greater than 4. See Constants $LOC_PIVOT_TBL_FIELD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iFunc not an Integer, less than 0 or greater than 12. See Constants $LOC_COMPUTE_FUNC_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = Pivot Table name called in $sName or automatically generated name, already exists in Sheet.
;                  @Error 1 @Extended 8 Return 0 = Range called in $oDestRange is within the source range.
;                  @Error 1 @Extended 9 Return 0 = Field name called in $sField not found in available fields for Pivot Table.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.table.CellAddress Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a Data Pilot Descriptor Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Destination Address.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Source Address.
;                  @Error 3 @Extended 3 Return 0 = Failed to insert Pivot Table.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning new Pivot Table's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you do not call a field in $sField, the resulting Pivot Table will display "Empty", and will need a field set either manually or using one of the other functions before it will appear normal.
;                  Any existing data within the Destination range will be overwritten.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotInsert(ByRef $oSourceRange, ByRef $oDestRange, $sName = "", $sField = "", $iFieldType = $LOC_PIVOT_TBL_FIELD_TYPE_COLUMN, $iFunc = $LOC_COMPUTE_FUNC_NONE)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPivotDesc, $oPivotTable
	Local $tCellAddr, $tSourceAddr, $tDestAddr
	Local $iCount = 1

	If Not IsObj($oSourceRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDestRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iFieldType, $LOC_PIVOT_TBL_FIELD_TYPE_HIDDEN, $LOC_PIVOT_TBL_FIELD_TYPE_DATA) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not __LO_IntIsBetween($iFunc, $LOC_COMPUTE_FUNC_NONE, $LOC_COMPUTE_FUNC_VARP) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If ($sName = "") Then
		While $oDestRange.Spreadsheet.DataPilotTables.hasByName("DataPilot" & $iCount)
			$iCount += 1

			Sleep((IsInt($iCount / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
		WEnd

		$sName = "DataPilot" & $iCount
	EndIf

	If $oDestRange.Spreadsheet.DataPilotTables.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$tDestAddr = $oDestRange.RangeAddress()
	If Not IsObj($tDestAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tCellAddr = __LO_CreateStruct("com.sun.star.table.CellAddress")
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tCellAddr.Sheet = $tDestAddr.Sheet()
	$tCellAddr.Column = $tDestAddr.StartColumn()
	$tCellAddr.Row = $tDestAddr.StartRow()

	$tSourceAddr = $oSourceRange.RangeAddress()
	If Not IsObj($tSourceAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If ($tSourceAddr.Sheet() = $tDestAddr.Sheet()) And _
			__LO_IntIsBetween($tDestAddr.StartColumn(), $tSourceAddr.StartColumn(), $tSourceAddr.EndColumn()) And _
			__LO_IntIsBetween($tDestAddr.StartRow(), $tSourceAddr.StartRow(), $tSourceAddr.EndRow()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	$oPivotDesc = $oDestRange.Spreadsheet.DataPilotTables.createDataPilotDescriptor()
	If Not IsObj($oPivotDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oPivotDesc.SourceRange = $tSourceAddr

	If ($sField <> "") Then
		If Not $oPivotDesc.DataPilotFields.hasByName($sField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oPivotDesc.DataPilotFields.getByName($sField).Orientation = $iFieldType

		If ($iFieldType = $LOC_PIVOT_TBL_FIELD_TYPE_DATA) And ($iFunc = $LOC_COMPUTE_FUNC_NONE) Then
			$oPivotDesc.DataPilotFields.getByName($sField).Function = $LOC_COMPUTE_FUNC_SUM

		Else
			$oPivotDesc.DataPilotFields.getByName($sField).Function = $iFunc
		EndIf
	EndIf

	$oPivotDesc.DrillDownOnDoubleClick = False ; These are set to True when creating the Descriptor, but are normally false on a new Pivot Table.
	$oPivotDesc.ShowFilterButton = False ; These are set to True when creating the Descriptor, but are normally false on a new Pivot Table.

	$oDestRange.Spreadsheet.DataPilotTables.insertNewByName($sName, $tCellAddr, $oPivotDesc)
	If Not $oDestRange.Spreadsheet.DataPilotTables.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oPivotTable = $oDestRange.Spreadsheet.DataPilotTables.getByName($sName)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPivotTable)
EndFunc   ;==>_LOCalc_RangePivotInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotName
; Description ...: Set or Retrieve the Pivot Table Name.
; Syntax ........: _LOCalc_RangePivotName(ByRef $oDoc, ByRef $oPivotTable[, $sName = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
;                  $sName               - [optional] a string value. Default is Null. The new name of the Pivot Table.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPivotTable not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = Document called in $oDoc does not contain the Pivot Table called in $oPivotTable.
;                  @Error 1 @Extended 5 Return 0 = Parent sheet of Pivot Table called in $oPivotTable already contains a Pivot Table with the name called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Pivot Table Parent Sheet.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Name was successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning Pivot Table's current Name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotName(ByRef $oDoc, ByRef $oPivotTable, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sName) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oPivotTable.Name())

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSheet = $oDoc.Sheets.getByIndex($oPivotTable.OutputRange.Sheet())
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oSheet.DataPilotTables.hasByName($oPivotTable.Name()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $oSheet.DataPilotTables.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oPivotTable.Name = $sName

	If Not ($oPivotTable.Name() = $sName) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangePivotName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotRefresh
; Description ...: Refresh the Pivot Table.
; Syntax ........: _LOCalc_RangePivotRefresh(ByRef $oPivotTable)
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Pivot Table was successfully refreshed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Refreshing a table re-creates it from the present source data.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotRefresh(ByRef $oPivotTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPivotTable.refresh()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangePivotRefresh

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotSettings
; Description ...: Set or Retrieve Pivot Table settings.
; Syntax ........: _LOCalc_RangePivotSettings(ByRef $oPivotTable[, $bIgnoreEmpty = Null[, $bIdentifyCat = Null[, $bTotalCol = Null[, $bTotalRow = Null[, $bAddFilter = Null[, $bEnableDrill = Null]]]]]])
; Parameters ....: $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
;                  $bIgnoreEmpty        - [optional] a boolean value. Default is Null. If True, empty fields in the source are ignored.
;                  $bIdentifyCat        - [optional] a boolean value. Default is Null. If True, Rows without labels are automatically assigned a label.
;                  $bTotalCol           - [optional] a boolean value. Default is Null. If True, a Total Column is present.
;                  $bTotalRow           - [optional] a boolean value. Default is Null. If True, a Total Row is present.
;                  $bAddFilter          - [optional] a boolean value. Default is Null. If True, a filter button is added based on spreadsheet data.
;                  $bEnableDrill        - [optional] a boolean value. Default is Null. If True, double-clicking on a item label will show or hide details for the item.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPivotTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bIgnoreEmpty not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bIdentifyCat not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bTotalCol not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bTotalRow not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bAddFilter not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bEnableDrill not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bIgnoreEmpty
;                  |                               2 = Error setting $bIdentifyCat
;                  |                               4 = Error setting $bTotalCol
;                  |                               8 = Error setting $bTotalRow
;                  |                               16 = Error setting $bAddFilter
;                  |                               32 = Error setting $bEnableDrill
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I was unable to find a setting for "Show Expand/Collapse buttons", therefore it is not settable currently.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotSettings(ByRef $oPivotTable, $bIgnoreEmpty = Null, $bIdentifyCat = Null, $bTotalCol = Null, $bTotalRow = Null, $bAddFilter = Null, $bEnableDrill = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $abPivotSettings[6]
	Local $iError = 0

	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bIgnoreEmpty, $bIdentifyCat, $bTotalCol, $bTotalRow, $bAddFilter, $bEnableDrill) Then
		__LO_ArrayFill($abPivotSettings, $oPivotTable.IgnoreEmptyRows(), $oPivotTable.RepeatIfEmpty(), $oPivotTable.ColumnGrand(), $oPivotTable.RowGrand(), _
				$oPivotTable.ShowFilterButton(), $oPivotTable.DrillDownOnDoubleClick())

		Return SetError($__LO_STATUS_SUCCESS, 1, $abPivotSettings)
	EndIf

	If ($bIgnoreEmpty <> Null) Then
		If Not IsBool($bIgnoreEmpty) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oPivotTable.IgnoreEmptyRows = $bIgnoreEmpty
		$iError = ($oPivotTable.IgnoreEmptyRows() = $bIgnoreEmpty) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bIdentifyCat <> Null) Then
		If Not IsBool($bIdentifyCat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPivotTable.RepeatIfEmpty = $bIdentifyCat
		$iError = ($oPivotTable.RepeatIfEmpty() = $bIdentifyCat) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bTotalCol <> Null) Then
		If Not IsBool($bTotalCol) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPivotTable.ColumnGrand = $bTotalCol
		$iError = ($oPivotTable.ColumnGrand() = $bTotalCol) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bTotalRow <> Null) Then
		If Not IsBool($bTotalRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPivotTable.RowGrand = $bTotalRow
		$iError = ($oPivotTable.RowGrand() = $bTotalRow) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bAddFilter <> Null) Then
		If Not IsBool($bAddFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPivotTable.ShowFilterButton = $bAddFilter
		$iError = ($oPivotTable.ShowFilterButton() = $bAddFilter) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bEnableDrill <> Null) Then
		If Not IsBool($bEnableDrill) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPivotTable.DrillDownOnDoubleClick = $bEnableDrill
		$iError = ($oPivotTable.DrillDownOnDoubleClick() = $bEnableDrill) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangePivotSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotsGetCount
; Description ...: Retrieve a count of Pivot tables contained in the Sheet.
; Syntax ........: _LOCalc_RangePivotsGetCount(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of Pivot Tables.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning a Count of Pivot tables contained in the Sheet.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotsGetCount(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oSheet.DataPilotTables.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOCalc_RangePivotsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotsGetNames
; Description ...: Retrieve an array of Pivot Tables contained in the Sheet.
; Syntax ........: _LOCalc_RangePivotsGetNames(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of Pivot Tables.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Pivot Table Name.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Pivot Table Names contained in the Sheet. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotsGetNames(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]
	Local $iCount

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oSheet.DataPilotTables.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asNames[$iCount]

	For $i = 0 To $iCount - 1
		$asNames[$i] = $oSheet.DataPilotTables.getByIndex($i).Name()
		If Not IsString($asNames[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asNames)
EndFunc   ;==>_LOCalc_RangePivotsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangePivotSource
; Description ...: Set or Retrieve the Pivot Table Source Range.
; Syntax ........: _LOCalc_RangePivotSource(ByRef $oDoc, ByRef $oPivotTable[, $oSourceRange = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oPivotTable         - [in/out] an object. A Pivot Table object returned by a previous _LOCalc_RangePivotInsert, _LOCalc_RangePivotGetObjByName or _LOCalc_RangePivotGetObjByIndex function.
;                  $oSourceRange        - [optional] an object. Default is Null. The Range containing the Data to use in the Pivot Table. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPivotTable not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSourceRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Source Range Parent Sheet.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Source Range Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oSourceRange
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Source Range was successfully set.
;                  @Error 0 @Extended 1 Return Object = Success. All optional parameters were called with Null, returning current source Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangePivotSource(ByRef $oDoc, ByRef $oPivotTable, $oSourceRange = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet, $oRange
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPivotTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($oSourceRange) Then
		$oSheet = $oDoc.Sheets.getByIndex($oPivotTable.SourceRange.Sheet())
		If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oRange = $oSheet.getCellRangeByPosition($oPivotTable.SourceRange.StartColumn(), $oPivotTable.SourceRange.StartRow(), $oPivotTable.SourceRange.EndColumn(), _
				$oPivotTable.SourceRange.EndRow())
		If Not IsObj($oRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $oRange)
	EndIf

	If Not IsObj($oSourceRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPivotTable.SourceRange = $oSourceRange.RangeAddress()

	$iError = (__LOCalc_RangeAddressIsSame($oPivotTable.SourceRange(), $oSourceRange.RangeAddress())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangePivotSource

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryColumnDiff
; Description ...: Query a Cell Range for differences on each column based on a specific row.
; Syntax ........: _LOCalc_RangeQueryColumnDiff(ByRef $oRange, $oCellToCompare)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range to look for differences in. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCellToCompare      - an object. A single Cell object (not a range) returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function. The Row this cell is located in will be used for the query.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCellToCompare not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oCellToCompare is not a single cell, cell ranges are not supported.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Address Struct from $oCellToCompare.
;                  @Error 3 @Extended 2 Return 0 = Failed to query column differences.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve query result cell addresses.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve cell range Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Looks for differences per column in the range, comparing the column to the value in the row $oCellToCompare is located. OOME 4.1. pg 488/489
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryColumnDiff(ByRef $oRange, $oCellToCompare)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $tCellAddr
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCellToCompare) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not ($oCellToCompare.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oCellToCompare.CellAddress()
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oRanges = $oRange.queryColumnDifferences($tCellAddr)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryColumnDiff

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryContents
; Description ...: Query a Cell or Cell range for specific cell contents.
; Syntax ........: _LOCalc_RangeQueryContents(ByRef $oRange, $iFlags)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFlags              - an integer value (1-1023). The Cell content type flag. Can be BitOR'd together. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFlags not an Integer, less than 1 or greater than 1023. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query cell content.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Empty cells in the range may be skipped depending on the flag used. For instance, when querying for styles, the returned ranges may not include empty cells even if styles are applied to those cells.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryContents(ByRef $oRange, $iFlags)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iFlags, $LOC_CELL_FLAG_VALUE, $LOC_CELL_FLAG_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRanges = $oRange.queryContentCells($iFlags)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryContents

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryDependents
; Description ...: Query a Cell or Cell Range for Dependents.
; Syntax ........: _LOCalc_RangeQueryDependents(ByRef $oRange[, $bRecursive = False])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bRecursive          - [optional] a boolean value. Default is False. If True, the query is repeated for each found cell.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bRecursive not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query cell dependents.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Dependent cells are cells which reference cells in this range. If $bRecursive is True, repeats query with all found cells (finds dependents of dependents, and so on).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryDependents(ByRef $oRange, $bRecursive = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bRecursive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRanges = $oRange.queryDependents($bRecursive)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryDependents

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryEmpty
; Description ...: Query a Cell or Cell Range for empty cells.
; Syntax ........: _LOCalc_RangeQueryEmpty(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query for empty cells.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryEmpty(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRanges = $oRange.queryEmptyCells()
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryEmpty

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryFormula
; Description ...: Query a Cell or Cell Range for formulas having a specific result.
; Syntax ........: _LOCalc_RangeQueryFormula(ByRef $oRange, $iResultType)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iResultType         - an integer value (1-7). The Formula result type. Can be BitOR'd together. See Constants $LOC_FORMULA_RESULT_TYPE_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iResultType not an Integer, less than 1 or greater than 7. See Constants $LOC_FORMULA_RESULT_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query cell formula results.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryFormula(ByRef $oRange, $iResultType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iResultType, $LOC_FORMULA_RESULT_TYPE_VALUE, $LOC_FORMULA_RESULT_TYPE_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRanges = $oRange.queryFormulaCells($iResultType)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryIntersection
; Description ...: Retrieve an array of cell ranges that intersect with a certain cell range.
; Syntax ........: _LOCalc_RangeQueryIntersection(ByRef $oRange, $oCell)
; Parameters ....: $oRange              - [in/out] an object. A Cell range that contains the cell or cell range called in $oCell. A Cell Range object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCell               - an object. A Cell or Cell Range located inside of the cell range called in $oRange. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Range Address Struct from $oCell.
;                  @Error 3 @Extended 2 Return 0 = Failed to query cell range intersections.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve query result cell addresses.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve cell range Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryIntersection(ByRef $oRange, $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $tRangeAddr
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tRangeAddr = $oCell.RangeAddress()
	If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oRanges = $oRange.queryIntersection($tRangeAddr)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oCell.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryIntersection

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryPrecedents
; Description ...: Query a Cell or Cell Range for Precedents.
; Syntax ........: _LOCalc_RangeQueryPrecedents(ByRef $oRange[, $bRecursive = False])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bRecursive          - [optional] a boolean value. Default is False. If True, the query is repeated for each found cell.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bRecursive not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query cell precedents.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Precedent cells are cells which are referenced by cells in this range. If $bRecursive is True, repeats query with all found cells (finds precedents of precedents, and so on).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryPrecedents(ByRef $oRange, $bRecursive = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bRecursive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRanges = $oRange.queryPrecedents($bRecursive)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryPrecedents

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryRowDiff
; Description ...: Query a Cell Range for differences on each row based on a specific column.
; Syntax ........: _LOCalc_RangeQueryRowDiff(ByRef $oRange, $oCellToCompare)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range to look for differences in. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCellToCompare      - an object. A single Cell object (not a range) returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function. The Column this cell is located in will be used for the query.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCellToCompare not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oCellToCompare is not a single cell, cell ranges are not supported.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Address Struct from $oCellToCompare.
;                  @Error 3 @Extended 2 Return 0 = Failed to query row differences.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve query result cell addresses.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve cell range Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Looks for differences per row in the range, comparing the row to the value in the column $oCellToCompare is located. OOME 4.1. pg 488/489
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryRowDiff(ByRef $oRange, $oCellToCompare)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $tCellAddr
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCellToCompare) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not ($oCellToCompare.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oCellToCompare.CellAddress()
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oRanges = $oRange.queryRowDifferences($tCellAddr)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryRowDiff

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryVisible
; Description ...: Query a Cell or Cell Range for visible cells.
; Syntax ........: _LOCalc_RangeQueryVisible(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query for visible cell.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryVisible(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRanges = $oRange.queryVisibleCells()
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeReplace
; Description ...: Replace the first instances of a search within a Range.
; Syntax ........: _LOCalc_RangeReplace(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $sReplaceString)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $sReplaceString      - a string value. A String of text or a regular expression to replace the first result with.
; Return values .: Success: 0 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;                  @Error 1 @Extended 5 Return 0 = $sReplaceString not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Found a result, but failed to replace it.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Search and replace was successful, no results found.
;                  @Error 0 @Extended 1 Return Object = Success. Search and Replace was successful, returning Object for Cell that the find and replace was performed upon.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office does not offer a Function to call to replace only one result within a Range, consequently I have had to create my own, which means this may not work exactly as expected.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_RangeFindAll, _LOCalc_RangeFindNext, _LOCalc_RangeReplaceAll,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeReplace(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $sReplaceString)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iReplacements
	Local $oResult

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplaceString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSrchDescript.SearchString = $sSearchString
	$oSrchDescript.ReplaceString = $sReplaceString

	$oResult = $oRange.findFirst($oSrchDescript)
	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_SUCCESS, 0, 1) ; No Results

	$iReplacements = $oResult.replaceAll($oSrchDescript)

	Return ($iReplacements > 0) ? (SetError($__LO_STATUS_SUCCESS, 1, $oResult)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_RangeReplace

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeReplaceAll
; Description ...: Replace all instances of a search.
; Syntax ........: _LOCalc_RangeReplaceAll(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $sReplaceString)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a Regular Expression to Search for.
;                  $sReplaceString      - a string value. A String of text or a Regular Expression to replace any results with.
; Return values .: Success: 0 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;                  @Error 1 @Extended 5 Return 0 = $sReplaceString not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Results were found, but failed to perform the replacement.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Search was successful, no results found.
;                  @Error 0 @Extended ? Return Array = Success. Search and Replace was successful, @Extended set to number of replacements made, returning array Cell/CellRange Objects of all Cells modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only the Sheet that contains the Range is searched, to search all Sheets you will have to cycle through and perform a search for each.
;                  Number of Replacements DOESN'T mean that is the size of the Array. If replacements where in several cells connected, the return will be a Cell Range for that area instead of individual cells.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_RangeFindAll, _LOCalc_RangeFindNext, _LOCalc_RangeReplace
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeReplaceAll(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $sReplaceString)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResults
	Local $aoResults[0]
	Local $iReplacements

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplaceString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSrchDescript.SearchString = $sSearchString
	$oSrchDescript.ReplaceString = $sReplaceString

	$oResults = $oRange.findAll($oSrchDescript)
	If Not IsObj($oResults) Then Return SetError($__LO_STATUS_SUCCESS, 0, 1)

	If ($oResults.getCount() > 0) Then
		ReDim $aoResults[$oResults.getCount]
		For $i = 0 To $oResults.getCount() - 1
			$aoResults[$i] = $oResults.getByIndex($i)
			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	$iReplacements = $oRange.replaceAll($oSrchDescript)

	Return ($iReplacements > 0) ? (SetError($__LO_STATUS_SUCCESS, $iReplacements, $aoResults)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_RangeReplaceAll

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowDelete
; Description ...: Delete Rows from a Sheet.
; Syntax ........: _LOCalc_RangeRowDelete(ByRef $oRange, $iRow[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iRow                - an integer value. The Row to begin deleting at. The Row called will be deleted. See remarks.
;                  $iCount              - [optional] an integer value. Default is 1. The number of rows to delete, including the row called in $iRow.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iRow not an Integer, less than 0 or greater than number of Rows contained in the Range.
;                  @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested rows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Rows in L.O. Calc are 0 based, to Delete Row 1 in the LibreOffice UI, you would call $iRow with 0.
;                  Deleting Rows does not decrease the Row count, it simply erases the row's contents in a specific area and shifts all after content higher.
; Related .......: _LOCalc_RangeRowInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowDelete(ByRef $oRange, $iRow, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oRange.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iRow, 0, $oRows.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iCount, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRows.removeByIndex($iRow, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeRowDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowGetObjByPosition
; Description ...: Retrieve a Row's Object for further Row related functions.
; Syntax ........: _LOCalc_RangeRowGetObjByPosition(ByRef $oRange, $iRow)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iRow                - an integer value. The Row number to retrieve the Row Object for. See remarks.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iRow not an Integer, less than 0 or greater than number of Rows contained in the Range.
;                  @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Row Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Success, returning Row's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Rows in L.O. Calc are 0 based, to retrieve Row 1 in the LibreOffice UI, you would call $iRow with 0.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowGetObjByPosition(ByRef $oRange, $iRow)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows, $oRow

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oRange.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iRow, 0, $oRows.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRow = $oRows.getByIndex($iRow)
	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oRow)
EndFunc   ;==>_LOCalc_RangeRowGetObjByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowHeight
; Description ...: Set or Retrieve the Row's Height settings.
; Syntax ........: _LOCalc_RangeRowHeight(ByRef $oRow[, $bOptimal = Null[, $iHeight = Null]])
; Parameters ....: $oRow                - an object. A Row object returned by a previous _LOCalc_RangeRowGetObjByPosition function.
;                  $bOptimal            - [optional] a boolean value. Default is Null. If True, the Optimal height is automatically chosen.
;                  $iHeight             - [optional] an integer value (0-34464). Default is Null. The Height of the row, set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bOptimal not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer, less than 0 or greater than 34464.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bOptimal
;                  |                               2 = Error setting $iHeight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I am presently unable to find a setting for Optimal Height "Add" Value.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowHeight(ByRef $oRow, $bOptimal = Null, $iHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avHeight[2]
	Local $iError = 0

	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bOptimal, $iHeight) Then
		__LO_ArrayFill($avHeight, $oRow.OptimalHeight(), $oRow.Height())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avHeight)
	EndIf

	If ($bOptimal <> Null) Then
		If Not IsBool($bOptimal) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oRow.OptimalHeight = $bOptimal
		$iError = ($oRow.OptimalHeight() = $bOptimal) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iHeight <> Null) Then
		If Not __LO_IntIsBetween($iHeight, 0, 34464) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oRow.Height = $iHeight
		$iError = (__LO_IntIsBetween($oRow.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeRowHeight

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowInsert
; Description ...: Insert blank rows from a specific row in a Range.
; Syntax ........: _LOCalc_RangeRowInsert(ByRef $oRange, $iRow[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iRow                - an integer value. The Row to begin inserting blank rows at. See remarks. All contents from this row down will be shifted down.
;                  $iCount              - [optional] an integer value. Default is 1. The number of blank rows to insert.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iRow not an Integer, less than 0 or greater than number of Rows contained in the Range.
;                  @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully inserted blank rows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Rows in L.O. Calc are 0 based, to add Rows in Row 1 in the LibreOffice UI, you would call $iRow with 0.
;                  Inserting Rows does not increase the Row count, it simply adds blanks in a specific area and shifts all after content lower.
; Related .......: _LOCalc_RangeRowDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowInsert(ByRef $oRange, $iRow, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oRange.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iRow, 0, $oRows.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iCount, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRows.insertByIndex($iRow, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeRowInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowPageBreak
; Description ...: Set or retrieve current Page Break settings for a Row.
; Syntax ........: _LOCalc_RangeRowPageBreak(ByRef $oRow[, $bManualPageBreak = Null[, $bStartOfPageBreak = Null]])
; Parameters ....: $oRow                - [in/out] an object. A Row object returned by a previous _LOCalc_RangeRowGetObjByPosition function.
;                  $bManualPageBreak    - [optional] a boolean value. Default is Null. If True, this row is the beginning of a manual Page Break.
;                  $bStartOfPageBreak   - [optional] a boolean value. Default is Null. If True, this row is the beginning of a start of Page Break. See Remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRow not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bManualPageBreak not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bStartOfPageBreak not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bManualPageBreak
;                  |                               2 = Error setting $bStartOfPageBreak
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Setting $bStartOfPageBreak to True will insert a Manual Page Break, the same as setting $bManualPageBreak to True would.
;                  $bStartOfPageBreak setting is available more for indicating where Calc is inserting Page Breaks rather than for applying a setting. You can retrieve the settings for each row, and check if this value is True or not. If the Page break is an automatically inserted one, the value for $bManualPageBreak would be False.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowPageBreak(ByRef $oRow, $bManualPageBreak = Null, $bStartOfPageBreak = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abBreak[2]

	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bManualPageBreak, $bStartOfPageBreak) Then
		__LO_ArrayFill($abBreak, $oRow.IsManualPageBreak(), $oRow.IsStartOfNewPage())

		Return SetError($__LO_STATUS_SUCCESS, 1, $abBreak)
	EndIf

	If ($bManualPageBreak <> Null) Then
		If Not IsBool($bManualPageBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oRow.IsManualPageBreak = $bManualPageBreak
		$iError = ($oRow.IsManualPageBreak() = $bManualPageBreak) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bStartOfPageBreak <> Null) Then
		If Not IsBool($bStartOfPageBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oRow.IsStartOfNewPage = $bStartOfPageBreak
		$iError = ($oRow.IsStartOfNewPage() = $bStartOfPageBreak) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeRowPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowsGetCount
; Description ...: Retrieve the total count of Rows contained in a Range.
; Syntax ........: _LOCalc_RangeRowsGetCount(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning number of Rows contained in the Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: There is a fixed number of Rows per sheet, but different L.O. versions contain different amounts of Rows. This can also help determine how many rows are in a Cell Range.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowsGetCount(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oRange.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oRows.Count())
EndFunc   ;==>_LOCalc_RangeRowsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowVisible
; Description ...: Set or Retrieve the Row's visibility setting.
; Syntax ........: _LOCalc_RangeRowVisible(ByRef $oRow[, $bVisible = Null])
; Parameters ....: $oRow                - an object. A Row object returned by a previous _LOCalc_RangeRowGetObjByPosition function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the Row is Visible.
; Return values .: Success: 1 or Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRow not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were called with Null, returning Row's current visibility setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowVisible(ByRef $oRow, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bVisible) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oRow.IsVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRow.IsVisible = $bVisible
	$iError = ($oRow.IsVisible() = $bVisible) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeRowVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeSort
; Description ...: Sort a Range of Data.
; Syntax ........: _LOCalc_RangeSort(ByRef $oDoc, ByRef $oRange, ByRef $tSortField[, $bSortColumns = False[, $bHasHeader = False[, $bBindFormat = True[, $bCopyOutput = False[, $oCellOutput = Null[, $tSortField2 = Null[, $tSortField3 = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $tSortField          - [in/out] a dll struct value. A Sort Field Struct created by a previous _LOCalc_SortFieldCreate function.
;                  $bSortColumns        - [optional] a boolean value. Default is False. If True, Columns contained in the Cell Range are sorted Left to Right. If False, Rows contained in the Cell Range are sorted top to bottom.
;                  $bHasHeader          - [optional] a boolean value. Default is False. If True, the Row or Column has a header that will not be sorted.
;                  $bBindFormat         - [optional] a boolean value. Default is True. If True, formatting will be moved with the data sorted.
;                  $bCopyOutput         - [optional] a boolean value. Default is False. If True, the data remains unmodified and instead is copied to a Cell Range after sorting.
;                  $oCellOutput         - [optional] an object. Default is Null. If $bCopyOutput is True, this is the Cell range where the data is copied to. See Remarks.
;                  $tSortField2         - [optional] a dll struct value. Default is Null. Another Sort Field Struct created by a previous _LOCalc_SortFieldCreate function.
;                  $tSortField3         - [optional] a dll struct value. Default is Null. Another Sort Field Struct created by a previous _LOCalc_SortFieldCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $tSortField not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bSortColumns not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bHasHeader not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bBindFormat not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bCopyOutput not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $tSortField2 not an Object.
;                  @Error 1 @Extended 9 Return 0 = $tSortField3 not an Object.
;                  @Error 1 @Extended 10 Return 0 = $bCopyOutput called with True, but $oCellOutput not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Sort Descriptor.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.table.CellAddress" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Column called in $tSortField is greater than number of Columns contained in called Range.
;                  @Error 3 @Extended 2 Return 0 = Row called in $tSortField is greater than number of Rows contained in called Range.
;                  @Error 3 @Extended 3 Return 0 = Column called in $tSortField2 is greater than number of Columns contained in called Range.
;                  @Error 3 @Extended 4 Return 0 = Row called in $tSortField2 is greater than number of Rows contained in called Range.
;                  @Error 3 @Extended 5 Return 0 = Column called in $tSortField3 is greater than number of Columns contained in called Range.
;                  @Error 3 @Extended 6 Return 0 = Row called in $tSortField3 is greater than number of Rows contained in called Range.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve output cell Range Address.
;                  @Error 3 @Extended 8 Return 0 = Failed to retrieve the Standard Macro library object.
;                  @Error 3 @Extended 9 Return 0 = Failed to insert temporary Macro.
;                  @Error 3 @Extended 10 Return 0 = Failed to retrieve temporary Macro Object.
;                  @Error 3 @Extended 11 Return 0 = Failed to remove temporary Macro.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Sort was successfully processed for requested Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You can sort up to 3 Columns/Rows per Sort call by using $tSortField2 and $tSortField3.
;                  Only one Sort Field per Column/Row per sort, may be used, otherwise only the first Sort Field for that Column/Row is used.
;                  $oCellOutput indicates the cell to begin the output data, and does not need to be the same size as $oRange. Any data will be overwritten in order to output the copied Sort Data that is within range.
;                  Due to some form of bug in LibreOffice, the sort function does not work appropriately when using the normal method, so a slight workaround has been implemented, this workaround involves inserting a temporary Macro into the Document, calling that Macro, and then deleting the Macro once finished.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeSort(ByRef $oDoc, ByRef $oRange, ByRef $tSortField, $bSortColumns = False, $bHasHeader = False, $bBindFormat = True, $bCopyOutput = False, $oCellOutput = Null, $tSortField2 = Null, $tSortField3 = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSortDesc
	Local $atSortField[1]
	Local $aoParam[3]
	Local $aDummyArray[0]
	Local $tCellInputAddr, $tCellAddr
	Local $oStandardLibrary, $oScript
	Local $sMacro

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($tSortField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bSortColumns) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bHasHeader) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bBindFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$avSortDesc = $oRange.createSortDescriptor()
	If Not IsArray($avSortDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $bSortColumns Then
		If Not __LO_IntIsBetween($tSortField.Field(), 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else
		If Not __LO_IntIsBetween($tSortField.Field(), 0, $oRange.Rows.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndIf

	$atSortField[0] = $tSortField

	If ($tSortField2 <> Null) Then
		If Not IsObj($tSortField2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		If $bSortColumns Then
			If Not __LO_IntIsBetween($tSortField2.Field(), 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Else
			If Not __LO_IntIsBetween($tSortField2.Field(), 0, $oRange.Rows.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
		EndIf

		ReDim $atSortField[2]
		$atSortField[1] = $tSortField2
	EndIf

	If ($tSortField3 <> Null) Then
		If Not IsObj($tSortField3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		If $bSortColumns Then
			If Not __LO_IntIsBetween($tSortField3.Field(), 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		Else
			If Not __LO_IntIsBetween($tSortField3.Field(), 0, $oRange.Rows.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
		EndIf

		ReDim $atSortField[UBound($atSortField) + 1]
		$atSortField[UBound($atSortField) - 1] = $tSortField3
	EndIf

	If ($bCopyOutput = True) Then
		If Not IsObj($oCellOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tCellInputAddr = $oCellOutput.RangeAddress()
		If Not IsObj($tCellInputAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

		$tCellAddr = __LO_CreateStruct("com.sun.star.table.CellAddress")
		If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tCellAddr.Sheet = $tCellInputAddr.Sheet()
		$tCellAddr.Column = $tCellInputAddr.StartColumn()
		$tCellAddr.Row = $tCellInputAddr.StartRow()
	EndIf

	For $i = 0 To UBound($avSortDesc) - 1
		Switch $avSortDesc[$i].Name()
			Case "IsSortColumns"
				$avSortDesc[$i].Value = $bSortColumns

			Case "ContainsHeader"
				$avSortDesc[$i].Value = $bHasHeader

			Case "SortFields"
				$avSortDesc[$i].Value = $atSortField

			Case "BindFormatsToContent"
				$avSortDesc[$i].Value = $bBindFormat

			Case "CopyOutputData"
				$avSortDesc[$i].Value = $bCopyOutput

			Case "OutputPosition"
				If ($bCopyOutput = True) Then $avSortDesc[$i].Value = $tCellAddr
		EndSwitch
	Next

;~ $oRange.Sort($avSortDesc); This doesn't sort properly, thus a work around method is required.

	$sMacro = "REM Macro for Performing a Sort Function. Created By an AutoIt Script." & @CR & _ ; Just a description of the Macro
			"Sub AU3LibreOffice_Sort(oRange, avSortDesc, atField)" & @CR & _ ; Macro header, Parameters, oRange = Range to Sort, avSortDesc = The array of Sort Descriptor settings,  atField = Sort Descriptor Column/Row settings.
			@CR & _
			"For i = LBound(avSortDesc) To UBound(avSortDesc) " & @CR & _ ; Loop through passed array, re-applying Array of Sort Fields, seems necessary to make sort work.
			"If (avSortDesc(i).Name() = ""SortFields"") Then avSortDesc(i).Value = atField" & @CR & _
			@CR & _
			"Next " & @CR & _
			@CR & _
			"oRange.Sort(avSortDesc())" & @CR & _
			"End Sub" & @CR

	; Retrieving the BasicLibrary.Standard Object fails when using a newly opened document, I found a workaround by updating the
	; following setting.
	$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

	$oStandardLibrary = $oDoc.BasicLibraries.Standard()
	If Not IsObj($oStandardLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then $oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")

	$oStandardLibrary.insertByName("AU3LibreOffice_UDF_Macros", $sMacro)
	If Not $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 9, 0)

	$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.AU3LibreOffice_UDF_Macros.AU3LibreOffice_Sort?language=Basic&location=document")
	If Not IsObj($oScript) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 10, 0)

	$aoParam[0] = $oRange
	$aoParam[1] = $avSortDesc
	$aoParam[2] = $atSortField

	$oScript.Invoke($aoParam, $aDummyArray, $aDummyArray)

	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then $oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")
	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 11, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeSort

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeSortAlt
; Description ...: An alternate version of Sort Data function.
; Syntax ........: _LOCalc_RangeSortAlt(ByRef $oDoc, ByRef $oRange, ByRef $tSortField[, $bSortColumns = False[, $bHasHeader = False[, $bBindFormat = True[, $bNaturalOrder = True[, $bIncludeComments = False[, $bIncludeImages = False[, $tSortField2 = Null[, $tSortField3 = Null]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $tSortField          - [in/out] a dll struct value. A Sort Field Struct created by a previous _LOCalc_SortFieldCreate function.
;                  $bSortColumns        - [optional] a boolean value. Default is False. If True, Columns contained in the Cell Range are sorted Left to Right. If False, Rows contained in the Cell Range are sorted top to bottom.
;                  $bHasHeader          - [optional] a boolean value. Default is False. If True, the Row or Column has a header that will not be sorted.
;                  $bBindFormat         - [optional] a boolean value. Default is True. If True, formatting will be moved with the data sorted.
;                  $bNaturalOrder       - [optional] a boolean value. Default is True. If True, sort using natural order is enabled. See remarks.
;                  $bIncludeComments    - [optional] a boolean value. Default is False. If True, boundary columns or boundary rows containing comments are also sorted.
;                  $bIncludeImages      - [optional] a boolean value. Default is False. If True, boundary columns or boundary rows containing images are also sorted.
;                  $tSortField2         - [optional] a dll struct value. Default is Null. Another Sort Field Struct created by a previous _LOCalc_SortFieldCreate function.
;                  $tSortField3         - [optional] a dll struct value. Default is Null. Another Sort Field Struct created by a previous _LOCalc_SortFieldCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $tSortField not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bSortColumns not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bHasHeader not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bBindFormat not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bNaturalOrder not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bIncludeComments not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bIncludeImages not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $tSortField2 not an Object.
;                  @Error 1 @Extended 11 Return 0 = $tSortField3 not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "Col1" Property.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "Ascending1" Property.
;                  @Error 2 @Extended 3 Return 0 = Failed to create "CaseSensitive" Property.
;                  @Error 2 @Extended 4 Return 0 = Failed to create "ByRows" Property.
;                  @Error 2 @Extended 5 Return 0 = Failed to create "HasHeader" Property.
;                  @Error 2 @Extended 6 Return 0 = Failed to create "IncludeAttribs" Property.
;                  @Error 2 @Extended 7 Return 0 = Failed to create "NaturalSort" Property.
;                  @Error 2 @Extended 8 Return 0 = Failed to create "IncludeComments" Property.
;                  @Error 2 @Extended 9 Return 0 = Failed to create "IncludeImages" Property.
;                  @Error 2 @Extended 10 Return 0 = Failed to create "UserDefIndex" Property.
;                  @Error 2 @Extended 11 Return 0 = Failed to create "Col2" Property.
;                  @Error 2 @Extended 12 Return 0 = Failed to create "Ascending2" Property.
;                  @Error 2 @Extended 13 Return 0 = Failed to create "Col3" Property.
;                  @Error 2 @Extended 14 Return 0 = Failed to create "Ascending3" Property.
;                  @Error 2 @Extended 15 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 16 Return 0 = Failed to create instance of "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Column called in $tSortField is greater than number of Columns contained in called Range.
;                  @Error 3 @Extended 2 Return 0 = Row called in $tSortField is greater than number of Rows contained in called Range.
;                  @Error 3 @Extended 3 Return 0 = Column called in $tSortField2 is greater than number of Columns contained in called Range.
;                  @Error 3 @Extended 4 Return 0 = Row called in $tSortField2 is greater than number of Rows contained in called Range.
;                  @Error 3 @Extended 5 Return 0 = Column called in $tSortField3 is greater than number of Columns contained in called Range.
;                  @Error 3 @Extended 6 Return 0 = Row called in $tSortField3 is greater than number of Rows contained in called Range.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Sort was successfully processed for requested Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This version uses a UNO dispatch command as an alternative to the other sort function.
;                  Any selections by the user will be lost after calling this function, and the called range will be selected instead.
;                  The first Sort Field determines case sensitivity for the entire sort.
;                  You can sort up to 3 Columns/Rows per Sort call by using $tSortField2 and $tSortField3.
;                  Only one Sort Field per Column/Row per sort, may be used, otherwise only the first Sort Field for that Column/Row is used.
;                  Natural sort is a sort algorithm that sorts string-prefixed numbers based on the value of the numerical element in each sorted number, instead of the traditional way of sorting them as ordinary strings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeSortAlt(ByRef $oDoc, ByRef $oRange, ByRef $tSortField, $bSortColumns = False, $bHasHeader = False, $bBindFormat = True, $bNaturalOrder = True, $bIncludeComments = False, $bIncludeImages = False, $tSortField2 = Null, $tSortField3 = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDispatcher
	Local $iCount = 10
	Local $avParam[10]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($tSortField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bSortColumns) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bHasHeader) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bBindFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bNaturalOrder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsBool($bIncludeComments) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not IsBool($bIncludeImages) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	If $bSortColumns Then
		If Not __LO_IntIsBetween($tSortField.Field(), 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else
		If Not __LO_IntIsBetween($tSortField.Field(), 0, $oRange.Rows.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndIf

	$avParam[0] = __LO_SetPropertyValue("Col1", $tSortField.Field() + 1) ; UNO Execute seems to be 1 based.
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$avParam[1] = __LO_SetPropertyValue("Ascending1", $tSortField.IsAscending())
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$avParam[2] = __LO_SetPropertyValue("CaseSensitive", $tSortField.IsCaseSensitive())
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	$avParam[3] = __LO_SetPropertyValue("ByRows", ($bSortColumns) ? (False) : (True))
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$avParam[4] = __LO_SetPropertyValue("HasHeader", $bHasHeader)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)

	$avParam[5] = __LO_SetPropertyValue("IncludeAttribs", $bBindFormat)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 6, 0)

	$avParam[6] = __LO_SetPropertyValue("NaturalSort", $bNaturalOrder)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 7, 0)

	$avParam[7] = __LO_SetPropertyValue("IncludeComments", $bIncludeComments)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 8, 0)

	$avParam[8] = __LO_SetPropertyValue("IncludeImages", $bIncludeImages)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 9, 0)

	$avParam[9] = __LO_SetPropertyValue("UserDefIndex", 0)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 10, 0)

	If ($tSortField2 <> Null) Then
		If Not IsObj($tSortField2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		If $bSortColumns Then
			If Not __LO_IntIsBetween($tSortField2.Field(), 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Else
			If Not __LO_IntIsBetween($tSortField2.Field(), 0, $oRange.Rows.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
		EndIf

		ReDim $avParam[UBound($avParam) + 2]

		$avParam[$iCount] = __LO_SetPropertyValue("Col2", $tSortField2.Field() + 1)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 11, 0)

		$iCount += 1

		$avParam[$iCount] = __LO_SetPropertyValue("Ascending2", $tSortField2.IsAscending())
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 12, 0)

		$iCount += 1
	EndIf

	If ($tSortField3 <> Null) Then
		If Not IsObj($tSortField3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		If $bSortColumns Then
			If Not __LO_IntIsBetween($tSortField3.Field(), 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		Else
			If Not __LO_IntIsBetween($tSortField3.Field(), 0, $oRange.Rows.Count() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
		EndIf

		ReDim $avParam[UBound($avParam) + 2]

		$avParam[$iCount] = __LO_SetPropertyValue("Col3", $tSortField3.Field() + 1)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 13, 0)

		$iCount += 1

		$avParam[$iCount] = __LO_SetPropertyValue("Ascending3", $tSortField3.IsAscending())
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 14, 0)

		$iCount += 1
	EndIf

	$oDoc.CurrentController.Select($oRange)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 15, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 16, 0)

	$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:DataSort", "", 0, $avParam)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeSortAlt

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeValidation
; Description ...: Set or Retrieve Validation settings for a Range.
; Syntax ........: _LOCalc_RangeValidation(ByRef $oRange[, $iType = Null[, $iCondition = Null[, $sValue1 = Null[, $sValue2 = Null[, $oBaseCell = Null[, $bIgnoreBlanks = Null[, $iShowList = Null]]]]]]])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iType               - [optional] an integer value (0-7). Default is Null. The Validity check type. See Constants $LOC_VALIDATION_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iCondition          - [optional] an integer value (0-9). Default is Null. The Condition to check the cell data with. See Constants $LOC_VALIDATION_COND_* as defined in LibreOfficeCalc_Constants.au3.
;                  $sValue1             - [optional] a string value. Default is Null. If Condition is such that it requires a value, enter it here as a string.
;                  $sValue2             - [optional] a string value. Default is Null. If Condition is such that it requires a second value, enter it here as a string.
;                  $oBaseCell           - [optional] an object. Default is Null. The Cell that is used as a base for relative references in the formulas.
;                  $bIgnoreBlanks       - [optional] a boolean value. Default is Null. If True, empty cells are allowed, and not marked as invalid.
;                  $iShowList           - [optional] an integer value (0-2). Default is Null. If $iType is set to $LOC_VALIDATION_TYPE_LIST, $iShowList determines the visibility of the list. See Constants $LOC_VALIDATION_LIST_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 0 or greater than 7. See Constants $LOC_VALIDATION_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iCondition not an Integer, less than 0 or greater than 9. See Constants $LOC_VALIDATION_COND_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $sValue1 not a String.
;                  @Error 1 @Extended 5 Return 0 = $sValue2 not a String.
;                  @Error 1 @Extended 6 Return 0 = $oBaseCell not an Object.
;                  @Error 1 @Extended 7 Return 0 = $oBaseCell not a single cell Object.
;                  @Error 1 @Extended 8 Return 0 = $bIgnoreBlanks not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iShowList not an Integer, less than 0 or greater than 2. See Constants $LOC_VALIDATION_LIST_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Validation Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Cell Address.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Cell Object for referenced Cell as base cell.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iType
;                  |                               2 = Error setting $iCondition
;                  |                               4 = Error setting $sValue1
;                  |                               8 = Error setting $sValue2
;                  |                               16 = Error setting $oBaseCell
;                  |                               32 = Error setting $bIgnoreBlanks
;                  |                               64 = Error setting $iShowList
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  When $iType is set to $LOC_VALIDATION_TYPE_LIST, $sValue1 is set to a single string of words that constitute the list, each word needs to be surrounded by quotations, and separated by semicolons, such as: '"abc";"def";"ghi"'
;                  When $iType is set to $LOC_VALIDATION_TYPE_LIST, call $iCondition with $LOC_VALIDATION_COND_EQUAL.
;                  The return for $oBaseCell will always be a cell object, whether or not it is currently set or not. If it has never been set before, it will generally be cell A1.
; Related .......: _LOCalc_RangeValidationSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeValidation(ByRef $oRange, $iType = Null, $iCondition = Null, $sValue1 = Null, $sValue2 = Null, $oBaseCell = Null, $bIgnoreBlanks = Null, $iShowList = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avValid[7]
	Local $oValidation, $oCell
	Local $tCellAddress
	Local $iError = 0

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oValidation = $oRange.Validation()
	If Not IsObj($oValidation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iType, $iCondition, $sValue1, $sValue2, $oBaseCell, $bIgnoreBlanks, $iShowList) Then
		$tCellAddress = $oValidation.getSourcePosition()
		If Not IsObj($tCellAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oCell = $oRange.Spreadsheet.getCellByPosition($tCellAddress.Column(), $tCellAddress.Row())
		If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		__LO_ArrayFill($avValid, $oValidation.Type(), $oValidation.getOperator(), $oValidation.getFormula1(), $oValidation.getFormula2(), $oCell, _
				$oValidation.IgnoreBlankCells(), $oValidation.ShowList())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avValid)
	EndIf

	If ($iType <> Null) Then
		If Not __LO_IntIsBetween($iType, $LOC_VALIDATION_TYPE_ANY, $LOC_VALIDATION_TYPE_CUSTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oValidation.Type = $iType
	EndIf

	If ($iCondition <> Null) Then
		If Not __LO_IntIsBetween($iCondition, $LOC_VALIDATION_COND_NONE, $LOC_VALIDATION_COND_FORMULA) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oValidation.setOperator($iCondition)
	EndIf

	If ($sValue1 <> Null) Then
		If Not IsString($sValue1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oValidation.setFormula1($sValue1)
	EndIf

	If ($sValue2 <> Null) Then
		If Not IsString($sValue2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oValidation.setFormula2($sValue2)
	EndIf

	If ($oBaseCell <> Null) Then
		If Not IsObj($oBaseCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not ($oBaseCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0) ; Only single cells supported.

		$tCellAddress = $oBaseCell.CellAddress()
		If Not IsObj($tCellAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oValidation.setSourcePosition($tCellAddress)
	EndIf

	If ($bIgnoreBlanks <> Null) Then
		If Not IsBool($bIgnoreBlanks) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oValidation.IgnoreBlankCells = $bIgnoreBlanks
	EndIf

	If ($iShowList <> Null) Then
		If Not __LO_IntIsBetween($iShowList, $LOC_VALIDATION_LIST_INVISIBLE, $LOC_VALIDATION_LIST_SORT_ASCENDING) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oValidation.ShowList = $iShowList
	EndIf

	$oRange.Validation = $oValidation

	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : ($oRange.Validation.Type() = $iType) ? ($iError) : (BitOR($iError, 1))
	$iError = (__LO_VarsAreNull($iCondition)) ? ($iError) : ($oRange.Validation.getOperator() = $iCondition) ? ($iError) : (BitOR($iError, 2))
	$iError = (__LO_VarsAreNull($sValue1)) ? ($iError) : ($oRange.Validation.getFormula1() = $sValue1) ? ($iError) : (BitOR($iError, 4))
	$iError = (__LO_VarsAreNull($sValue2)) ? ($iError) : ($oRange.Validation.getFormula2() = $sValue2) ? ($iError) : (BitOR($iError, 8))
	$iError = (__LO_VarsAreNull($oBaseCell)) ? ($iError) : (__LOCalc_CellAddressIsSame($oRange.Validation.getSourcePosition(), $tCellAddress)) ? ($iError) : (BitOR($iError, 16))
	$iError = (__LO_VarsAreNull($bIgnoreBlanks)) ? ($iError) : ($oRange.Validation.IgnoreBlankCells() = $bIgnoreBlanks) ? ($iError) : (BitOR($iError, 32))
	$iError = (__LO_VarsAreNull($iShowList)) ? ($iError) : ($oRange.Validation.ShowList() = $iShowList) ? ($iError) : (BitOR($iError, 64))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeValidation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeValidationSettings
; Description ...: Set or Retrieve Range Validation settings.
; Syntax ........: _LOCalc_RangeValidationSettings(ByRef $oRange[, $bInputMsg = Null[, $sInputTitle = Null[, $sInputMsg = Null[, $bErrorMsg = Null[, $iErrorStyle = Null[, $sErrorTitle = Null[, $sErrorMsg = Null]]]]]]])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bInputMsg           - [optional] a boolean value. Default is Null. If True, a input message is displayed when the cell is clicked.
;                  $sInputTitle         - [optional] a string value. Default is Null. If $bInputMsg is True, the Title of the Input tip to display.
;                  $sInputMsg           - [optional] a string value. Default is Null. If $bInputMsg is True, the Message of the Input tip to display.
;                  $bErrorMsg           - [optional] a boolean value. Default is Null. If True, a error message is displayed when invalid data is entered into a cell.
;                  $iErrorStyle         - [optional] an integer value (0-3). Default is Null. The Error alert style. See Constants $LOC_VALIDATION_ERROR_ALERT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $sErrorTitle         - [optional] a string value. Default is Null. If $bErrorMsg is True, the Title of the error alert to display.
;                  $sErrorMsg           - [optional] a string value. Default is Null. If $bErrorMsg is True, the Message of the error alert to display.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bInputMsg not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sInputTitle not a String.
;                  @Error 1 @Extended 4 Return 0 = $sInputMsg not a String.
;                  @Error 1 @Extended 5 Return 0 = $bErrorMsg not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iErrorStyle not an Integer, less than 0 or greater than 3. See Constants $LOC_VALIDATION_ERROR_ALERT_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $sErrorTitle not a String.
;                  @Error 1 @Extended 8 Return 0 = $sErrorMsg not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Validation Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bInputMsg
;                  |                               2 = Error setting $sInputTitle
;                  |                               4 = Error setting $sInputMsg
;                  |                               8 = Error setting $bErrorMsg
;                  |                               16 = Error setting $iErrorStyle
;                  |                               32 = Error setting $sErrorTitle
;                  |                               64 = Error setting $sErrorMsg
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  When $iErrorStyle is set to $LOC_VALIDATION_ERROR_ALERT_MACRO, $sErrorTitle is called with the macro address to execute, the macro address will look similar to the following, filling in the data between the"<>", including the last parameter for location, which will be either application, or document: "vnd.sun.star.script:<LibraryName>.<ModuleName>.<MacroName>?language=Basic&location=<application|document>"
;                  At this time I have no functions for locating or creating macros. They may be added later.
; Related .......: _LOCalc_RangeValidation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeValidationSettings(ByRef $oRange, $bInputMsg = Null, $sInputTitle = Null, $sInputMsg = Null, $bErrorMsg = Null, $iErrorStyle = Null, $sErrorTitle = Null, $sErrorMsg = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avValid[7]
	Local $oValidation
	Local $iError = 0

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oValidation = $oRange.Validation()
	If Not IsObj($oValidation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bInputMsg, $sInputTitle, $sInputMsg, $bErrorMsg, $iErrorStyle, $sErrorTitle, $sErrorMsg) Then
		__LO_ArrayFill($avValid, $oValidation.ShowInputMessage(), $oValidation.InputTitle(), $oValidation.InputMessage(), $oValidation.ShowErrorMessage(), _
				$oValidation.ErrorAlertStyle(), $oValidation.ErrorTitle(), $oValidation.ErrorMessage())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avValid)
	EndIf

	If ($bInputMsg <> Null) Then
		If Not IsBool($bInputMsg) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oValidation.ShowInputMessage = $bInputMsg
	EndIf

	If ($sInputTitle <> Null) Then
		If Not IsString($sInputTitle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oValidation.InputTitle = $sInputTitle
	EndIf

	If ($sInputMsg <> Null) Then
		If Not IsString($sInputMsg) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oValidation.InputMessage = $sInputMsg
	EndIf

	If ($bErrorMsg <> Null) Then
		If Not IsBool($bErrorMsg) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oValidation.ShowErrorMessage = $bErrorMsg
	EndIf

	If ($iErrorStyle <> Null) Then
		If Not __LO_IntIsBetween($iErrorStyle, $LOC_VALIDATION_ERROR_ALERT_STOP, $LOC_VALIDATION_ERROR_ALERT_MACRO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oValidation.ErrorAlertStyle = $iErrorStyle
	EndIf

	If ($sErrorTitle <> Null) Then
		If Not IsString($sErrorTitle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oValidation.ErrorTitle = $sErrorTitle
	EndIf

	If ($sErrorMsg <> Null) Then
		If Not IsString($sErrorMsg) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oValidation.ErrorMessage = $sErrorMsg
	EndIf

	$oRange.Validation = $oValidation

	$iError = (__LO_VarsAreNull($bInputMsg)) ? ($iError) : ($oRange.Validation.ShowInputMessage() = $bInputMsg) ? ($iError) : (BitOR($iError, 1))
	$iError = (__LO_VarsAreNull($sInputTitle)) ? ($iError) : ($oRange.Validation.InputTitle() = $sInputTitle) ? ($iError) : (BitOR($iError, 2))
	$iError = (__LO_VarsAreNull($sInputMsg)) ? ($iError) : ($oRange.Validation.InputMessage() = $sInputMsg) ? ($iError) : (BitOR($iError, 4))
	$iError = (__LO_VarsAreNull($bErrorMsg)) ? ($iError) : ($oRange.Validation.ShowErrorMessage() = $bErrorMsg) ? ($iError) : (BitOR($iError, 8))
	$iError = (__LO_VarsAreNull($iErrorStyle)) ? ($iError) : ($oRange.Validation.ErrorAlertStyle() = $iErrorStyle) ? ($iError) : (BitOR($iError, 16))
	$iError = (__LO_VarsAreNull($sErrorTitle)) ? ($iError) : ($oRange.Validation.ErrorTitle() = $sErrorTitle) ? ($iError) : (BitOR($iError, 32))
	$iError = (__LO_VarsAreNull($sErrorMsg)) ? ($iError) : ($oRange.Validation.ErrorMessage() = $sErrorMsg) ? ($iError) : (BitOR($iError, 64))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeValidationSettings
