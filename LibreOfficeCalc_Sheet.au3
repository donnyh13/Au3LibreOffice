#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc
#include "LibreOfficeCalc_Doc.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Removing, etc. L.O. Calc document Sheets.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_SheetActivate
; _LOCalc_SheetAdd
; _LOCalc_SheetCopy
; _LOCalc_SheetCreateCursor
; _LOCalc_SheetDetectiveClear
; _LOCalc_SheetDetectiveDependent
; _LOCalc_SheetDetectiveInvalidData
; _LOCalc_SheetDetectivePrecedent
; _LOCalc_SheetDetectiveTraceError
; _LOCalc_SheetExists
; _LOCalc_SheetGetActive
; _LOCalc_SheetGetObjByName
; _LOCalc_SheetGetObjByPosition
; _LOCalc_SheetImport
; _LOCalc_SheetIsActive
; _LOCalc_SheetIsProtected
; _LOCalc_SheetLink
; _LOCalc_SheetLinkModify
; _LOCalc_SheetMove
; _LOCalc_SheetName
; _LOCalc_SheetPrintColumnsRepeat
; _LOCalc_SheetPrintRangeModify
; _LOCalc_SheetPrintRowsRepeat
; _LOCalc_SheetProtect
; _LOCalc_SheetRemove
; _LOCalc_SheetsGetCount
; _LOCalc_SheetsGetNames
; _LOCalc_SheetTabColor
; _LOCalc_SheetUnprotect
; _LOCalc_SheetVisible
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetActivate
; Description ...: Activate a Sheet in a Calc Document.
; Syntax ........: _LOCalc_SheetActivate(ByRef $oDoc, ByRef $oSheet)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Sheet was successfully activated.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetIsActive
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetActivate(ByRef $oDoc, ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.setActiveSheet($oSheet)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetActivate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetAdd
; Description ...: Insert a new Sheet into a Calc Document.
; Syntax ........: _LOCalc_SheetAdd(ByRef $oDoc[, $sName = Null[, $iPosition = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - [optional] a string value. Default is Null. The Name of the new Sheet. See remarks.
;                  $iPosition           - [optional] an integer value. Default is Null. The position to insert the new sheet. If left as Null, new sheet is inserted at the end. See remarks.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Document already contains a Sheet named the same as called in $sName.
;                  @Error 1 @Extended 4 Return 0 = $iPosition not an Integer, less than 0 or greater than number of sheets present in the document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve new Sheet's Object. New Sheet may not have been inserted successfully.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. New sheet was successfully inserted, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sName is left as Null, the sheet will be automatically named "Sheet?" where "?" is a digit.
;                  If $iPosition is left as Null, the sheet will be inserted at the end of the list of Sheets.
;                  Calling $iPosition with the number of Sheets in the Document will place the added sheet at the end of the sheet list.
; Related .......: _LOCalc_SheetRemove, _LOCalc_SheetExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetAdd(ByRef $oDoc, $sName = Null, $iPosition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets, $oSheet
	Local $iCount = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sName = Null) Then
		$sName = "Sheet" & ($oSheets.Count() + 1)
		If $oSheets.hasByName($sName) Then
			$iCount = $oSheets.Count()
			While $oSheets.hasByName($sName)
				$iCount += 1
				$sName = "Sheet" & $iCount
			WEnd
		EndIf
	EndIf

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($iPosition = Null) Then $iPosition = $oSheets.Count()

	If Not __LO_IntIsBetween($iPosition, 0, $oSheets.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSheets.insertNewByName($sName, $iPosition)

	$oSheet = $oSheets.getByName($sName)

	Return (IsObj($oSheet)) ? (SetError($__LO_STATUS_SUCCESS, 0, $oSheet)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0))
EndFunc   ;==>_LOCalc_SheetAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetCopy
; Description ...: Create a copy of a particular Sheet.
; Syntax ........: _LOCalc_SheetCopy(ByRef $oDoc, ByRef $oSheet[, $sNewName = Null[, $iPosition = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sNewName            - [optional] a string value. Default is Null. The name to assign to the newly copied Sheet.
;                  $iPosition           - [optional] an integer value. Default is Null. The position to place the copied sheet at. 0 = the beginning.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sNewName not a String.
;                  @Error 1 @Extended 4 Return 0 = Document already contains a Sheet with the same name as called in $sNewName.
;                  @Error 1 @Extended 5 Return 0 = $iPosition not an Integer, less than 0, or greater than number of Sheets contained in the document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve original Sheet's name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Object for new Sheet.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully copied the Sheet. Returning the new Sheet's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sNewName is left as Null, the original Sheet's name is used, with "_" and a digit appended.
;                  If $iPosition is left as Null, the copied sheet will be placed at the end of the list.
;                  Calling $iPosition with the number of Sheets in the Document will place the copied sheet at the end of the sheet list.
; Related .......: _LOCalc_SheetExists, _LOCalc_SheetMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetCopy(ByRef $oDoc, ByRef $oSheet, $sNewName = Null, $iPosition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets, $oNewSheet
	Local $sName
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sName = $oSheet.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If ($sNewName = Null) Then
		$sNewName = $sName & "_" & 2
		If $oSheets.hasByName($sNewName) Then
			$iCount = 2
			While $oSheets.hasByName($sNewName)
				$iCount += 1
				$sNewName = $sName & "_" & $iCount
			WEnd
		EndIf
	EndIf

	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oSheets.hasByName($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($iPosition = Null) Then $iPosition = $oSheets.Count()

	If Not __LO_IntIsBetween($iPosition, 0, $oSheets.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSheets.copyByName($sName, $sNewName, $iPosition)

	$oNewSheet = $oSheets.getByName($sNewName)
	If Not IsObj($oNewSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNewSheet)
EndFunc   ;==>_LOCalc_SheetCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetCreateCursor
; Description ...: Create a Sheet Cursor for an entire Sheet.
; Syntax ........: _LOCalc_SheetCreateCursor(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Sheet Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully created a Sheet Cursor, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Sheet Cursor can be used in functions accepting a range. When created, the Cursor will have the entire Sheet selected.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetCreateCursor(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheetCursor

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheetCursor = $oSheet.createCursor()
	If Not IsObj($oSheetCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheetCursor)
EndFunc   ;==>_LOCalc_SheetCreateCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetDetectiveClear
; Description ...: Clear all Detective marking arrows for a Sheet.
; Syntax ........: _LOCalc_SheetDetectiveClear(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. All Detective arrows were successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetDetectivePrecedent, _LOCalc_SheetDetectiveDependent, _LOCalc_SheetDetectiveTraceError, _LOCalc_SheetDetectiveInvalidData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetDetectiveClear(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheet.clearArrows()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetDetectiveClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetDetectiveDependent
; Description ...: Show or Hide Dependent marking arrows for a Cell.
; Syntax ........: _LOCalc_SheetDetectiveDependent(ByRef $oCell[, $bShow = True])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
;                  $bShow               - [optional] a boolean value. Default is True. If True, Dependent marking arrows will be added one level, if False, marking arrows will be removed one level.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell not a Cell Object.
;                  @Error 1 @Extended 3 Return 0 = $bShow not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve called Cell's Address.
;                  @Error 3 @Extended 2 Return 0 = Failed to process Show/Hide Dependent function.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if Dependents were marked/cleared one level, else false.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Each time that $bShow is called with True, another level of Dependents is marked with arrows.
; Related .......: _LOCalc_SheetDetectivePrecedent, _LOCalc_SheetDetectiveTraceError, _LOCalc_SheetDetectiveClear, _LOCalc_SheetDetectiveInvalidData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetDetectiveDependent(ByRef $oCell, $bShow = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn
	Local $tCellAddress

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.
	If Not IsBool($bShow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tCellAddress = $oCell.CellAddress()
	If Not IsObj($tCellAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $bShow Then
		$bReturn = $oCell.Spreadsheet.showDependents($tCellAddress)
		If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Else
		$bReturn = $oCell.Spreadsheet.hideDependents($tCellAddress)
		If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOCalc_SheetDetectiveDependent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetDetectiveInvalidData
; Description ...: Mark all cells containing invalid data.
; Syntax ........: _LOCalc_SheetDetectiveInvalidData(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to process Invalid Data function.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if invalid data was marked, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Validity rules must be set up for Cell(s) before invalid data will be marked.
; Related .......: _LOCalc_RangeValidation, _LOCalc_RangeValidationSettings, _LOCalc_SheetDetectivePrecedent, _LOCalc_SheetDetectiveDependent, _LOCalc_SheetDetectiveTraceError, _LOCalc_SheetDetectiveClear
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetDetectiveInvalidData(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bReturn = $oSheet.showInvalid()
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOCalc_SheetDetectiveInvalidData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetDetectivePrecedent
; Description ...: Show or Hide Precedent marking arrows for a Cell.
; Syntax ........: _LOCalc_SheetDetectivePrecedent(ByRef $oCell[, $bShow = True])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
;                  $bShow               - [optional] a boolean value. Default is True. If True, Precedent marking arrows will be added one level, if False, marking arrows will be removed one level.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell not a Cell Object.
;                  @Error 1 @Extended 3 Return 0 = $bShow not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve called Cell's Address.
;                  @Error 3 @Extended 2 Return 0 = Failed to process Show/Hide Precedent function.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if precedents were marked/cleared one level, else false.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Each time that $bShow is called with True, another level of precedents is marked with arrows.
; Related .......: _LOCalc_SheetDetectiveDependent, _LOCalc_SheetDetectiveTraceError, _LOCalc_SheetDetectiveClear, _LOCalc_SheetDetectiveInvalidData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetDetectivePrecedent(ByRef $oCell, $bShow = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn
	Local $tCellAddress

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.
	If Not IsBool($bShow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddress = $oCell.CellAddress()
	If Not IsObj($tCellAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $bShow Then
		$bReturn = $oCell.Spreadsheet.showPrecedents($tCellAddress)
		If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Else
		$bReturn = $oCell.Spreadsheet.hidePrecedents($tCellAddress)
		If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOCalc_SheetDetectivePrecedent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetDetectiveTraceError
; Description ...: Draw arrows from a Cell with errors to the cells causing the errors.
; Syntax ........: _LOCalc_SheetDetectiveTraceError(ByRef $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell not a Cell Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve called Cell's Address.
;                  @Error 3 @Extended 2 Return 0 = Failed to process Trace Error function.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if Errors were marked, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetDetectivePrecedent, _LOCalc_SheetDetectiveDependent, _LOCalc_SheetDetectiveClear, _LOCalc_SheetDetectiveInvalidData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetDetectiveTraceError(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn
	Local $tCellAddress

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	$tCellAddress = $oCell.CellAddress()
	If Not IsObj($tCellAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$bReturn = $oCell.Spreadsheet.showErrors($tCellAddress)
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOCalc_SheetDetectiveTraceError

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetExists
; Description ...: Check whether a Calc document has a Sheet with a specific name.
; Syntax ........: _LOCalc_SheetExists(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - a string value. The sheet name to check for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the document contains a Sheet matching $sName, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetExists(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_Success, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOCalc_SheetExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetGetActive
; Description ...: Retrieve a Sheet object for the currently active Sheet.
; Syntax ........: _LOCalc_SheetGetActive(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Active Sheet's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully retrieved the Active Sheet, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetGetActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheet = $oDoc.CurrentController.getActiveSheet()
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheet)
EndFunc   ;==>_LOCalc_SheetGetActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetGetObjByName
; Description ...: Retrieve a Sheet Object for a specific Sheet by name.
; Syntax ........: _LOCalc_SheetGetObjByName(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - a string value. The sheet name to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Document does not contain a sheet with name called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve requested Sheet's object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Sheet's object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetGetObjByName(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet, $oSheets

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSheet = $oSheets.getByName($sName)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheet)
EndFunc   ;==>_LOCalc_SheetGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetGetObjByPosition
; Description ...: Retrieve a Sheet Object for a specific Sheet by position.
; Syntax ........: _LOCalc_SheetGetObjByPosition(ByRef $oDoc, $iPosition)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iPosition           - an integer value. The 0 based position of the Sheet, to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iPosition not an Integer, less than 0 or greater than number of Sheets contained in the document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve requested Sheet's object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Sheet's object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Sheet position aligns with the order they are displayed at the bottom of the document. 0 based.
; Related .......: _LOCalc_SheetsGetCount, _LOCalc_SheetGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetGetObjByPosition(ByRef $oDoc, $iPosition)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet, $oSheets

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iPosition, 0, $oDoc.Sheets.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSheet = $oSheets.getByIndex($iPosition)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheet)
EndFunc   ;==>_LOCalc_SheetGetObjByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetImport
; Description ...: Import a Sheet from another Document. L.O. 3.5 and up.
; Syntax ........: _LOCalc_SheetImport(ByRef $oSourceDoc, ByRef $oDestDoc, $sSheetName[, $bInsertAfter = False])
; Parameters ....: $oSourceDoc          - [in/out] an object. The Document containing the desired Sheet. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oDestDoc            - [in/out] an object. The Document to Import the Sheet to. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sSheetName          - a string value. The Sheet's name to import from the Source Document.
;                  $bInsertAfter        - [optional] a boolean value. Default is False. If True, the Sheet is inserted after the currently active Sheet. If False, the Sheet is inserted before the currently active Sheet.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSourceDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oDestDoc not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sSheetName not a String.
;                  @Error 1 @Extended 4 Return 0 = Document called in $oSourceDoc does not have a Sheet with the name called in $sSheetName.
;                  @Error 1 @Extended 5 Return 0 = $bInsertAfter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Destination Document's currently active Sheet's position.
;                  @Error 3 @Extended 2 Return 0 = Failed to import the Sheet.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve new Sheet's Object.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office Version less than 3.5.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully imported the requested Sheet, returning the new Sheet's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetLink
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetImport(ByRef $oSourceDoc, ByRef $oDestDoc, $sSheetName, $bInsertAfter = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iPosition, $iNewSheet
	Local $oSheet

	If Not __LO_VersionCheck(3.5) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSourceDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDestDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSheetName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not $oSourceDoc.Sheets.hasByName($sSheetName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bInsertAfter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$iPosition = $oDestDoc.CurrentController.getActiveSheet().RangeAddress.Sheet()
	If Not IsInt($iPosition) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iPosition = ($bInsertAfter) ? ($iPosition + 1) : ($iPosition)

	$iNewSheet = $oDestDoc.Sheets.importSheet($oSourceDoc, $sSheetName, $iPosition)
	If Not IsInt($iNewSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSheet = $oDestDoc.Sheets.getByIndex($iNewSheet)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheet)
EndFunc   ;==>_LOCalc_SheetImport

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetIsActive
; Description ...: Check if a particular Sheet is the active Sheet.
; Syntax ........: _LOCalc_SheetIsActive(ByRef $oDoc, ByRef $oSheet)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the called Sheet is the currently active sheet, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetActivate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetIsActive(ByRef $oDoc, ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet2

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheet2 = $oDoc.CurrentController.getActiveSheet()
	If Not IsObj($oSheet2) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oSheet.AbsoluteName() = $oSheet2.AbsoluteName()) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOCalc_SheetIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetIsProtected
; Description ...: Check whether a Sheet is password protected or not.
; Syntax ........: _LOCalc_SheetIsProtected(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query Sheet's current protection status.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Successfully queried Sheet's protection status, returning a boolean indicating if the sheet is currently protected (True), or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetProtect, _LOCalc_SheetUnprotect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetIsProtected(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bReturn = $oSheet.isProtected()
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOCalc_SheetIsProtected

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetLink
; Description ...: Create a new Sheet and Link to an external Sheet in another Document.
; Syntax ........: _LOCalc_SheetLink(ByRef $oSourceDoc, ByRef $oDestDoc, $sSheetName[, $iLinkMode = $LOC_SHEET_LINK_MODE_NORMAL[, $bInsertAfter = False]])
; Parameters ....: $oSourceDoc          - [in/out] an object. The Document containing the desired Sheet. Must have been previously saved to a location. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oDestDoc            - [in/out] an object. The Document to Import the Sheet to. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sSheetName          - a string value. The Sheet's name to import from the Source Document.
;                  $iLinkMode           - [optional] an integer value (0-2). Default is $LOC_SHEET_LINK_MODE_NORMAL. The content to link from the Sheet. See Constants $LOC_SHEET_LINK_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bInsertAfter        - [optional] a boolean value. Default is False. If True, the Sheet is inserted after the currently active Sheet. If False, the Sheet is inserted before the currently active Sheet.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSourceDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oDestDoc not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sSheetName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iLinkMode not an Integer, less than 0, or greater than 2. See Constants $LOC_SHEET_LINK_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = Document called in $oSourceDoc has no save location.
;                  @Error 1 @Extended 6 Return 0 = Document called in $oSourceDoc does not have a Sheet with the name called in $sSheetName.
;                  @Error 1 @Extended 7 Return 0 = $bInsertAfter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Create a name for new Sheet in Destination Document.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Document's currently active Sheet's position.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve new Sheet's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully inserted and linked the new Sheet, returning the new Sheet's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetImport
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetLink(ByRef $oSourceDoc, ByRef $oDestDoc, $sSheetName, $iLinkMode = $LOC_SHEET_LINK_MODE_NORMAL, $bInsertAfter = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName
	Local $iCount, $iPosition
	Local $oSheet

	If Not IsObj($oSourceDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDestDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSheetName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iLinkMode, $LOC_SHEET_LINK_MODE_NONE, $LOC_SHEET_LINK_MODE_VALUE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($oSourceDoc.URL() = "") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not $oSourceDoc.Sheets.hasByName($sSheetName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bInsertAfter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If $oDestDoc.Sheets.hasByName($sSheetName) Then
		$sName = $sSheetName & "_2"
		If $oDestDoc.Sheets.hasByName($sName) Then
			$iCount = 2
			While $oDestDoc.Sheets.hasByName($sName)
				$iCount += 1
				$sName = $sSheetName & "_" & $iCount
			WEnd
		EndIf

	Else
		$sName = $sSheetName
	EndIf

	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iPosition = $oDestDoc.CurrentController.getActiveSheet().RangeAddress.Sheet()
	If Not IsInt($iPosition) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iPosition = ($bInsertAfter) ? ($iPosition + 1) : ($iPosition - 1)

	$oDestDoc.Sheets.insertNewByName($sName, $iPosition)

	$oSheet = $oDestDoc.Sheets.getByName($sName)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oSheet.link($oSourceDoc.URL(), $sSheetName, "", "", $iLinkMode)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheet)
EndFunc   ;==>_LOCalc_SheetLink

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetLinkModify
; Description ...: Set or Retrieve the Linked settings of a Sheet.
; Syntax ........: _LOCalc_SheetLinkModify(ByRef $oSheet[, $oNewDoc = Null[, $sSheetName = Null[, $iLinkMode = Null]]])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oNewDoc             - [optional] an object. Default is Null. A New Document to link a new sheet from. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sSheetName          - [optional] a string value. Default is Null. The Sheet name to link to.
;                  $iLinkMode           - [optional] an integer value (0-2). Default is Null. The content to link from the Sheet. See Constants $LOC_SHEET_LINK_MODE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oNewDoc not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sSheetName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iLinkMode not an Integer, less than 0, or greater than 2. See Constants $LOC_SHEET_LINK_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = Document called in $oNewDoc has no save location.
;                  @Error 1 @Extended 6 Return 0 = Document called in $oNewDoc does not have a Sheet with the name called in $sSheetName.
;                  @Error 1 @Extended 7 Return 0 = Document currently set as source does not contain a Sheet with the name as called in $sSheetName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to set $sSheetName, Sheet has not been linked before, and has no source document to find $sSheetName in.
;                  @Error 3 @Extended 2 Return 0 = Failed to open source document set for this Sheet.
;                  @Error 3 @Extended 3 Return 0 = Failed to set $iLinkMode, Sheet isn't currently linked.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oNewDoc
;                  |                               2 = Error setting $sSheetName
;                  |                               4 = Error setting $iLinkMode
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  The Return for $oNewDoc is the File path of the linked Document.
;                  If you set $oNewDoc, you must set $sSheetName and $iLinkMode also.
;                  To unlink a Sheet, call $iLinkMode with $LOC_SHEET_LINK_MODE_NONE.
; Related .......: _LOCalc_SheetLink
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetLinkModify(ByRef $oSheet, $oNewDoc = Null, $sSheetName = Null, $iLinkMode = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSourceDoc
	Local $iError = 0
	Local $bClose = True
	Local $avSheet[3]

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($oNewDoc, $sSheetName, $iLinkMode) Then
		__LO_ArrayFill($avSheet, _LO_PathConvert($oSheet.LinkUrl(), $LO_PATHCONV_PCPATH_RETURN), $oSheet.LinkSheetName(), $oSheet.LinkMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSheet)
	EndIf

	If ($oNewDoc <> Null) Then
		If Not IsObj($oNewDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		If Not IsString($sSheetName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If Not __LO_IntIsBetween($iLinkMode, $LOC_SHEET_LINK_MODE_NONE, $LOC_SHEET_LINK_MODE_VALUE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If ($oNewDoc.URL() = "") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not $oNewDoc.Sheets.hasByName($sSheetName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oSheet.link($oSourceDoc.URL(), $sSheetName, "", "", $iLinkMode)

		$iError = ($oSheet.LinkUrl() = $oNewDoc.URL()) ? ($iError) : (BitOR($iError, 1))
		$iError = ($oSheet.LinkSheetName() = $sSheetName) ? ($iError) : (BitOR($iError, 2))
		$iError = ($oSheet.LinkMode() = $iLinkMode) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($sSheetName <> Null) Then
		If Not IsString($sSheetName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If ($oSheet.LinkUrl() = "") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oSourceDoc = _LOCalc_DocOpen(_LO_PathConvert($oSheet.LinkUrl(), $LO_PATHCONV_PCPATH_RETURN), True, True)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$bClose = (@extended = 2) ? (True) : (False)
		If Not $oSourceDoc.Sheets.hasByName($sSheetName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $oSourceDoc.Close(True))

		If $bClose Then $oSourceDoc.Close(True)

		$oSheet.LinkSheetName = $sSheetName
		$iError = ($oSheet.LinkSheetName() = $sSheetName) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iLinkMode <> Null) Then
		If Not __LO_IntIsBetween($iLinkMode, $LOC_SHEET_LINK_MODE_NONE, $LOC_SHEET_LINK_MODE_VALUE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If ($iLinkMode = $LOC_SHEET_LINK_MODE_NONE) Then
			With $oSheet
				.LinkUrl = ""
				.LinkSheetName = ""
				.LinkMode = $iLinkMode
			EndWith
			$iError = ($oSheet.LinkUrl() = "") ? ($iError) : (BitOR($iError, 1))
			$iError = ($oSheet.LinkSheetName() = "") ? ($iError) : (BitOR($iError, 2))
			$iError = ($oSheet.LinkMode() = $iLinkMode) ? ($iError) : (BitOR($iError, 4))

		Else
			If ($oSheet.LinkUrl() = "") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$oSheet.LinkMode = $iLinkMode
			$iError = ($oSheet.LinkMode() = $iLinkMode) ? ($iError) : (BitOR($iError, 4))
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_SheetLinkModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetMove
; Description ...: Set or Retrieve a Sheet's position in the list of Sheets in a Calc Document.
; Syntax ........: _LOCalc_SheetMove(ByRef $oDoc, ByRef $oSheet[, $iPosition = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iPosition           - [optional] an integer value. Default is Null.The Position the move the Sheet to, 0 being the beginning.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iPosition not an Integer, less than 0 or greater than number of sheets contained in the document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheet's name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Sheets Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Sheet was successfully moved.
;                  @Error 0 @Extended 0 Return Integer = Success. $iPosition called with Null, returning Sheet's current position.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling $iPosition with the number of Sheets in the Document will place the moved sheet at the end of the sheet list.
; Related .......: _LOCalc_SheetCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetMove(ByRef $oDoc, ByRef $oSheet, $iPosition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets
	Local $sName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($iPosition = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oSheet.RangeAddress.Sheet())

	$sName = $oSheet.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iPosition, 0, $oSheets.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSheets.moveByName($sName, $iPosition)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetName
; Description ...: Set or Retrieve a Sheet's name.
; Syntax ........: _LOCalc_SheetName(ByRef $oDoc, ByRef $oSheet[, $sName = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sName               - [optional] a string value. Default is Null. The new name for the Sheet.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = Document already has a Sheet named the same as called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Sheet's new name was successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Sheet's current name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOCalc_SheetExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetName(ByRef $oDoc, ByRef $oSheet, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($sName = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oSheet.Name())

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSheet.Name = $sName

	$iError = ($oSheet.Name() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_SheetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetPrintColumnsRepeat
; Description ...: Set or Retrieve settings for Column Header print settings.
; Syntax ........: _LOCalc_SheetPrintColumnsRepeat(ByRef $oSheet[, $oRange = Null[, $bRepeatColumns = Null]])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRange              - [optional] an object. Default is Null. The Range containing the Column Headers. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bRepeatColumns      - [optional] a boolean value. Default is Null. If True, the Column Titles are repeated on each printed page to the right.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bRepeatColumns not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Range Object of current Column Headers.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Cell/Cell Range Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve current Column Header Range.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve called Range Address.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oRange
;                  |                               2 = Error setting $bRepeatColumns
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call $oRange with the Default keyword to reset the Column Range to none.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  The Return for $oRange will be a Cell Range or Cell Object, if no Column is set, the return will always be the first cell (A1) of the Sheet.
; Related .......: _LOCalc_SheetPrintRowsRepeat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetPrintColumnsRepeat(ByRef $oSheet, $oRange = Null, $bRepeatColumns = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tRangeAddr
	Local $oCell, $oCellRange
	Local $iError = 0
	Local $avPrintColumn[2]

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($oRange, $bRepeatColumns) Then
		$tRangeAddr = $oSheet.getTitleColumns()
		If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If ($tRangeAddr.StartColumn() = $tRangeAddr.EndColumn()) And ($tRangeAddr.StartRow() = $tRangeAddr.EndRow()) Then
			$oCell = $oSheet.getCellByPosition($tRangeAddr.StartColumn(), $tRangeAddr.StartRow())
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			__LO_ArrayFill($avPrintColumn, $oCell, $oSheet.PrintTitleColumns())

		Else
			$oCellRange = $oSheet.getCellRangeByPosition($tRangeAddr.StartColumn(), $tRangeAddr.StartRow(), $tRangeAddr.EndColumn(), $tRangeAddr.EndRow())
			If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			__LO_ArrayFill($avPrintColumn, $oCellRange, $oSheet.PrintTitleColumns())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPrintColumn)
	EndIf

	If ($oRange = Default) Then
		$tRangeAddr = $oSheet.getTitleColumns()
		If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		With $tRangeAddr
			.Sheet = 0
			.StartColumn = 0
			.EndColumn = 0
			.StartRow = 0
			.EndRow = 0
		EndWith
		$oSheet.setTitleColumns($tRangeAddr)
		$iError = (__LOCalc_RangeAddressIsSame($oSheet.getTitleColumns(), $tRangeAddr)) ? ($iError) : (BitOR($iError, 1))

	ElseIf ($oRange <> Null) Then
		If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tRangeAddr = $oRange.RangeAddress()
		If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oSheet.setTitleColumns($tRangeAddr)
		$iError = (__LOCalc_RangeAddressIsSame($oSheet.getTitleColumns(), $tRangeAddr)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bRepeatColumns <> Null) Then
		If Not IsBool($bRepeatColumns) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oSheet.PrintTitleColumns = $bRepeatColumns
		$iError = ($oSheet.PrintTitleColumns() = $bRepeatColumns) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_SheetPrintColumnsRepeat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetPrintRangeModify
; Description ...: Set or Retrieve an array of Print Ranges for a Sheet.
; Syntax ........: _LOCalc_SheetPrintRangeModify(ByRef $oSheet[, $aoRange = Null])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $aoRange             - [optional] an array of objects. Default is Null. An Array of Ranges that are marked for printing. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $aoRange not an Array.
;                  @Error 1 @Extended 3 Return ? = Element in Array called in $aoRange not an Object, returning problem array element number.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Print Area Addresses.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Cell Range Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $aoRange
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Print Range was successfully set.
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were set to Null, returning Array of Range Objects set to be printed. @extended set to number of Ranges. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call $aoRange with an empty Array (local $aArray[0]) to set the whole sheet to be printed (default), instead of a specific range.
;                  Likewise, when retrieving the current print range, if the returned array is empty, that means the entire sheet (all of the used cells) is set to be printed.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  If you set a Print Range for one sheet, any other Sheet without a Print Range set will NOT be printed.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetPrintRangeModify(ByRef $oSheet, $aoRange = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($aoRange = Null) Then
		$aoRange = $oSheet.getPrintAreas()
		If Not IsArray($aoRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		For $i = 0 To UBound($aoRange) - 1
			$aoRange[$i] = $oSheet.getCellRangeByPosition($aoRange[$i].StartColumn(), $aoRange[$i].StartRow(), $aoRange[$i].EndColumn(), $aoRange[$i].EndRow())
			If Not IsObj($aoRange[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($aoRange), $aoRange)
	EndIf

	If Not IsArray($aoRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	For $i = 0 To UBound($aoRange) - 1
		If Not IsObj($aoRange[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, $i)

		$aoRange[$i] = $aoRange[$i].RangeAddress()
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$oSheet.setPrintAreas($aoRange)

	Return (UBound($oSheet.getPrintAreas()) = UBound($aoRange)) ? SetError($__LO_STATUS_SUCCESS, 0, 1) : SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)
EndFunc   ;==>_LOCalc_SheetPrintRangeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetPrintRowsRepeat
; Description ...: Set or Retrieve settings for Row Header print settings.
; Syntax ........: _LOCalc_SheetPrintRowsRepeat(ByRef $oSheet[, $oRange = Null[, $bRepeatRows = Null]])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRange              - [optional] an object. Default is Null. The Range containing the Row Headers. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bRepeatRows         - [optional] a boolean value. Default is Null. If True, the Row Titles are repeated on each printed page to the bottom.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bRepeatRows not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Range Object of current Row Headers.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Cell/Cell Range Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve current Row Header Range.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve called Range Address.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oRange
;                  |                               2 = Error setting $bRepeatRows
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call $oRange with the Default keyword to reset the Row Range to none.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  The Return for $oRange will be a Cell Range or Cell Object, if no Row is set, the return will always be the first cell (A1) of the Sheet.
; Related .......: _LOCalc_SheetPrintColumnsRepeat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetPrintRowsRepeat(ByRef $oSheet, $oRange = Null, $bRepeatRows = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tRangeAddr
	Local $oCell, $oCellRange
	Local $iError = 0
	Local $avPrintRow[2]

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($oRange, $bRepeatRows) Then
		$tRangeAddr = $oSheet.getTitleRows()
		If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If ($tRangeAddr.StartColumn() = $tRangeAddr.EndColumn()) And ($tRangeAddr.StartRow() = $tRangeAddr.EndRow()) Then
			$oCell = $oSheet.getCellByPosition($tRangeAddr.StartColumn(), $tRangeAddr.StartRow())
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			__LO_ArrayFill($avPrintRow, $oCell, $oSheet.PrintTitleRows())

		Else
			$oCellRange = $oSheet.getCellRangeByPosition($tRangeAddr.StartColumn(), $tRangeAddr.StartRow(), $tRangeAddr.EndColumn(), $tRangeAddr.EndRow())
			If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			__LO_ArrayFill($avPrintRow, $oCellRange, $oSheet.PrintTitleRows())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPrintRow)
	EndIf

	If ($oRange = Default) Then
		$tRangeAddr = $oSheet.getTitleRows()
		If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		With $tRangeAddr
			.Sheet = 0
			.StartColumn = 0
			.EndColumn = 0
			.StartRow = 0
			.EndRow = 0
		EndWith
		$oSheet.setTitleRows($tRangeAddr)
		$iError = (__LOCalc_RangeAddressIsSame($oSheet.getTitleRows(), $tRangeAddr)) ? ($iError) : (BitOR($iError, 1))

	ElseIf ($oRange <> Null) Then
		If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tRangeAddr = $oRange.RangeAddress()
		If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oSheet.setTitleRows($tRangeAddr)
		$iError = (__LOCalc_RangeAddressIsSame($oSheet.getTitleRows(), $tRangeAddr)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bRepeatRows <> Null) Then
		If Not IsBool($bRepeatRows) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oSheet.PrintTitleRows = $bRepeatRows
		$iError = ($oSheet.PrintTitleRows() = $bRepeatRows) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_SheetPrintRowsRepeat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetProtect
; Description ...: Password protect a sheet from modification.
; Syntax ........: _LOCalc_SheetProtect(ByRef $oSheet, $sPassword)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sPassword           - a string value. The password to protect the sheet with.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sPassword not a String.
;                  @Error 1 @Extended 3 Return 0 = String called in $sPassword contains no letters, digits, or underscores.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to protect the sheet.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Sheet was successfully protected with the called password.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetUnprotect, _LOCalc_SheetIsProtected
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetProtect(ByRef $oSheet, $sPassword)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($sPassword = "") Or Not StringRegExp($sPassword, "[\w]") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Password contains no letters, digits, or underscores.

	$oSheet.Protect($sPassword)

	Return ($oSheet.isProtected()) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_SheetProtect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetRemove
; Description ...: Remove a Sheet from a Calc Document.
; Syntax ........: _LOCalc_SheetRemove(ByRef $oDoc, ByRef $oSheet)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the Sheet's name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Sheets Object.
;                  @Error 3 @Extended 3 Return 0 = Attempted to delete the Sheet, but a Sheet by that name still exists.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully removed the requested sheet.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetAdd, _LOCalc_SheetGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetRemove(ByRef $oDoc, ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets
	Local $sName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sName = $oSheet.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSheets.removeByName($sName)

	If $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetRemove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetsGetCount
; Description ...: Retrieve a count of Sheets contained in a Calc Document.
; Syntax ........: _LOCalc_SheetsGetCount(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Sheets contained in the Calc Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetsGetCount(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheets.Count())
EndFunc   ;==>_LOCalc_SheetsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetsGetNames
; Description ...: Retrieve an array of Sheet names for a Calc Document.
; Syntax ........: _LOCalc_SheetsGetNames(ByRef $oDoc[, $bLinkedOnly = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bLinkedOnly         - [optional] a boolean value. Default is False. If True, only Sheets that are linked will be returned. If False, all Sheets will be returned.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bLinkedOnly not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of Sheets.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve SheetLinks Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve count of Sheet Links.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Linked Sheet's object.
;                  @Error 3 @Extended 6 Return 0 =Failed to retrieve Sheet's object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Sheet names for this document. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetsGetNames(ByRef $oDoc, $bLinkedOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets, $oSheetLinks, $oLinkedSheet, $oSheet
	Local $iLinkedSheetCount = 0, $iSheetsCount = 0
	Local $asNames[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bLinkedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iSheetsCount = $oSheets.Count()
	If Not IsInt($iSheetsCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If $bLinkedOnly Then
		$oSheetLinks = $oDoc.SheetLinks()
		If Not IsObj($oSheetLinks) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$iLinkedSheetCount = $oSheetLinks.Count()
		If Not IsInt($iLinkedSheetCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		ReDim $asNames[$iLinkedSheetCount]

		For $i = 0 To $iLinkedSheetCount - 1
			$oLinkedSheet = $oSheetLinks.getByIndex($i)
			If Not IsObj($oLinkedSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

			For $k = 0 To $iSheetsCount - 1
				$oSheet = $oSheets.getByIndex($k)
				If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

				If ($oLinkedSheet.Url() = $oSheet.LinkUrl()) Then
					$asNames[$i] = $oSheet.Name()
					ExitLoop
				EndIf
				Sleep((IsInt($k / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
			Next
		Next

	Else
		ReDim $asNames[$iSheetsCount]

		For $i = 0 To $iSheetsCount - 1
			$asNames[$i] = $oSheets.getByIndex($i).Name()

			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOCalc_SheetsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetTabColor
; Description ...: Set or Retrieve a Sheet's Tab Color.
; Syntax ........: _LOCalc_SheetTabColor(ByRef $oSheet[, $iColor = Null])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is Null. The tab color in Long Color format. Set to $LO_COLOR_OFF(-1) to set to Default color setting. Can also be one of the constants $LO_COLOR_* as defined in LibreOffice_Constants.au3
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current Tab Color as an Integer
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetTabColor(ByRef $oSheet, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($iColor = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oSheet.TabColor())

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheet.TabColor = $iColor
	If Not ($oSheet.TabColor() = $iColor) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetTabColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetUnprotect
; Description ...: Remove password protection from a Sheet.
; Syntax ........: _LOCalc_SheetUnprotect(ByRef $oSheet, $sPassword)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sPassword           - a string value. The password previously used to protect the sheet.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sPassword not a String.
;                  @Error 1 @Extended 3 Return 0 = String called in $sPassword contains no letters, digits, or underscores.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Password called in $sPassword is incorrect.
;                  @Error 3 @Extended 2 Return 0 = Failed to unprotect the sheet.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Sheet was successfully unprotected with the called password.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetProtect, _LOCalc_SheetIsProtected
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetUnprotect(ByRef $oSheet, $sPassword)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($sPassword = "") Or Not StringRegExp($sPassword, "[\w]") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Password contains no letters, digits, or underscores.

	$oSheet.Unprotect($sPassword)

	If ($oCOM_ErrorHandler.number() = -2147352567) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Wrong password

	Return ($oSheet.isProtected()) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_SheetUnprotect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetVisible
; Description ...: Set or Retrieve a Sheet's current visibility setting.
; Syntax ........: _LOCalc_SheetVisible(ByRef $oSheet[, $bVisible = Null])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the Sheet is visible in the Libre Office UI.
; Return values .: Success: 1 or Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisiblee
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Sheet Visibility setting was successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bVisible set to Null, returning current visibility setting. True indicates the Sheet is currently visible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetVisible(ByRef $oSheet, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oSheet.IsVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheet.IsVisible = $bVisible

	Return ($oSheet.IsVisible = $bVisible) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_SheetVisible
