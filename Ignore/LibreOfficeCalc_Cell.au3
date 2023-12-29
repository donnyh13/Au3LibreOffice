#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc


; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Removing, etc. L.O. Calc document Cells.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_CellFormula
; _LOCalc_CellGetType
; _LOCalc_CellText
; _LOCalc_CellValue
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellFormula
; Description ...: Set or Retrieve a Cell's Formula.
; Syntax ........: _LOCalc_CellFormula(ByRef $oCell[, $sFormula = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function.
;                  $sFormula            - [optional] a string value. Default is Null. The Formula to set the Cell to. Overwrites any previous data.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;				   @Error 1 @Extended 3 Return 0 = $sFormula not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $sFormula
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Cell's current formula as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Only individual cells are supported, not cell ranges.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current Cell content.
; Related .......: _LOCalc_CellGetType, _LOCalc_CellText, _LOCalc_CellValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellFormula(ByRef $oCell, $sFormula = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($oCell.ImplementationName() <> "ScCellObj") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	If ($sFormula = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oCell.getFormula())

	If Not IsString($sFormula) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCell.setFormula($sFormula)
	If ($oCell.getFormula() <> $sFormula) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CellFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellGetType
; Description ...: Retrieve the type of data that is contained in a Cell.
; Syntax ........: _LOCalc_CellGetType(ByRef $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Data Type.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Returning type of data contained in the Cell. Return value will be one of the constants, $LOC_CELL_TYPE_* as defined in LibreOfficeCalc_Constants.au3
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Only individual cells are supported, not cell ranges.
; Related .......: _LOCalc_CellText, _LOCalc_CellFormula, _LOCalc_CellValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellGetType(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCellType

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($oCell.ImplementationName() <> "ScCellObj") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	$iCellType = $oCell.getType()
	If Not IsInt($iCellType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCellType)
EndFunc   ;==>_LOCalc_CellGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellText
; Description ...: Set or Retrieve a Cell's Text content.
; Syntax ........: _LOCalc_CellText(ByRef $oCell[, $sText = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function.
;                  $sText               - [optional] a string value. Default is Null. The Text to set the Cell contents to. Overwrites any previous data.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;				   @Error 1 @Extended 3 Return 0 = $sText not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $sText
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Cell's contents as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Only individual cells are supported, not cell ranges.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current Cell content.
; Related .......: _LOCalc_CellGetType, _LOCalc_CellFormula, _LOCalc_CellValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellText(ByRef $oCell, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($oCell.ImplementationName() <> "ScCellObj") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	If ($sText = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oCell.getString())

	If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCell.setString($sText)
	If ($oCell.getString() <> $sText) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CellText

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellValue
; Description ...: Set or Retrieve a Cell's Value.
; Syntax ........: _LOCalc_CellValue(ByRef $oCell[, $nValue = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function.
;                  $nValue              - [optional] a general number value. Default is Null. The Value to set the Cell to. Overwrites any previous data.
; Return values .: Success: 1 or Number.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;				   @Error 1 @Extended 3 Return 0 = $nValue not a Number.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $nValue
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Number = Success. All optional parameters were set to Null, returning the Cell's current number value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Only individual cells are supported, not cell ranges.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current Cell content.
; Related .......: _LOCalc_CellGetType, _LOCalc_CellText, _LOCalc_CellFormula
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellValue(ByRef $oCell, $nValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($oCell.ImplementationName() <> "ScCellObj") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	If ($nValue = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oCell.getValue())

	If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCell.setValue($nValue)
	If ($oCell.getValue() <> $nValue) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CellValue
