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

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Removing, etc. L.O. Calc document Cells.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_CellBackColor
; _LOCalc_CellBorderColor
; _LOCalc_CellBorderPadding
; _LOCalc_CellBorderStyle
; _LOCalc_CellBorderWidth
; _LOCalc_CellCreateTextCursor
; _LOCalc_CellEffect
; _LOCalc_CellFont
; _LOCalc_CellFontColor
; _LOCalc_CellFormula
; _LOCalc_CellGetType
; _LOCalc_CellNumberFormat
; _LOCalc_CellOverline
; _LOCalc_CellProtection
; _LOCalc_CellShadow
; _LOCalc_CellStrikeOut
; _LOCalc_CellString
; _LOCalc_CellTextAlign
; _LOCalc_CellTextOrient
; _LOCalc_CellTextProperties
; _LOCalc_CellUnderline
; _LOCalc_CellValue
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellBackColor
; Description ...: Set or Retrieve the Cell or Cell Range Background color.
; Syntax ........: _LOCalc_CellBackColor(ByRef $oCell[, $iBackColor = Null[, $bBackTransparent = Null]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Cell background color as a Long Integer. Set to $LO_COLOR_OFF(-1) to disable Background color. Can also be one of the constants $LO_COLOR_* as defined in LibreOffice_Constants.au3
;                  $bBackTransparent    - [optional] a boolean value. Default is Null. If True, the background color is transparent.
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Passed Object to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iBackColor not an Integer, set less than -1 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $bBackTransparent not an Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  |                               2 = Error setting $bBackTransparent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameters with Null keyword to skip it.
; Related .......: _LO_ConvertColorToLong, _LO_ConvertColorFromLong, _LOCalc_CellStyleBackColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellBackColor(ByRef $oCell, $iBackColor = Null, $bBackTransparent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellBackColor($oCell, $iBackColor, $bBackTransparent)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellBorderColor
; Description ...: Set and Retrieve the Cell or Cell Range Border Line Color. Libre Office Version 3.6 and Up.
; Syntax ........: _LOCalc_CellBorderColor(ByRef $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $iVert = Null[, $iHori = Null[, $iTLBRDiag = Null[, $iBLTRDiag = Null]]]]]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. Set the Top Border Line Color of the Cell Range in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. Set the Bottom Border Line Color of the Cell Range in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. Set the Left Border Line Color of the Cell Range in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. Set the Right Border Line Color of the Cell Range in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iVert               - [optional] an integer value. Default is Null. Set the Vertical Border Line Color of the Cell Range in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iHori               - [optional] an integer value. Default is Null. Set the Horizontal Border Line Color of the Cell Range in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iTLBRDiag           - [optional] an integer value (0-16777215). Default is Null. Set the Top-Left to Bottom-Right Diagonal Border Line Color of the Cell Range in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBLTRDiag           - [optional] an integer value (0-16777215). Default is Null. Set the Bottom-Left to Top-Right Diagonal Border Line Color of the Cell Range in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 7 Return 0 = $iVert not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 8 Return 0 = $iHori not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 9 Return 0 = $iTLBRDiag not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 10 Return 0 = $iBLTRDiag not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 11 Return 0 = Variable passed to internal function not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Error Retrieving TableBorder2 Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Color when Right Border width not set.
;                  @Error 4 @Extended 5 Return 0 = Cannot set Vertical Border Color when Vertical Border width not set.
;                  @Error 4 @Extended 6 Return 0 = Cannot set Horizontal Border Color when Horizontal Border width not set.
;                  @Error 4 @Extended 7 Return 0 = Cannot set Top-Left to Bottom-Right Diagonal Border Color when Top-Left to Bottom-Right Diagonal Border width not set.
;                  @Error 4 @Extended 8 Return 0 = Cannot set Bottom-Left to Top-Right Diagonal Border Color when Bottom-Left to Top-Right Diagonal Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellBorderWidth, _LOCalc_CellBorderStyle, _LOCalc_CellBorderColor, _LO_ConvertColorToLong, _LO_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellBorderColor(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $iVert = Null, $iHori = Null, $iTLBRDiag = Null, $iBLTRDiag = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($iVert <> Null) And Not __LO_IntIsBetween($iVert, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If ($iHori <> Null) And Not __LO_IntIsBetween($iHori, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If ($iTLBRDiag <> Null) And Not __LO_IntIsBetween($iTLBRDiag, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If ($iBLTRDiag <> Null) And Not __LO_IntIsBetween($iBLTRDiag, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

	$vReturn = __LOCalc_CellBorder($oCell, False, False, True, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori, $iTLBRDiag, $iBLTRDiag)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellBorderPadding
; Description ...: Set or retrieve the Cell or Cell Range Border Padding settings.
; Syntax ........: _LOCalc_CellBorderPadding(ByRef $oCell[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Micrometers (uM).
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Distance between the Border and Cell contents, in Micrometers(uM).
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Distance between the Border and Cell contents, in Micrometers(uM).
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Distance between the Border and Cell contents, in Micrometers(uM).
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Distance between the Border and Cell contents, in Micrometers(uM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iAll not an Integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iTop not an Integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iBottom not an Integer, or less than 0.
;                  @Error 1 @Extended 7 Return 0 = $iLeft not an Integer, or less than 0.
;                  @Error 1 @Extended 8 Return 0 = $iRight not an Integer, or less than 0.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $iAll returns an integer value if all (Top, Bottom, Left, Right) padding values are equal, else Null is returned.
; Related .......: _LOCalc_CellStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellBorderPadding(ByRef $oCell, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellBorderPadding($oCell, $iAll, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellBorderStyle
; Description ...: Set and retrieve the Cell or Cell Range Border Line style. Libre Office Version 3.6 and Up.
; Syntax ........: _LOCalc_CellBorderStyle(ByRef $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $iVert = Null[, $iHori = Null[, $iTLBRDiag = Null[, $iBLTRDiag = Null]]]]]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Top Border Line Style of the Cell Range using one of the line style constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Bottom Border Line Style of the Cell Range using one of the line style constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Left Border Line Style of the Cell Range using one of the line style constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Right Border Line Style of the Cell Range using one of the line style constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iVert               - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Vertical Border Line Style of the Cell Range using one of the line style constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iHori               - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Horizontal Border Line Style of the Cell Range using one of the line style constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iTLBRDiag           - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Top-Left to Bottom-Right Diagonal Border Line Style of the Cell Range using one of the line style constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBLTRDiag           - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Bottom-Left to Top-Right Diagonal Border Line Style of the Cell Range using one of the line style constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0. See Constants $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0. See Constants $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0. See Constants $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0. See Constants $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iVert not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0. See Constants $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iHori not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0. See Constants $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $iTLBRDiag not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0. See Constants $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $iBLTRDiag not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0. See Constants $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = Variable passed to internal function not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Error Retrieving TableBorder2 Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style when Right Border width not set.
;                  @Error 4 @Extended 5 Return 0 = Cannot set Vertical Border Style when Vertical Border width not set.
;                  @Error 4 @Extended 6 Return 0 = Cannot set Horizontal Border Style when Horizontal Border width not set.
;                  @Error 4 @Extended 7 Return 0 = Cannot set Top-Left to Bottom-Right Diagonal Border Style when Top-Left to Bottom-Right Diagonal Border width not set.
;                  @Error 4 @Extended 8 Return 0 = Cannot set Bottom-Left to Top-Right Diagonal Border Style when Bottom-Left to Top-Right Diagonal Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellBorderWidth, _LOCalc_CellBorderColor, _LOCalc_CellStyleBorderStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellBorderStyle(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $iVert = Null, $iHori = Null, $iTLBRDiag = Null, $iBLTRDiag = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($iVert <> Null) And Not __LO_IntIsBetween($iVert, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If ($iHori <> Null) And Not __LO_IntIsBetween($iHori, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If ($iTLBRDiag <> Null) And Not __LO_IntIsBetween($iTLBRDiag, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If ($iBLTRDiag <> Null) And Not __LO_IntIsBetween($iBLTRDiag, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

	$vReturn = __LOCalc_CellBorder($oCell, False, True, False, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori, $iTLBRDiag, $iBLTRDiag)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellBorderWidth
; Description ...: Set and retrieve the Cell or Cell Range Border Line Width settings. Libre Office Version 3.6 and Up.
; Syntax ........: _LOCalc_CellBorderWidth(ByRef $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $iVert = Null[, $iHori = Null[, $iTLBRDiag = Null[, $iBLTRDiag = Null]]]]]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Border Line width of the Cell Range in Micrometers. Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Border Line width of the Cell Range in Micrometers. Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Border Line width of the Cell Range in Micrometers. Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Border Line width of the Cell Range in Micrometers. Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iVert               - [optional] an integer value. Default is Null. Set the Vertical Border Line width of the Cell Range in Micrometers. Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iHori               - [optional] an integer value. Default is Null.Set the Horizontal Border Line width of the Cell Range in Micrometers. Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iTLBRDiag           - [optional] an integer value. Default is Null. Set the Top-Left to Bottom-Right Diagonal Border Line width of the Cell Range in Micrometers. Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBLTRDiag           - [optional] an integer value. Default is Null. Set the Bottom-Left to Top-Right Diagonal Border Line width of the Cell Range in Micrometers. Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or less than 0.
;                  @Error 1 @Extended 7 Return 0 = $iVert not an integer, or less than 0.
;                  @Error 1 @Extended 8 Return 0 = $iHori not an integer, or less than 0.
;                  @Error 1 @Extended 9 Return 0 = $iTLBRDiag not an integer, or less than 0.
;                  @Error 1 @Extended 10 Return 0 = $iBLTRDiag not an integer, or less than 0.
;                  @Error 1 @Extended 11 Return 0 = Variable passed to internal function not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Error Retrieving TableBorder2 Object.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: For some reason, Horizontal line width may change depending on either top/bottom line widths or vertical line width.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellBorderStyle, _LOCalc_CellBorderColor, _LOCalc_CellStyleBorderWidth
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellBorderWidth(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $iVert = Null, $iHori = Null, $iTLBRDiag = Null, $iBLTRDiag = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($iVert <> Null) And Not __LO_IntIsBetween($iVert, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If ($iHori <> Null) And Not __LO_IntIsBetween($iHori, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If ($iTLBRDiag <> Null) And Not __LO_IntIsBetween($iTLBRDiag, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If ($iBLTRDiag <> Null) And Not __LO_IntIsBetween($iBLTRDiag, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

	$vReturn = __LOCalc_CellBorder($oCell, True, False, False, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori, $iTLBRDiag, $iBLTRDiag)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellCreateTextCursor
; Description ...: Create a Text Cursor in a single Cell.
; Syntax ........: _LOCalc_CellCreateTextCursor(ByRef $oCell[, $bAtEnd = False])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
;                  $bAtEnd              - [optional] a boolean value. Default is False. If True, the Text Cursor is created at the end of the Text, else it will be created at the beginning.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell not a Single Cell Object. Only Single Cells supported.
;                  @Error 1 @Extended 3 Return 0 = $bAtEnd not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create Text Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully created a Text Cursor in the requested cell, returning the Text Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Cells that are considered Values instead of Strings may be considered Strings if you modify them using a text cursor.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellCreateTextCursor(ByRef $oCell, $bAtEnd = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.
	If Not IsBool($bAtEnd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextCursor = $oCell.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $bAtEnd Then
		$oTextCursor.gotoEnd(False)

	Else
		$oTextCursor.gotoStart(False)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOCalc_CellCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellEffect
; Description ...: Set or Retrieve the Font Effect settings for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellEffect(ByRef $oCell[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iRelief             - [optional] an integer value (0-2). Default is Null. The Character Relief style. See Constants, $LOC_RELIEF_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bOutline            - [optional] a boolean value. Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] a boolean value. Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants, $LOC_RELIEF_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bOutline not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bShadow not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRelief
;                  |                               2 = Error setting $bOutline
;                  |                               4 = Error setting $bShadow
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellStyleEffect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellEffect(ByRef $oCell, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellEffect($oCell, $iRelief, $bOutline, $bShadow)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellFont
; Description ...: Set and Retrieve the Font Settings for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellFont(ByRef $oCell[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sFontName           - [optional] a string value. Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] a general number value. Default is Null. The new Font size.
;                  $iPosture            - [optional] an integer value (0-5). Default is Null. The Font Italic setting. See Constants, $LOC_POSTURE_* as defined in LibreOfficeCalc_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] an integer value(0,50-200). Default is Null. The Font Bold settings see Constants, $LOC_WEIGHT_* as defined in LibreOfficeCalc_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Font called in $sFontName not available.
;                  @Error 1 @Extended 4 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 5 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 6 Return 0 = $nFontSize not a number.
;                  @Error 1 @Extended 7 Return 0 = $iPosture not an Integer, less than 0, or greater than 5. See Constants, $LOC_POSTURE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iWeight not an Integer, less than 50 but not equal to 0, or greater than 200. See Constants, $LOC_WEIGHT_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $nFontSize
;                  |                               4 = Error setting $iPosture
;                  |                               8 = Error setting $iWeight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;                  Libre Calc accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......: _LOCalc_FontsGetNames, _LOCalc_CellStyleFont
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellFont(ByRef $oCell, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	If ($sFontName <> Null) And Not _LOCalc_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$vReturn = __LOCalc_CellFont($oCell, $sFontName, $nFontSize, $iPosture, $iWeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellFontColor
; Description ...: Set or Retrieve the Font Color for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellFontColor(ByRef $oCell[, $iFontColor = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFontColor          - [optional] an integer value (-1-16777215). Default is Null. The Color value in Long Integer format to make the font, can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for Auto color.
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iFontColor not an Integer, less than 0, or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iFontColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current Font Color as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Though Transparency is present on the Font Effects page in the UI, there is (as best as I can find) no setting for it available to read and modify. And further, it seems even in L.O. the setting does not affect the font's transparency, though it may change the color value.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LO_ConvertColorToLong, _LO_ConvertColorFromLong, _LOCalc_CellStyleFontColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellFontColor(ByRef $oCell, $iFontColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellFontColor($oCell, $iFontColor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellFormula
; Description ...: Set or Retrieve a Cell's Formula.
; Syntax ........: _LOCalc_CellFormula(ByRef $oCell[, $sFormula = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
;                  $sFormula            - [optional] a string value. Default is Null. The Formula to set the Cell to. Overwrites any previous data.
; Return values .: Success: 1 or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;                  @Error 1 @Extended 3 Return 0 = $sFormula not a String.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFormula
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Cell's current formula as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only individual cells are supported, not cell ranges.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current Cell content.
; Related .......: _LOCalc_CellGetType, _LOCalc_CellString, _LOCalc_CellValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellFormula(ByRef $oCell, $sFormula = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

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
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Data Type.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning type of data contained in the Cell. Return value will be one of the constants, $LOC_CELL_TYPE_* as defined in LibreOfficeCalc_Constants.au3
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only individual cells are supported, not cell ranges.
; Related .......: _LOCalc_CellString, _LOCalc_CellFormula, _LOCalc_CellValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellGetType(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCellType

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	$iCellType = $oCell.getType()
	If Not IsInt($iCellType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCellType)
EndFunc   ;==>_LOCalc_CellGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellNumberFormat
; Description ...: Set or Retrieve Cell or Cell Range Number Format settings.
; Syntax ........: _LOCalc_CellNumberFormat(ByRef $oDoc, ByRef $oCell[, $iFormatKey = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFormatKey          - [optional] an integer value. Default is Null. A Format Key from a previous _LOCalc_FormatKeyCreate or _LOCalc_FormatKeysGetList function.
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 4 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 5 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 6 Return 0 = Format Key called in $iFormatKey not found in document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iFormatKey
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current setting as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellStyleNumberFormat, _LOCalc_FormatKeyCreate, _LOCalc_FormatKeysGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellNumberFormat(ByRef $oDoc, ByRef $oCell, $iFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Row Obj

	$vReturn = __LOCalc_CellNumberFormat($oDoc, $oCell, $iFormatKey)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellNumberFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellOverline
; Description ...: Set and retrieve the Overline settings for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellOverline(ByRef $oCell[, $bWordOnly = Null[, $iOverLineStyle = Null[, $bOLHasColor = Null[, $iOLColor = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not Overlined.
;                  $iOverLineStyle      - [optional] an integer value (0-18). Default is Null. The style of the Overline line, see constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3. See Remarks.
;                  $bOLHasColor         - [optional] a boolean value. Default is Null. If True, the Overline is colored, must be set to true in order to set the Overline color.
;                  $iOLColor            - [optional] an integer value (-1-16777215). Default is Null. The Overline color, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iOverLineStyle not an Integer, less than 0, or greater than 18. See constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3. See Remarks.
;                  @Error 1 @Extended 6 Return 0 = $bOLHasColor not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iOLColor not an Integer, less than -1, or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $iOverLineStyle
;                  |                               4 = Error setting $bOLHasColor
;                  |                               8 = Error setting $iOLColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Overline line style uses the same constants as underline style.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellUnderline, _LO_ConvertColorToLong, _LO_ConvertColorFromLong, _LOCalc_CellStyleOverline
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellOverline(ByRef $oCell, $bWordOnly = Null, $iOverLineStyle = Null, $bOLHasColor = Null, $iOLColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellOverLine($oCell, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellOverline

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellProtection
; Description ...: Set or Retrieve Cell or Cell Range protection settings.
; Syntax ........: _LOCalc_CellProtection(ByRef $oCell[, $bHideAll = Null[, $bProtected = Null[, $bHideFormula = Null[, $bHideWhenPrint = Null]]]])
; Parameters ....: $oCell               - [in/out] an object.
;                  $bHideAll            - [optional] a boolean value. Default is Null. If True, Hides formulas and contents of the cells in the range.
;                  $bProtected          - [optional] a boolean value. Default is Null. If True, Prevents the cells selected from being modified.
;                  $bHideFormula        - [optional] a boolean value. Default is Null. If True, Hides formulas in the cells in the selection.
;                  $bHideWhenPrint      - [optional] a boolean value. Default is Null. If True, cells in the selection are kept from being printed.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bHideAll not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bProtected not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bHideFormula not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bHideWhenPrint not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Protection Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bHideAll
;                  |                               2 = Error setting $bProtected
;                  |                               4 = Error setting $bHideFormula
;                  |                               8 = Error setting $bHideWhenPrint
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Cell protection only takes effect if you also protect the sheet. (Tools - Protect Sheet)
; Related .......: _LOCalc_CellStyleProtection
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellProtection(ByRef $oCell, $bHideAll = Null, $bProtected = Null, $bHideFormula = Null, $bHideWhenPrint = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellProtection($oCell, $bHideAll, $bProtected, $bHideFormula, $bHideWhenPrint)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellProtection

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellShadow
; Description ...: Set or Retrieve the Shadow settings for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellShadow(ByRef $oCell[, $iWidth = Null[, $iColor = Null[, $bTransparent = Null[, $iLocation = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iWidth              - [optional] an integer value (0-5009). Default is Null. The shadow width, set in Micrometers.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The color of the shadow, set in Long Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bTransparent        - [optional] a boolean value. Default is Null. If True, the shadow is transparent.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The location of the shadow compared to the Cell. See Constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iWidth not an Integer, less than 0, or greater than 5009.
;                  @Error 1 @Extended 5 Return 0 = $iColor not an Integer, less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 6 Return 0 = $bTransparent not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, less than 0, or greater than 4. See Constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shadow Format Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $bTransparent
;                  |                               8 = Error setting $iLocation
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  LibreOffice may change the shadow width +/- a Micrometer.
; Related .......: _LO_ConvertColorToLong, _LO_ConvertColorFromLong, _LOCalc_CellStyleShadow
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellShadow(ByRef $oCell, $iWidth = Null, $iColor = Null, $bTransparent = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellShadow($oCell, $iWidth, $iColor, $bTransparent, $iLocation)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellStrikeOut
; Description ...: Set or Retrieve the Strikeout settings for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellStrikeOut(ByRef $oCell[, $bWordOnly = Null[, $bStrikeOut = Null[, $iStrikeLineStyle = Null]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If True, strike out is applied to words only, skipping whitespaces.
;                  $bStrikeOut          - [optional] a boolean value. Default is Null. If True, strikeout is applied to characters.
;                  $iStrikeLineStyle    - [optional] an integer value (0-6). Default is Null. The Strikeout Line Style, see constants, $LOC_STRIKEOUT_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bStrikeOut not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iStrikeLineStyle not an Integer, less than 0 or greater than 6. See constants, $LOC_STRIKEOUT_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $bStrikeOut
;                  |                               4 = Error setting $iStrikeLineStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellStyleStrikeOut
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellStrikeOut(ByRef $oCell, $bWordOnly = Null, $bStrikeOut = Null, $iStrikeLineStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellStrikeOut($oCell, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellString
; Description ...: Set or Retrieve a Cell's Text content.
; Syntax ........: _LOCalc_CellString(ByRef $oCell[, $sText = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
;                  $sText               - [optional] a string value. Default is Null. The Text to set the Cell contents to. Overwrites any previous data.
; Return values .: Success: 1 or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;                  @Error 1 @Extended 3 Return 0 = $sText not a String.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sText
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Cell's contents as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only individual cells are supported, not cell ranges.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current Cell content.
; Related .......: _LOCalc_CellGetType, _LOCalc_CellFormula, _LOCalc_CellValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellString(ByRef $oCell, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	If ($sText = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oCell.getString())

	If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCell.setString($sText)
	If ($oCell.getString() <> $sText) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CellString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellTextAlign
; Description ...: Set and Retrieve text Alignment settings for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellTextAlign(ByRef $oCell[, $iHoriAlign = Null[, $iVertAlign = Null[, $iIndent = Null]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iHoriAlign          - [optional] an integer value (0-6). Default is Null. The Horizontal alignment of the text. See Constants, $LOC_CELL_ALIGN_HORI_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-5). Default is Null. The Vertical alignment of the text. See Constants, $LOC_CELL_ALIGN_VERT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iIndent             - [optional] an integer value. Default is Null. The amount of indentation from the left side of the cell, in micrometers.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iHoriAlign not an Integer, less than 0, or greater than 6. See Constants, $LOC_CELL_ALIGN_HORI_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iVertAlign not an Integer, less than 0, or greater than 5. See Constants, $LOC_CELL_ALIGN_VERT_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iIndent not an Integer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iHoriAlign
;                  |                               2 = Error setting $iVertAlign
;                  |                               4 = Error setting $iIndent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellTextOrient, _LOCalc_CellTextProperties, _LOCalc_CellStyleTextAlign
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellTextAlign(ByRef $oCell, $iHoriAlign = Null, $iVertAlign = Null, $iIndent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellTextAlign($oCell, $iHoriAlign, $iVertAlign, $iIndent)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellTextAlign

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellTextOrient
; Description ...: Set or Retrieve Text Orientation settings for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellTextOrient(ByRef $oCell[, $iRotate = Null[, $iReference = Null[, $bVerticalStack = Null[, $bAsianLayout = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iRotate             - [optional] an integer value (0-359). Default is Null. The rotation angle of the text in the cell.
;                  $iReference          - [optional] an integer value (0,1,3). Default is Null. The cell edge from which to write the rotated text. See Constants $LOC_CELL_ROTATE_REF_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bVerticalStack      - [optional] a boolean value. Default is Null. If True, Aligns text vertically. Only available after you enable support for Asian languages in Libre Office settings.
;                  $bAsianLayout        - [optional] a boolean value. Default is Null. If True, Aligns Asian characters one below the other. Only available after you enable support for Asian languages in Libre Office settings, and enable vertical text.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iRotate not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 5 Return 0 = $iReference not an Integer, less than 0, or greater than 1 but not equal to 3. See Constants $LOC_CELL_ROTATE_REF_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bVerticalStack not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bAsianLayout not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRotate
;                  |                               2 = Error setting $iReference
;                  |                               4 = Error setting $bVerticalStack
;                  |                               8 = Error setting $bAsianLayout
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellTextAlign, _LOCalc_CellTextProperties, _LOCalc_CellStyleTextOrient
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellTextOrient(ByRef $oCell, $iRotate = Null, $iReference = Null, $bVerticalStack = Null, $bAsianLayout = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellTextOrient($oCell, $iRotate, $iReference, $bVerticalStack, $bAsianLayout)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellTextOrient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellTextProperties
; Description ...: Set or Retrieve Text property settings for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellTextProperties(ByRef $oCell[, $bAutoWrapText = Null[, $bHyphen = Null[, $bShrinkToFit = Null[, $iTextDirection = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bAutoWrapText       - [optional] a boolean value. Default is Null. If True, Wraps text onto another line at the cell border.
;                  $bHyphen             - [optional] a boolean value. Default is Null. If True, Enables word hyphenation for text wrapping to the next line.
;                  $bShrinkToFit        - [optional] a boolean value. Default is Null. If True, Reduces the apparent size of the font so that the contents of the cell fit into the current cell width.
;                  $iTextDirection      - [optional] an integer value (0,1,4). Default is Null. The Text Writing Direction. See Constants, $LOC_TXT_DIR_* as defined in LibreOfficeCalc_Constants.au3. [Libre Office Default is 4]
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bAutoWrapText not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bHyphen not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bShrinkToFitnot a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iTextDirection not an Integer, less than 0, or greater than 1 but not equal to 4. See Constants, $LOC_TXT_DIR_* as defined in LibreOfficeCalc_Constants.au3. [Libre Office Default is 4]
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bAutoWrapText
;                  |                               2 = Error setting $bHyphen
;                  |                               4 = Error setting $bShrinkToFit
;                  |                               8 = Error setting $iTextDirection
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellTextAlign, _LOCalc_CellTextOrient, _LOCalc_CellStyleTextProperties
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellTextProperties(ByRef $oCell, $bAutoWrapText = Null, $bHyphen = Null, $bShrinkToFit = Null, $iTextDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellTextProperties($oCell, $bAutoWrapText, $bHyphen, $bShrinkToFit, $iTextDirection)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellTextProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellUnderline
; Description ...: Set and retrieve the Underline settings for a Cell or Cell Range.
; Syntax ........: _LOCalc_CellUnderline(ByRef $oCell[, $bWordOnly = Null[, $iUnderLineStyle = Null[, $bULHasColor = Null[, $iULColor = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not underlined.
;                  $iUnderLineStyle     - [optional] an integer value (0-18). Default is Null. The Underline line style, see constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bULHasColor         - [optional] a boolean value. Default is Null. If True, the underline is colored, must be set to true in order to set the underline color.
;                  $iULColor            - [optional] an integer value (-1-16777215). Default is Null. The color of the underline, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell does not support Character properties, or Table Column, or Table Row service.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iUnderLineStyle not an Integer, less than 0, or greater than 18. See constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3. See Remarks.
;                  @Error 1 @Extended 6 Return 0 = $bULHasColor not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iULColor not an Integer, less than -1, or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $iUnderLineStyle
;                  |                               4 = Error setting $bULHasColor
;                  |                               8 = Error setting $iULColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CellOverline, _LO_ConvertColorToLong, _LO_ConvertColorFromLong, _LOCalc_CellStyleUnderline
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellUnderline(ByRef $oCell, $bWordOnly = Null, $iUnderLineStyle = Null, $bULHasColor = Null, $iULColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.supportsService("com.sun.star.style.CharacterProperties") _
			And Not $oCell.supportsService("com.sun.star.table.TableColumn") _ ; Column Obj
			And Not $oCell.supportsService("com.sun.star.table.TableRow") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Row Obj

	$vReturn = __LOCalc_CellUnderLine($oCell, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_CellUnderline

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellValue
; Description ...: Set or Retrieve a Cell's Value.
; Syntax ........: _LOCalc_CellValue(ByRef $oCell[, $nValue = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
;                  $nValue              - [optional] a general number value. Default is Null. The Value to set the Cell to. Overwrites any previous data.
; Return values .: Success: 1 or Number.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;                  @Error 1 @Extended 3 Return 0 = $nValue not a Number.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $nValue
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Number = Success. All optional parameters were set to Null, returning the Cell's current number value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only individual cells are supported, not cell ranges.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current Cell content.
; Related .......: _LOCalc_CellGetType, _LOCalc_CellString, _LOCalc_CellFormula
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellValue(ByRef $oCell, $nValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	If ($nValue = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oCell.getValue())

	If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCell.setValue($nValue)
	If ($oCell.getValue() <> $nValue) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CellValue
