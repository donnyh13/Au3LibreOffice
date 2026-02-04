#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; Other includes for Writer
#include "LibreOfficeWriter_Page.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Inserting Tables in L.O. Writer.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_TableBackColor
; _LOWriter_TableBorderColor
; _LOWriter_TableBorderPadding
; _LOWriter_TableBorderStyle
; _LOWriter_TableBorderWidth
; _LOWriter_TableBreak
; _LOWriter_TableCellsGetNames
; _LOWriter_TableColumnDelete
; _LOWriter_TableColumnGetCount
; _LOWriter_TableColumnInsert
; _LOWriter_TableCreate
; _LOWriter_TableCreateCursor
; _LOWriter_TableCursor
; _LOWriter_TableDelete
; _LOWriter_TableExists
; _LOWriter_TableGetCellObjByCursor
; _LOWriter_TableGetCellObjByName
; _LOWriter_TableGetCellObjByPosition
; _LOWriter_TableGetData
; _LOWriter_TableGetObjByCursor
; _LOWriter_TableGetObjByName
; _LOWriter_TableMargin
; _LOWriter_TableProperties
; _LOWriter_TableRowBackColor
; _LOWriter_TableRowDelete
; _LOWriter_TableRowGetCount
; _LOWriter_TableRowInsert
; _LOWriter_TableRowProperty
; _LOWriter_TableSetData
; _LOWriter_TablesGetNames
; _LOWriter_TableShadow
; _LOWriter_TableStyleCurrent
; _LOWriter_TableStyleExists
; _LOWriter_TableStylesGetNames
; _LOWriter_TableWidth
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableBackColor
; Description ...: Set or Retrieve the Background color of a Table.
; Syntax ........: _LOWriter_TableBackColor(ByRef $oTable[, $iBackColor = Null])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Table background color, as a RGB Color Integer. See Remarks. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for no background color.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Background color.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableBackColor(ByRef $oTable, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iColor

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iBackColor) Then
		$iColor = $oTable.BackColor()
		If Not IsInt($iColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTable.BackColor = $iBackColor
	$iError = ($oTable.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1)) ; Error setting color.

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_TableBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableBorderColor
; Description ...: Set or Retrieve the Table Border Line Color. Libre Office Version 3.6 and Up.
; Syntax ........: _LOWriter_TableBorderColor(ByRef $oTable[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $iVert = Null[, $iHori = Null]]]]]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. The Top Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. The Bottom Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. The Left Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. The Right Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iVert               - [optional] an integer value (0-16777215). Default is Null. The Vertical Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iHori               - [optional] an integer value (0-16777215). Default is Null. The Horizontal Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTop not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 3 Return 0 = $iBottom not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $iLeft not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iRight not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 6 Return 0 = $iVert not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 7 Return 0 = $iHori not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Object "TableBorder2".
;                  @Error 3 @Extended 3 Return 0 = Cannot set Top Border Color when Top Border width not set.
;                  @Error 3 @Extended 4 Return 0 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error 3 @Extended 5 Return 0 = Cannot set Left Border Color when Left Border width not set.
;                  @Error 3 @Extended 6 Return 0 = Cannot set Right Border Color when Right Border width not set.
;                  @Error 3 @Extended 7 Return 0 = Cannot set Vertical Border Color when Vertical Border width not set.
;                  @Error 3 @Extended 8 Return 0 = Cannot set Horizontal Border Color when Horizontal Border width not set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  |                               16 = Error setting $iVert
;                  |                               32 = Error setting $iHori
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_TableBorderWidth, _LOWriter_TableBorderStyle, _LOWriter_TableBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableBorderColor(ByRef $oTable, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $iVert = Null, $iHori = Null)
	Local $vReturn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iVert <> Null) And Not __LO_IntIsBetween($iVert, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($iHori <> Null) And Not __LO_IntIsBetween($iHori, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$vReturn = __LOWriter_TableBorder($oTable, False, False, True, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_TableBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableBorderPadding
; Description ...: Set or retrieve the Border Padding (spacing between the Table text and border) settings.
; Syntax ........: _LOWriter_TableBorderPadding(ByRef $oTable[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iTop                - [optional] an integer value. Default is Null. The Top Distance between the Border and Table contents in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Distance between the Border and Table contents in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Distance between the Border and Table contents in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The Right Distance between the Border and Table contents in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iBottom not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $Left not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iRight not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving TableBorderDistances Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving TableBorderDistances Object for Error checking.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTop border distance
;                  |                               2 = Error setting $iBottom border distance
;                  |                               4 = Error setting $iLeft border distance
;                  |                               8 = Error setting $iRight border distance
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LO_UnitConvert, _LOWriter_TableBorderWidth, _LOWriter_TableBorderStyle, _LOWriter_TableBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableBorderPadding(ByRef $oTable, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tBD
	Local $aiBPadding[4]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oTable.TableBorderDistances.TopDistance(), $oTable.TableBorderDistances.BottomDistance(), _
				$oTable.TableBorderDistances.LeftDistance(), $oTable.TableBorderDistances.RightDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	$tBD = $oTable.TableBorderDistances()
	If Not IsObj($tBD) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tBD.TopDistance = $iTop
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tBD.BottomDistance = $iBottom
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tBD.LeftDistance = $iLeft
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tBD.RightDistance = $iRight
	EndIf

	$oTable.TableBorderDistances = $tBD
	; Error Checking.
	$tBD = $oTable.TableBorderDistances()
	If Not IsObj($tBD) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iError = (__LO_VarsAreNull($iTop)) ? ($iError) : ((__LO_IntIsBetween($tBD.TopDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iBottom)) ? ($iError) : ((__LO_IntIsBetween($tBD.BottomDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iLeft)) ? ($iError) : ((__LO_IntIsBetween($tBD.LeftDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iRight)) ? ($iError) : ((__LO_IntIsBetween($tBD.RightDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_TableBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableBorderStyle
; Description ...: Set or Retrieve the Table Border Line style. Libre Office Version 3.6 and Up.
; Syntax ........: _LOWriter_TableBorderStyle(ByRef $oTable[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $iVert = Null[, $iHori = Null]]]]]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. The Top Border Line Style of the Table. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. The Bottom Border Line Style of the Table. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. The Left Border Line Style of the Table. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. The Right Border Line Style of the Table. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVert               - [optional] an integer value (0x7FFF,0-17). Default is Null. The internal Vertical Border Line Styles of the Table. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iHori               - [optional] an integer value (0x7FFF,0-17). Default is Null. The internal Horizontal Border Line Styles of the Table. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 3 Return 0 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 4 Return 0 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 5 Return 0 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 6 Return 0 = $iVert not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 7 Return 0 = $iHori not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Object "TableBorder2".
;                  @Error 3 @Extended 3 Return 0 = Cannot set Top Border Style when Top Border width not set.
;                  @Error 3 @Extended 4 Return 0 = Cannot set Bottom Border Style when Bottom Border width not set.
;                  @Error 3 @Extended 5 Return 0 = Cannot set Left Border Style when Left Border width not set.
;                  @Error 3 @Extended 6 Return 0 = Cannot set Right Border Style when Right Border width not set.
;                  @Error 3 @Extended 7 Return 0 = Cannot set Vertical Border Style when Vertical Border width not set.
;                  @Error 3 @Extended 8 Return 0 = Cannot set Horizontal Border Style when Horizontal Border width not set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  |                               16 = Error setting $iVert
;                  |                               32 = Error setting $iHori
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableBorderWidth, _LOWriter_TableBorderColor, _LOWriter_TableBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableBorderStyle(ByRef $oTable, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $iVert = Null, $iHori = Null)
	Local $vReturn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iVert <> Null) And Not __LO_IntIsBetween($iVert, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($iHori <> Null) And Not __LO_IntIsBetween($iHori, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$vReturn = __LOWriter_TableBorder($oTable, False, True, False, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_TableBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableBorderWidth
; Description ...: Set or Retrieve the Table Border Line Width. Libre Office Version 3.6 and Up.
; Syntax ........: _LOWriter_TableBorderWidth(ByRef $oTable[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $iVert = Null[, $iHori = Null]]]]]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iTop                - [optional] an integer value. Default is Null. The Top Border Line width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Border Line Width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Border Line width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. The Right Border Line Width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVert               - [optional] an integer value. Default is Null. The Internal Vertical Border Line width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iHori               - [optional] an integer value. Default is Null. The Internal Horizontal Border Line width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTop not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = $iBottom not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iLeft not an Integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iRight not an Integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iVert not an Integer, or less than 0.
;                  @Error 1 @Extended 7 Return 0 = $iHori not an Integer, or less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Object "TableBorder2".
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  |                               16 = Error setting $iVert
;                  |                               32 = Error setting $iHori
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set them to 0
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LO_UnitConvert, _LOWriter_TableBorderStyle, _LOWriter_TableBorderColor, _LOWriter_TableBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableBorderWidth(ByRef $oTable, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $iVert = Null, $iHori = Null)
	Local $vReturn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iVert <> Null) And Not __LO_IntIsBetween($iVert, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($iHori <> Null) And Not __LO_IntIsBetween($iHori, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$vReturn = __LOWriter_TableBorder($oTable, True, False, False, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_TableBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableBreak
; Description ...: Set or retrieve the Paragraph break settings for before or after the Table.
; Syntax ........: _LOWriter_TableBreak(ByRef $oDoc, ByRef $oTable[, $iBreakType = Null[, $sPageStyle = Null[, $iPgNumOffSet = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iBreakType          - [optional] an integer value (0-6). Default is Null. The Type of break to insert, see constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPageStyle          - [optional] a string value. Default is Null. The New Page Style to begin with after the paragraph break. If Set, to remove the break you must set this to "".
;                  $iPgNumOffSet        - [optional] an integer value. Default is Null. If a page break property is set at the table, this property contains the new value for the page number.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTable not an Object
;                  @Error 1 @Extended 3 Return 0 = $iBreakType not an Integer, less than 0 or greater than 6. See Constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $sPageStyle not a String.
;                  @Error 1 @Extended 5 Return 0 = $sPageStyle not found in current document.
;                  @Error 1 @Extended 6 Return 0 = $iPgNumOffSet not an Integer, or less than 0.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBreakType
;                  |                               2 = Error setting $sPageStyle
;                  |                               4 = Error setting $iPgNumOffSet
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Break Type must be set before Page Style will be able to be set, and page style needs set before $iPgNumOffSet can be set.
;                  LibreOffice doesn't directly show in its User interface options for Break type constants #3 and #6 (Column both) and (Page both), but doesn't throw an error when being set to either one, so they are included here, though I'm not sure if they will work correctly.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_PageStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableBreak(ByRef $oDoc, ByRef $oTable, $iBreakType = Null, $sPageStyle = Null, $iPgNumOffSet = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avBreaks[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iBreakType, $sPageStyle, $iPgNumOffSet) Then
		__LO_ArrayFill($avBreaks, $oTable.BreakType(), $oTable.PageDescName(), $oTable.PageNumberOffset())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avBreaks)
	EndIf

	If ($iBreakType <> Null) Then
		If Not __LO_IntIsBetween($iBreakType, 0, 6) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTable.BreakType = $iBreakType
		$iError = ($oTable.BreakType() = $iBreakType) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sPageStyle <> Null) Then
		If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If ($sPageStyle <> "") And Not _LOWriter_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTable.PageDescName = $sPageStyle
		$iError = (__LOWriter_PageStyleCompare($oDoc, $oTable.PageDescName(), $sPageStyle)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iPgNumOffSet <> Null) Then
		If Not __LO_IntIsBetween($iPgNumOffSet, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTable.PageNumberOffset = $iPgNumOffSet
		$iError = ($oTable.PageNumberOffset() = $iPgNumOffSet) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) ; error setting Properties.
EndFunc   ;==>_LOWriter_TableBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableCellsGetNames
; Description ...: Retrieve an array of all Cell names from a Table.
; Syntax ........: _LOWriter_TableCellsGetNames(ByRef $oTable)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Cell Names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Array of Cell names. @Extended set to number of names returned in the array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableCellsGetNames(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asCellNames

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0) ; Not an Object.

	$asCellNames = $oTable.getCellNames()
	If Not IsArray($asCellNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; failed to get array of names.

	Return SetError($__LO_STATUS_SUCCESS, UBound($asCellNames), $asCellNames)
EndFunc   ;==>_LOWriter_TableCellsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableColumnDelete
; Description ...: Delete a column from a Text Table.
; Syntax ........: _LOWriter_TableColumnDelete(ByRef $oTable, $iColumn[, $iCount = 1])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iColumn             - an integer value. The Column to delete.
;                  $iCount              - [optional] an integer value. Default is 1. Number of columns to delete starting at the column called in $iColumn and moving right.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;                  @Error 1 @Extended 4 Return 0 = Requested column greater than number of columns contained in table.
;                  --Success--
;                  @Error 0 @Extended ? Return 1 = Full amount of columns deleted. @Extended set to total columns deleted.
;                  @Error 0 @Extended ? Return 2 = $iCount greater than amount of columns contained in Table; deleted all columns from $iColumn over. @Extended set to total columns deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice counts columns and Rows starting at 0. So to delete the first column in a Table you would set $iColumn to 0.
;                  If you attempt to delete more columns than are present all columns from $iColumn over will be deleted.
;                  If you delete all columns starting from column 0, the entire Table is deleted.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableColumnGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableColumnDelete(ByRef $oTable, $iColumn, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iColumnCount, $iReturn = 0

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iColumn) Or ($iColumn < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iCount) Or ($iCount < 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iColumnCount = $oTable.getColumns.getCount()
	If ($iColumnCount <= $iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Requested column out of bounds.

	$iCount = ($iCount > ($iColumnCount - $iColumn)) ? ($iColumnCount - $iColumn) : ($iCount)
	$iReturn = ($iCount > ($iColumnCount - $iColumn)) ? (2) : (1) ; Return 1 if full amount deleted else 2 if only partial.
	$oTable.getColumns.removeByIndex($iColumn, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $iReturn)
EndFunc   ;==>_LOWriter_TableColumnDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableColumnGetCount
; Description ...: Retrieves the number of Columns in a table.
; Syntax ........: _LOWriter_TableColumnGetCount(ByRef $oTable)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Column count.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Returning Column Count as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableColumnGetCount(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iColumnSize = 0

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0) ; Not an Object.

	$iColumnSize = $oTable.getColumns.getCount()
	If ($iColumnSize = 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Failed to retrieve column count.

	Return SetError($__LO_STATUS_SUCCESS, 0, $iColumnSize)
EndFunc   ;==>_LOWriter_TableColumnGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableColumnInsert
; Description ...: Insert columns into a Text Table
; Syntax ........: _LOWriter_TableColumnInsert(ByRef $oTable, $iCount[, $iColumn = -1])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iCount              - an integer value. Number of columns to insert.
;                  $iColumn             - [optional] an integer value. Default is -1. The column to insert columns after. See Remarks.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iCount not an Integer, or less than 1.
;                  @Error 1 @Extended 3 Return 0 = $iColumn not an Integer, or less than -1.
;                  @Error 1 @Extended 4 Return 0 = Column called in $iColumn greater than number of columns contained in table.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to insert columns.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Successfully inserted the number of desired columns.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you do not set $iColumn, the new columns will be placed at the end of the Table.
;                  LibreOffice counts the Table columns/Rows starting at 0. The columns are placed behind the desired column when inserted.
;                  To insert a column at the left most of the Table you would set $iColumn to 0. To insert columns at the Right of a table you would set $iColumn to one higher than the last column. e.g. a Table containing 3 columns, would be numbered as follows: 0(first-Column), 1(second-Column), 2(third-Column), to insert columns at the very Right of the columns, you would set $iColumn to 3.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableColumnGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableColumnInsert(ByRef $oTable, $iCount, $iColumn = -1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iColumnCount

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iCount, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iColumn, -1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iColumnCount = $oTable.getColumns.getCount()
	If ($iColumnCount < $iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Requested column out of bounds.

	$iColumn = ($iColumn <= -1) ? ($iColumnCount) : ($iColumn)
	$oTable.getColumns.insertByIndex($iColumn, $iCount)

	Return ($oTable.getColumns.getCount() = ($iColumnCount + $iCount)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0))
EndFunc   ;==>_LOWriter_TableColumnInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableCreate
; Description ...: Create and insert a Text Table into a document.
; Syntax ........: _LOWriter_TableCreate(ByRef $oDoc, ByRef $oCursor[, $iColumns = 2[, $iRows = 3[, $iBackColor = Null[, $sTableName = Null[, $bHeading = Null[, $sStyle = Null[, $bSplit = Null[, $bRepeatHeading = Null[, $iHeadingRows = Null]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions. See Remarks.
;                  $iColumns            - [optional] an integer value. Default is 2. The number of columns to create the table with.
;                  $iRows               - [optional] an integer value. Default is 3. The number of rows to create the table with.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Table background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF (-1) for no background color.
;                  $sTableName          - [optional] a string value. Default is Null. The unique table name. If set to Null, the Table is automatically named.
;                  $bHeading            - [optional] a boolean value. Default is Null. If True, the first row of a Table is a Heading row.
;                  $sStyle              - [optional] a string value. Default is Null. The TableStyle to apply to the Table.
;                  $bSplit              - [optional] a boolean value. Default is Null. If True, the table is allowed to split across two pages.
;                  $bRepeatHeading      - [optional] a boolean value. Default is Null. If True, the Heading is repeated on each subsequent page.
;                  $iHeadingRows        - [optional] an integer value. Default is Null. The number of rows to count as a Heading. See remarks.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iColumns not an Integer, or less than 1.
;                  @Error 1 @Extended 4 Return 0 = $iRows not an Integer, or less than 1.
;                  @Error 1 @Extended 5 Return 0 = $oCursor Object located in a Foot/EndNote.
;                  @Error 1 @Extended 6 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 7 Return 0 = $sTableName not a String.
;                  @Error 1 @Extended 8 Return 0 = Name called in $sTableName already exists.
;                  @Error 1 @Extended 9 Return 0 = $sStyle not a String.
;                  @Error 1 @Extended 10 Return 0 = Table Style called in $sStyle not found.
;                  @Error 1 @Extended 11 Return 0 = $bSplit not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $bRepeatHeading not a Boolean.
;                  @Error 1 @Extended 13 Return 0 = $iHeadingRows not an Integer, or less than 1.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.text.TextTable" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error getting Text Object from Cursor.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iBackColor
;                  |                               2 = Error setting $sTableName
;                  |                               4 = Error setting $bHeading
;                  |                               8 = Error setting $sStyle
;                  |                               16 = Error setting $bSplit
;                  |                               32 = Error setting $bRepeatHeading
;                  |                               64 = Error setting $iHeadingRows
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Table was created successfully and inserted at cursor position. Returning Table Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;                  Text Tables cannot be inserted into Foot/Endnotes. And it is not best to place them into other tables, though it is possible.
;                  You can set the $oCursor parameter to either a ViewCursor or a Text cursor in an acceptable data type, the table will be inserted at the cursor position.
;                  $iHeadingRows accepts values from 1 to 1 less then the number of rows.
;                  If a property setting error occurs, the table will have still been successfully inserted, and the Table's object will still be returned.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableCreate(ByRef $oDoc, ByRef $oCursor, $iColumns = 2, $iRows = 3, $iBackColor = Null, $sTableName = Null, $bHeading = Null, $sStyle = Null, $bSplit = Null, $bRepeatHeading = Null, $iHeadingRows = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTable, $oText, $oTextCursor
	Local $iCursorDataType, $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) And ($oCursor <> Default) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iColumns, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iRows, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oText = __LOWriter_CursorGetText($oDoc, $oCursor)
	$iCursorDataType = @extended
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorDataType = $LOW_CURDATA_FOOTNOTE) Or ($iCursorDataType = $LOW_CURDATA_ENDNOTE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; Unable to insert tables in footnotes/ EndNotes

	$oTable = $oDoc.createInstance("com.sun.star.text.TextTable")
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTable.initialize($iRows, $iColumns)

	If ($iBackColor <> Null) Then ; Set color before inserting the Table for cleaner appearance.
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTable.BackColor = $iBackColor
		$iError = ($oTable.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sTableName <> Null) Then
		If Not IsString($sTableName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If _LOWriter_TableExists($oDoc, $sTableName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	EndIf

	$oText.insertTextContent($oCursor, $oTable, False)

	If ($sTableName <> Null) Then ; Set name after inserting, otherwise the name changes.
		$oTable.setName($sTableName)
		$iError = ($oTable.Name() = $sTableName) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If $bHeading Then
		$oTable.RepeatHeadline = True
		$oTable.HeaderRowCount = 1
		For $i = 0 To $oTable.getColumns.getCount() - 1
			$oTextCursor = $oTable.getCellByPosition($i, 0).Text.createTextCursor()
			_LOWriter_ParStyleCurrent($oDoc, $oTextCursor, "Table Heading")
			If @error Then
				$iError = BitOR($iError, 4)
				ExitLoop
			EndIf
		Next

	Else
		$oTable.RepeatHeadline = False
		$oTable.HeaderRowCount = 0
		For $i = 0 To $oTable.getColumns.getCount() - 1
			$oTextCursor = $oTable.getCellByPosition($i, 0).Text.createTextCursor()
			_LOWriter_ParStyleCurrent($oDoc, $oTextCursor, "Table Contents")
			If @error Then
				$iError = BitOR($iError, 4)
				ExitLoop
			EndIf
		Next
	EndIf

	If ($sStyle <> Null) Then
		If Not IsString($sStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not _LOWriter_TableStyleExists($oDoc, $sStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oTable.TableTemplateName = $sStyle
		$iError = (__LOWriter_TableStyleCompare($oDoc, $oTable.TableTemplateName(), $sStyle)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bSplit <> Null) Then
		If Not IsBool($bSplit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oTable.Split = $bSplit
		$iError = ($oTable.Split() = $bSplit) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bRepeatHeading <> Null) Then
		If Not IsBool($bRepeatHeading) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oTable.RepeatHeadline = $bRepeatHeading
		$iError = ($oTable.RepeatHeadline() = $bRepeatHeading) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iHeadingRows <> Null) Then
		If ($iHeadingRows <> 1) And Not __LO_IntIsBetween($iHeadingRows, 1, $iRows - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oTable.HeaderRowCount = $iHeadingRows
		$iError = ($oTable.HeaderRowCount() = $iHeadingRows) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, $oTable)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oTable)) ; error setting Properties.
EndFunc   ;==>_LOWriter_TableCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableCreateCursor
; Description ...: Create a Table Cursor for modifying Text-Table properties.
; Syntax ........: _LOWriter_TableCreateCursor(ByRef $oDoc, $oTable[, $sCellName = ""[, $oCursor = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oTable              - an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function. See remarks.
;                  $sCellName           - [optional] a string value. Default is "". The Table Cell name to create a Text Table Cursor in. See Remarks.
;                  $oCursor             - [optional] an object. Default is Null. A Cursor Object returned from any Cursor Object creation or retrieval functions. See Remarks.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTable and $oCursor not Objects.
;                  @Error 1 @Extended 3 Return 0 = $oTable and $oCursor both Objects.
;                  @Error 1 @Extended 4 Return 0 = $sCellName not a String.
;                  @Error 1 @Extended 5 Return 0 = $oCursor not in a Table.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Cursor by Cell Name or by first Cell name in Table.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failure retrieving Table by Table Name from Cursor.
;                  @Error 3 @Extended 2 Return 0 = Failure retrieving list of Table Cell Names.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success, TableCursor object was created successfully. Returning Table Cursor Object for further Table manipulation functions.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oTable can be either called with a Table object, or Null Keyword with $oCursor called with a Cursor object, $oCursor can be either called with a cursor object currently located in a Table (such as a ViewCursor) or a TextCursor located in a table.
;                  $sCellName can be left blank, which will place the TextTableCursor at the first cell (Typically "A1") if $oTable is called with an Object, else if $oCursor is used, the cell the cursor is currently located in is used.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableCellsGetNames, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableCreateCursor(ByRef $oDoc, $oTable, $sCellName = "", $oCursor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTableCursor
	Local $asCells

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTable) And Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If IsObj($oTable) And IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If IsObj($oCursor) Then
		Switch __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor)
			Case $LOW_CURDATA_CELL ; Transform to TextTableCursor
				$oTable = $oDoc.TextTables.getByName($oCursor.TextTable.Name)
				If Not IsObj($oTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				$sCellName = ($sCellName = "") ? ($oCursor.Cell.CellName) : ($sCellName)

			Case Else

				Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; Wrong Cursor Data Type
		EndSwitch
	EndIf

	If ($sCellName = "") Then ; If cell name undefined, get first cell.
		$asCells = $oTable.getCellNames()
		If Not IsArray($asCells) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; no cells

		$sCellName = $asCells[0]
	EndIf

	$oTableCursor = $oTable.createCursorByCellName($sCellName)
	If Not IsObj($oTableCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTableCursor)
EndFunc   ;==>_LOWriter_TableCreateCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableCursor
; Description ...: Commands related to a Table Cursor.
; Syntax ........: _LOWriter_TableCursor(ByRef $oCursor[, $sGoToCellByName = Null[, $bSelect = False[, $bMergeRange = Null[, $iSplitRangeInto = Null[, $bSplitRangeHori = False]]]]])
; Parameters ....: $oCursor             - [in/out] an object. A Table Cursor Object returned from a _LOWriter_TableCreateCursor function.
;                  $sGoToCellByName     - [optional] a string value. Default is Null. Move the cursor to the cell with the specified name, Case Sensitive; See also $bSelect.
;                  $bSelect             - [optional] a boolean value. Default is False. If True, selection is expanded when moving to a specific cell with $sGoToCellByName.
;                  $bMergeRange         - [optional] a boolean value. Default is Null. Merge the selected range of cells.
;                  $iSplitRangeInto     - [optional] an integer value. Default is Null. Create n new cells in each cell selected by the cursor. See also $bSplitRangeHori.
;                  $bSplitRangeHori     - [optional] a boolean value. Default is False. If True, splits the selected cell or cell range horizontally, else, False for vertically.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not a Table Cursor.
;                  @Error 1 @Extended 3 Return 0 = $sGoToCellByName not a String.
;                  @Error 1 @Extended 4 Return 0 = $bSelect not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iSplitRangeInto not an Integer, or less than 1.
;                  @Error 1 @Extended 6 Return 0 = $bSplitRangeHori not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended ? Return 0 = Some commands were not successfully completed. Use BitAND to test @Extended for the following values:
;                  |                                1 = Failed while processing $sGoToCellByName.
;                  |                                2 = Failed while processing $bMergeRange.
;                  |                                4 = Failed while processing $iSplitRangeInto.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Command was successfully completed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableCreateCursor, _LOWriter_CursorMove, _LOWriter_TableCellsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableCursor(ByRef $oCursor, $sGoToCellByName = Null, $bSelect = False, $bMergeRange = Null, $iSplitRangeInto = Null, $bSplitRangeHori = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $iError = 0

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ((__LOWriter_Internal_CursorGetType($oCursor)) <> $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($sGoToCellByName <> Null) Then
		If Not IsString($sGoToCellByName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$vReturn = $oCursor.gotoCellByName($sGoToCellByName, $bSelect)
		$iError = ($vReturn = True) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bMergeRange = True) Then
		$vReturn = $oCursor.mergeRange()
		$iError = ($vReturn = True) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iSplitRangeInto <> Null) Then
		If Not __LO_IntIsBetween($iSplitRangeInto, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not IsBool($bSplitRangeHori) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$vReturn = $oCursor.splitRange($iSplitRangeInto, $bSplitRangeHori)
		$iError = ($vReturn = True) ? ($iError) : (BitOR($iError, 4, 0))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROCESSING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_TableCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableDelete
; Description ...: Delete a table from the document.
; Syntax ........: _LOWriter_TableDelete(ByRef $oDoc, ByRef $oTable)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Table name.
;                  @Error 3 @Extended 2 Return 0 = Failed to delete Table.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Table was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableDelete(ByRef $oDoc, ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sTableName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sTableName = $oTable.getName()
	If Not IsString($sTableName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTable.dispose()
	If ($oDoc.TextTables.hasByName($sTableName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Document still contains Table named the same.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_TableDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableExists
; Description ...: Check if a Document contains a Table by name.
; Syntax ........: _LOWriter_TableExists(ByRef $oDoc, $sTableName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
;                  $sTableName          - a string value. The Table name to search for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sTableName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Text Tables Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If a table was found matching $sTableName, True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableExists(ByRef $oDoc, $sTableName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTables

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sTableName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTables = $oDoc.TextTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oTables.hasByName($sTableName)) Then Return SetError($__LO_STATUS_SUCCESS, 1, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; No matches
EndFunc   ;==>_LOWriter_TableExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableGetCellObjByCursor
; Description ...: Retrieve a single Cell Object or a Cell Range by Cursor.
; Syntax ........: _LOWriter_TableGetCellObjByCursor(ByRef $oDoc, ByRef $oTable, ByRef $oCursor)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 4 Return 0 = $oCursor is not currently located inside of a Table Cell.
;                  @Error 1 @Extended 5 Return 0 = $oCursor unknown cursor type.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failure Retrieving Cell Object
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = A Cell object or a Cell Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will accept a Table Cursor, a ViewCursor, or a Text Cursor.
;                  A TableCursor and ViewCursor can retrieve the single cell they are located in, or a range of cells that have been selected by them.
;                  A TextCursor can only retrieve the single cell it is located in.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableGetCellObjByCursor(ByRef $oDoc, ByRef $oTable, ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType, $iCursorDataType
	Local $oCell, $oSelection
	Local $sCellRange

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)

	Switch $iCursorType
		Case $LOW_CURTYPE_TABLE_CURSOR
			$sCellRange = $oCursor.getRangeName()
			$oCell = (StringInStr($sCellRange, ":")) ? ($oTable.getCellRangeByName($sCellRange)) : ($oTable.getCellByName($sCellRange))

		Case $LOW_CURTYPE_TEXT_CURSOR
			$iCursorDataType = __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor)
			If Not ($iCursorDataType = $LOW_CURDATA_CELL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Cursor not in a Table cell.

			$oCell = $oTable.getCellByName($oCursor.Cell.CellName)

		Case $LOW_CURTYPE_VIEW_CURSOR
			$oSelection = $oDoc.CurrentSelection()
			If ($oSelection.ImplementationName() = "SwXTextTableCursor") Then
				$oCell = $oTable.getCellRangeByName($oSelection.getRangeName())

			Else
				$iCursorDataType = __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor)
				If Not ($iCursorDataType = $LOW_CURDATA_CELL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Cursor not in a Table cell.

				$oCell = $oTable.getCellByName($oCursor.Cell.CellName)
			EndIf

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; Unknown cursor type.
	EndSwitch

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCell)
EndFunc   ;==>_LOWriter_TableGetCellObjByCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableGetCellObjByName
; Description ...: Retrieve a Cell Object or a Cell range by name.
; Syntax ........: _LOWriter_TableGetCellObjByName(ByRef $oTable, $sCellName[, $sToCellName = $sCellName])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $sCellName           - a string value. A Cell Name. Note: Case Sensitive. See remarks.
;                  $sToCellName         - [optional] a string value. Default is $sCellName. The Cell name to end the Cell Range. Note: Case Sensitive.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable is not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sCellName not a String.
;                  @Error 1 @Extended 3 Return 0 = $sToCellName not a String.
;                  @Error 1 @Extended 4 Return 0 = Table does not contain the Requested Cell name as called in $sCellName.
;                  @Error 1 @Extended 5 Return 0 = Table does not contain the Requested Cell name as called in $sToCellName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. A Cell object or a Cell Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Cell names are Case Sensitive. LibreOffice first goes from A to Z, and then a to z and then AA to ZZ etc.
;                  $sCellName can contain a Cell name that is located after $sToCellName in the Table.
;                  If $sToCellName is left blank, a cell object is returned instead of a Cell Range.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableCellsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableGetCellObjByName(ByRef $oTable, $sCellName, $sToCellName = $sCellName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPALL = 8
	Local $oCell

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sToCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$sCellName = StringStripWS($sCellName, $__STR_STRIPALL)
	$sToCellName = StringStripWS($sToCellName, $__STR_STRIPALL)
	If Not __LOWriter_TableHasCellName($oTable, $sCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; CellName not contained in Table
	If Not __LOWriter_TableHasCellName($oTable, $sToCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; ToCellName not contained in Table

	$oCell = ($sCellName = $sToCellName) ? ($oTable.getCellByName($sCellName)) : ($oTable.getCellRangeByName($sCellName & ":" & $sToCellName))
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCell)
EndFunc   ;==>_LOWriter_TableGetCellObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableGetCellObjByPosition
; Description ...: Retrieve a Cell object or Cell Range by position. See Remarks
; Syntax ........: _LOWriter_TableGetCellObjByPosition(ByRef $oTable, $iColumn, $iRow[, $iToColumn = Null[, $iToRow = Null]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iColumn             - an integer value. The column the desired cell is located in, or where to start the the cell range from.
;                  $iRow                - an integer value. The row the desired cell is located in, or where to start the the cell range from.
;                  $iToColumn           - [optional] an integer value. Default is Null. The column containing the cell where to end the the cell range. Can be the same as $iRow or higher. If left blank $iToColumn will be the same as $iColumn.
;                  $iToRow              - [optional] an integer value. Default is Null. The row containing the cell where to end the the cell range. Can be the same as $iRow or higher. If left blank $iToRow will be the same as $iRow.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, or less than 0 or greater than number of Columns contained in the table.
;                  @Error 1 @Extended 3 Return 0 = $iRow not an Integer, or less than 0 or greater than number of Rows contained in the table.
;                  @Error 1 @Extended 4 Return 0 = $iToColumn not an Integer, or less than $iColumn or greater than number of Columns contained in the table.
;                  @Error 1 @Extended 5 Return 0 = $iToRow not an Integer, or less than $iRow or greater than number of Rows contained in the table.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. A Cell object or a Cell Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function can fail with complex Tables. Complex tables are tables that contain cells that have been split or joined.
;                  Rows and Columns in a Table are 0 based, meaning they start their count at 0. The first cell is column 0 row 0.
;                  To retrieve a single cell, only call the $iColumn and $iRow parameters.
;                  To retrieve a cell range, call $iColumn with the lowest Integer value column and then $iToColumn with the highest Integer value column. As also for $iRow and $iToRow.
;                  You may request the same row in both $iRow and $iToRow, but neither $iToRow or $iToColumn may be a lower Integer value than $iRow and $iColumn respectively.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableColumnGetCount, _LOWriter_TableRowGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableGetCellObjByPosition(ByRef $oTable, $iColumn, $iRow, $iToColumn = Null, $iToRow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iColumn, 0, ($oTable.getColumns.getCount() - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iRow, 0, ($oTable.getRows.getCount() - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iToColumn, $iToRow) Then
		$oCell = $oTable.getCellByPosition($iColumn, $iRow)

	Else
		If Not __LO_IntIsBetween($iToColumn, $iColumn, ($oTable.getColumns.getCount() - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If Not __LO_IntIsBetween($iToRow, $iRow, ($oTable.getRows.getCount() - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$iToColumn = ($iToColumn = Null) ? ($iColumn) : ($iToColumn)
		$iToRow = ($iToRow = Null) ? ($iRow) : ($iToRow)

		$oCell = $oTable.getCellRangeByPosition($iColumn, $iRow, $iToColumn, $iToRow)
	EndIf

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCell)
EndFunc   ;==>_LOWriter_TableGetCellObjByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableGetData
; Description ...: Retrieve current text of a Text Table.
; Syntax ........: _LOWriter_TableGetData(ByRef $oTable[, $iColumn = -1[, $iRow = -1]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iColumn             - [optional] an integer value. Default is -1. The desired Column, See Remarks.
;                  $iRow                - [optional] an integer value. Default is -1. The desired Row, See Remarks.
; Return values .: Success: Array or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, less than -1 or greater than number of Columns containted in the Table.
;                  @Error 1 @Extended 3 Return 0 = $iRow not an Integer, less than -1 or greater than number of Rows containted in the Table.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Table data.
;                  --Success--
;                  @Error 0 @Extended 1 Return Array of Arrays = Array of Table data.
;                  @Error 0 @Extended 2 Return Array = Returning a specific row of data.
;                  @Error 0 @Extended 3 Return Array = Returning a specific column of data.
;                  @Error 0 @Extended 4 Return String = Returning the data of a specific cell.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If only a Table object is called, an Array of Arrays is returned, The main array will have the same number of elements as there are rows. Each internal array will have the same number of elements as there are columns.
;                  If You input a specific Row, a Array will be returned with the data from that specific row, one element per column.
;                  If You input a Row and a column, a String will be returned with the specified Cell's data.
;                  If you want only a certain column, set $iRow to -1 and $iColumn to the desired column.
;                  LibreOffice Tables start at 0, so to get the first Row/Column, you would set $iRow or $iColumn to 0.
;                  This function can fail if the Table is "complex", meaning it has joined or split cells.
;                  Strings returned will have CRLF for hard newlines, and LF for soft newlines.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableColumnGetCount, _LOWriter_TableRowGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableGetData(ByRef $oTable, $iColumn = Null, $iRow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avTableDataReturn[0], $avTableData

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iColumn <> Null) And Not __LO_IntIsBetween($iColumn, 0, ($oTable.getColumns.getCount() - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iRow <> Null) And Not __LO_IntIsBetween($iRow, 0, ($oTable.getRows.getCount() - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$avTableData = $oTable.getDataArray() ; Will fail if Columns are joined
	If Not IsArray($avTableData) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iColumn, $iRow) Then ; Returning whole Table of Data.

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTableData)

	ElseIf __LO_VarsAreNull($iColumn) Then ; Returning a specific Row's data.
		$avTableDataReturn = $avTableData[$iRow]

		Return SetError($__LO_STATUS_SUCCESS, 2, $avTableDataReturn)

	ElseIf __LO_VarsAreNull($iRow) Then ; Returning a specific Column's data.
		ReDim $avTableDataReturn[UBound($avTableData)]
		For $i = 0 To UBound($avTableData) - 1
			$avTableDataReturn[$i] = ($avTableData[$i])[$iColumn]
		Next

		Return SetError($__LO_STATUS_SUCCESS, 3, $avTableDataReturn)
	EndIf

	; Returning a specific Cell's data.
	$avTableDataReturn = ($avTableData[$iRow])[$iColumn]

	Return SetError($__LO_STATUS_SUCCESS, 4, $avTableDataReturn)
EndFunc   ;==>_LOWriter_TableGetData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableGetObjByCursor
; Description ...: Returns a Table Object, for later Table related functions.
; Syntax ........: _LOWriter_TableGetObjByCursor(ByRef $oDoc, ByRef $oCursor)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions. Cursor object must be located in a Table.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oCursor not located in a Table.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Table Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning Table Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableGetObjByName, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableGetObjByCursor(ByRef $oDoc, ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTable

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetDataType($oDoc, $oCursor) <> $LOW_CURDATA_CELL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Cursor not in Table

	$oTable = $oDoc.TextTables.getByName($oCursor.TextTable.Name)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTable)
EndFunc   ;==>_LOWriter_TableGetObjByCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableGetObjByName
; Description ...: Retrieve a Table Object.
; Syntax ........: _LOWriter_TableGetObjByName(ByRef $oDoc, $sTableName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sTableName          - a string value. Table Name to retrieve the Object for.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sTableName not a String.
;                  @Error 1 @Extended 3 Return 0 = No table matching $sTableName found in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Table Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning an Object for the requested Table.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableGetObjByCursor, _LOWriter_TablesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableGetObjByName(ByRef $oDoc, $sTableName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTable

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sTableName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_TableExists($oDoc, $sTableName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTable = $oDoc.TextTables.getByName($sTableName)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTable)
EndFunc   ;==>_LOWriter_TableGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableMargin
; Description ...: Set or Retrieve the Table Margins.
; Syntax ........: _LOWriter_TableMargin(ByRef $oTable[, $iTopMargin = Null[, $iBottomMargin = Null[, $iLeftMargin = Null[, $iRightMargin = Null]]]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iTopMargin          - [optional] an integer value. Default is Null. The top table margin in Hundredths of a Millimeter (HMM).
;                  $iBottomMargin       - [optional] an integer value. Default is Null. The Bottom table margin in Hundredths of a Millimeter (HMM).
;                  $iLeftMargin         - [optional] an integer value. Default is Null. The Left table margin in Hundredths of a Millimeter (HMM). See Remarks
;                  $iRightMargin        - [optional] an integer value. Default is Null. The Right table margin in Hundredths of a Millimeter (HMM). See Remarks.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTopMargin not an Integer, less than 0 or greater than 100,000.
;                  @Error 1 @Extended 3 Return 0 = $iBottomMargin not an Integer, less than 0 or greater than 100,000
;                  @Error 1 @Extended 4 Return 0 = $iLeftMargin not an Integer, or less than -100,000.
;                  @Error 1 @Extended 5 Return 0 = $iRightMargin not an Integer, or less than -100,000.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Unable to set Left Margin with orientation set to $LOW_ORIENT_HORI_FULL(6) Or $LOW_ORIENT_HORI_LEFT(3).
;                  @Error 3 @Extended 2 Return 0 = Unable to set Right Margin with orientation set to other than $LOW_ORIENT_HORI_NONE(0) Or $LOW_ORIENT_HORI_LEFT(3).
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTopMargin
;                  |                               2 = Error setting $iBottomMargin
;                  |                               4 = Error setting $iLeftMargin
;                  |                               8 = Error setting $iRightMargin
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Left Margin cannot be set unless Table Orientation is set to other than $LOW_ORIENT_HORI_FULL(6), or $LOW_ORIENT_HORI_LEFT(3).
;                  Right Margin cannot be set unless the table orientation is set to $LOW_ORIENT_HORI_NONE(0), or $LOW_ORIENT_HORI_LEFT(3).
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableMargin(ByRef $oTable, $iTopMargin = Null, $iBottomMargin = Null, $iLeftMargin = Null, $iRightMargin = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiMargins[4]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTopMargin, $iBottomMargin, $iLeftMargin, $iRightMargin) Then
		__LO_ArrayFill($aiMargins, $oTable.TopMargin(), $oTable.BottomMargin(), $oTable.LeftMargin(), $oTable.RightMargin())

		Return SetError($__LO_STATUS_SUCCESS, 0, $aiMargins)
	EndIf

	If ($iTopMargin <> Null) Then
		If Not __LO_IntIsBetween($iTopMargin, 0, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oTable.TopMargin = $iTopMargin
		$iError = (__LO_IntIsBetween($oTable.TopMargin(), $iTopMargin - 1, $iTopMargin + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iBottomMargin <> Null) Then
		If Not __LO_IntIsBetween($iBottomMargin, 0, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTable.BottomMargin = $iBottomMargin
		$iError = (__LO_IntIsBetween($oTable.BottomMargin(), $iBottomMargin - 1, $iBottomMargin + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iLeftMargin <> Null) Then
		If Not __LO_IntIsBetween($iLeftMargin, -100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (($oTable.HoriOrient() = $LOW_ORIENT_HORI_FULL) Or ($oTable.HoriOrient() = $LOW_ORIENT_HORI_LEFT)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Can't set Left Margin with orientation set to Auto(6/Full) Or Left (3)

		$oTable.LeftMargin = $iLeftMargin
		$iError = (__LO_IntIsBetween($oTable.LeftMargin(), $iLeftMargin - 1, $iLeftMargin + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iRightMargin <> Null) Then
		If Not __LO_IntIsBetween($iRightMargin, -100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not (($oTable.HoriOrient() = $LOW_ORIENT_HORI_LEFT) Or ($oTable.HoriOrient() = $LOW_ORIENT_HORI_NONE)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Can't set Right Margin with orientation set to other than Manual(0/None) Or Left (3)

		$oTable.RightMargin = $iRightMargin
		$iError = (__LO_IntIsBetween($oTable.RightMargin(), $iRightMargin - 1, $iRightMargin + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_TableMargin

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableProperties
; Description ...: Set or Retrieve Table properties.
; Syntax ........: _LOWriter_TableProperties(ByRef $oDoc, ByRef $oTable[, $iTableAlign = Null[, $bKeepTogether = Null[, $sTableName = Null[, $bSplit = Null[, $bSplitRows = Null[, $bRepeatHeading = Null[, $iHeaderRows = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iTableAlign         - [optional] an integer value (0-7). Default is Null. The horizontal alignment of the Table. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3. Libre Office Default is $LOW_ORIENT_HORI_FULL.
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. If True, prevents page or column breaks between this table and the following paragraph or text table.
;                  $sTableName          - [optional] a string value. Default is Null. The new table name. See Remarks.
;                  $bSplit              - [optional] a boolean value. Default is Null. If False, the table will not split across two pages.
;                  $bSplitRows          - [optional] a boolean value. Default is Null. If True, the content in a Table row is allowed to split at page splits, else if False, Content is not allowed to split across pages.
;                  $bRepeatHeading      - [optional] a boolean value. Default is Null. If True, the first row of the table is repeated on every new page.
;                  $iHeaderRows         - [optional] an integer value. Default is Null. The number of rows to include in the heading.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iTableAlign not an Integer, less than 0 or greater than 7. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $bKeepTogether not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $sTableName not a String.
;                  @Error 1 @Extended 6 Return 0 = Table with same name as called in $sTableName already exists.
;                  @Error 1 @Extended 7 Return 0 = $bSplit not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bSplitRows not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bRepeatHeading not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $iHeaderRows not an Integer, less than 0 or greater than number of rows in table.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTableAlign
;                  |                               2 = Error setting $bKeepTogether
;                  |                               4 = Error setting $sTableName
;                  |                               8 = Error setting $bSplit
;                  |                               16 = Error setting $bSplitRows
;                  |                               32 = Error setting $bRepeatHeading
;                  |                               64 = Error setting $bRepeatHeading
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $bSplitRows will return 0 instead of a boolean if the Table's rows have different settings for $bSplitRows.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableProperties(ByRef $oDoc, ByRef $oTable, $iTableAlign = Null, $bKeepTogether = Null, $sTableName = Null, $bSplit = Null, $bSplitRows = Null, $bRepeatHeading = Null, $iHeaderRows = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avProperties[7]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iTableAlign, $bKeepTogether, $sTableName, $bSplit, $bSplitRows, $bRepeatHeading, $iHeaderRows) Then
		__LO_ArrayFill($avProperties, $oTable.HoriOrient(), $oTable.KeepTogether(), $oTable.getName(), $oTable.Split(), _
				__LOWriter_TableRowSplitToggle($oTable), $oTable.RepeatHeadline(), $oTable.HeaderRowCount())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProperties)
	EndIf

	If ($iTableAlign <> Null) Then
		If Not __LO_IntIsBetween($iTableAlign, $LOW_ORIENT_HORI_NONE, $LOW_ORIENT_HORI_LEFT_AND_WIDTH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTable.HoriOrient = $iTableAlign
		$iError = ($oTable.HoriOrient() = $iTableAlign) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bKeepTogether <> Null) Then
		If Not IsBool($bKeepTogether) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTable.KeepTogether = $bKeepTogether
		$iError = ($oTable.KeepTogether() = $bKeepTogether) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sTableName <> Null) Then
		If Not IsString($sTableName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If _LOWriter_TableExists($oDoc, $sTableName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTable.setName($sTableName)
		$iError = ($oTable.Name() = $sTableName) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bSplit <> Null) Then
		If Not IsBool($bSplit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oTable.Split = $bSplit
		$iError = ($oTable.Split() = $bSplit) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bSplitRows <> Null) Then
		If Not IsBool($bSplitRows) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		__LOWriter_TableRowSplitToggle($oTable, $bSplitRows)
		$iError = (__LOWriter_TableRowSplitToggle($oTable) = $bSplitRows) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bRepeatHeading <> Null) Then
		If Not IsBool($bRepeatHeading) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oTable.RepeatHeadline = $bRepeatHeading
		$iError = ($oTable.RepeatHeadline() = $bRepeatHeading) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iHeaderRows <> Null) Then
		If Not __LO_IntIsBetween($iHeaderRows, 0, $oTable.getRows.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oTable.HeaderRowCount = $iHeaderRows
		$iError = ($oTable.HeaderRowCount() = $iHeaderRows) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) ; error setting Properties.
EndFunc   ;==>_LOWriter_TableProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableRowBackColor
; Description ...: Set the background color of an entire Table row.
; Syntax ........: _LOWriter_TableRowBackColor(ByRef $oTable, $iRow[, $iBackColor = Null])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iRow                - an integer value. The row number to set the background color for. Rows are 0 based.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Row background color, as a RGB Color Integer. See Remarks. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) to disable background color.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iRow not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = Requested row greater than number of rows contained in Table.
;                  @Error 1 @Extended 4 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failure retrieving specified Row object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Background color.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_TableRowGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableRowBackColor(ByRef $oTable, $iRow, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oRow
	Local $iColor

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iRow) Or ($iRow < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oTable.getRows.getCount() < $iRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Requested Row out of bounds.

	$oRow = $oTable.getRows.getByIndex($iRow)
	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iBackColor) Then
		$iColor = $oRow.BackColor()
		If Not IsInt($iColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $iColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oRow.BackColor = $iBackColor
	$iError = ($oRow.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1)) ; Error setting color.

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOWriter_TableRowBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableRowDelete
; Description ...: Delete a row from a Text Table.
; Syntax ........: _LOWriter_TableRowDelete(ByRef $oTable, $iRow[, $iCount = 1])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iRow                - an integer value. The row number to delete. Rows are 0 based.
;                  $iCount              - [optional] an integer value. Default is 1. Number of rows to delete starting at $iRow and moving down.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iRow not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;                  @Error 1 @Extended 4 Return 0 = Requested row greater than number of rows contained in table.
;                  --Success--
;                  @Error 0 @Extended ? Return 1 = Full amount of Rows deleted. @Extended set to total rows deleted.
;                  @Error 0 @Extended ? Return 2 = $iCount greater than amount of rows contained in Table; deleted all rows from $iRow over. @Extended set to total rows deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice counts Rows starting at 0. So to delete the first Row in a Table you would set $iRow to 0.
;                  If you attempt to delete more rows than are present, all rows from $iRow over will be deleted.
;                  If you delete all Rows starting from Row 0, the entire Table is deleted.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableRowGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableRowDelete(ByRef $oTable, $iRow, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iRowCount, $iReturn = 0

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iRow) Or ($iRow < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iCount) Or ($iCount < 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iRowCount = $oTable.getRows.getCount()
	If ($iRowCount <= $iRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Requested Row out of bounds.

	$iCount = ($iCount > ($iRowCount - $iRow)) ? ($iRowCount - $iRow) : ($iCount)
	$iReturn = ($iCount > ($iRowCount - $iRow)) ? (2) : (1) ; Return 1 if full amount deleted else 2 if only partial.
	$oTable.getRows.removeByIndex($iRow, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $iReturn)
EndFunc   ;==>_LOWriter_TableRowDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableRowGetCount
; Description ...: Retrieves the number of Rows in a table.
; Syntax ........: _LOWriter_TableRowGetCount(ByRef $oTable)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Row count.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Returning Row count.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableRowGetCount(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iRowSize = 0

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0) ; Not an Object.

	$iRowSize = $oTable.getRows.getCount()
	If ($iRowSize = 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Failed to retrieve Row count.

	Return SetError($__LO_STATUS_SUCCESS, 0, $iRowSize)
EndFunc   ;==>_LOWriter_TableRowGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableRowInsert
; Description ...: Insert a row into a Text Table
; Syntax ........: _LOWriter_TableRowInsert(ByRef $oTable, $iCount[, $iRow = -1])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iCount              - an integer value. Number of rows to insert.
;                  $iRow                - [optional] an integer value. Default is -1. The row to insert rows after. See Remarks.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iCount not an Integer, or less than 1.
;                  @Error 1 @Extended 3 Return 0 = $iRow not an Integer, or less than -1.
;                  @Error 1 @Extended 4 Return 0 = Requested Row greater than number of Rows contained in table.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to insert Rows.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Successfully inserted requested number of rows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you do not set $iRow, the new Rows will be placed at the Bottom of the Table.
;                  LibreOffice counts the Table Rows starting at 0. The Rows are placed above the desired Row when inserted.
;                  To insert a Row at the top most of the Table you would set $iRow to 0.
;                  To insert rows at the bottom of a table you would set $iRow to one higher than the last row. e.g. a Table containing 3 rows, would be numbered as follows: 0(first-row), 1(second-row), 2(third-row), to insert rows at the very bottom of the rows, I would set $iRow to 3.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableRowGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableRowInsert(ByRef $oTable, $iCount, $iRow = -1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iRowCount

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iCount) Or ($iCount < 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iRow) Or ($iRow < -1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iRowCount = $oTable.getRows.getCount()
	If ($iRowCount < $iRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Requested Row out of bounds.

	$iRow = ($iRow <= -1) ? ($iRowCount) : ($iRow)
	$oTable.getRows.insertByIndex($iRow, $iCount)

	Return ($oTable.getRows.getCount() = ($iRowCount + $iCount)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0))
EndFunc   ;==>_LOWriter_TableRowInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableRowProperty
; Description ...: Set or Retrieve properties for a Text Table.
; Syntax ........: _LOWriter_TableRowProperty(ByRef $oTable, $iRow[, $iHeight = Null[, $bIsAutoHeight = Null[, $bIsSplitAllowed = Null]]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iRow                - an integer value. The Row to set the properties for.
;                  $iHeight             - [optional] an integer value. Default is Null. The row height in Hundredths of a Millimeter (HMM).
;                  $bIsAutoHeight       - [optional] a boolean value. Default is Null. If True, the row's height is automatically adjusted.
;                  $bIsSplitAllowed     - [optional] a boolean value. Default is Null. If False, the row can not be split at a page boundary.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iRow not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = Requested row greater than number of rows contained in Table.
;                  @Error 1 @Extended 4 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $bIsAutoHeight not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bIsSplitAllowed not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failure retrieving specified Row object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iHeight
;                  |                               2 = Error setting $bIsAutoHeight
;                  |                               4 = Error setting $bIsSplitAllowed
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The First row number contained in a table is 0.
;                  None of these properties can be set if the Table is not inserted yet.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LOWriter_TableRowGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableRowProperty(ByRef $oTable, $iRow, $iHeight = Null, $bIsAutoHeight = Null, $bIsSplitAllowed = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRow
	Local $iError = 0
	Local $avProperties[3]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iRow) Or ($iRow < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oTable.getRows.getCount() <= $iRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Requested Row out of bounds.

	$oRow = $oTable.getRows.getByIndex($iRow)
	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iHeight, $bIsAutoHeight, $bIsSplitAllowed) Then
		__LO_ArrayFill($avProperties, $oRow.Height(), $oRow.IsAutoHeight(), $oRow.IsSplitAllowed())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProperties)
	EndIf

	If ($iHeight <> Null) Then
		If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; not an integer

		$oRow.Height = $iHeight
		$iError = (__LO_IntIsBetween($oRow.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bIsAutoHeight <> Null) Then
		If Not IsBool($bIsAutoHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; not a Boolean

		$oRow.IsAutoHeight = $bIsAutoHeight
		$iError = ($oRow.IsAutoHeight() = $bIsAutoHeight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bIsSplitAllowed <> Null) Then
		If Not IsBool($bIsSplitAllowed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; not a Boolean

		$oRow.IsSplitAllowed = $bIsSplitAllowed
		$iError = ($oRow.IsSplitAllowed() = $bIsSplitAllowed) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOWriter_TableRowProperty

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableSetData
; Description ...: Fill a Text Table with Data.
; Syntax ........: _LOWriter_TableSetData(ByRef $oTable, ByRef $avData)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $avData              - [in/out] an array of variants. See Remarks.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $avData not an Array.
;                  @Error 1 @Extended 3 Return 0 = $avData Array does not contain the same number of elements as Rows in the Table.
;                  @Error 1 @Extended 4 Return ? = $avData sub arrays do not contain enough elements to match columns contained in Table. Return set to element number in main array containing faulty array.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Table data was successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The array must be an array of Arrays.
;                  The Main Array must contain the same number of elements as there are rows, and each sub Array must have the same number of Elements as there are columns.
;                  To skip a Cell, just leave the sub array element blank you want to skip.
;                  This will replace all previous data in the Table.
;                  The array is cycled through, and all @CRLF's are replaced with @CR so that additional new lines are not added when inserting the data into LibreOffice.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableSetData(ByRef $oTable, ByRef $avData)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iColumns
	Local $avTemp

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($avData) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (UBound($avData) <> $oTable.getRows.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Array doesn't contain enough elements to match Table.

	$iColumns = $oTable.getColumns.getCount()
	For $i = 0 To UBound($avData) - 1
		If (UBound($avData[$i]) <> $iColumns) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i) ; Array contains too small of array for Table column count.

		$avTemp = $avData[$i]
		For $j = 0 To UBound($avTemp) - 1
			If IsString($avTemp[$j]) Then $avTemp[$j] = StringRegExpReplace($avTemp[$j], @CRLF, @CR)
			Sleep((IsInt($j / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
		$avData[$i] = $avTemp
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next
	$oTable.setDataArray($avData)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_TableSetData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TablesGetNames
; Description ...: Retrieve an array of names for all tables contained in a document.
; Syntax ........: _LOWriter_TablesGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failure retrieving Table objects.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Returning Array of Table Names. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TablesGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTables
	Local $asTableNames[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oTables = $oDoc.TextTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asTableNames = $oTables.getElementNames()

	Return SetError($__LO_STATUS_SUCCESS, UBound($asTableNames), $asTableNames)
EndFunc   ;==>_LOWriter_TablesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableShadow
; Description ...: Set or Retrieve the shadow settings for a Table Border.
; Syntax ........: _LOWriter_TableShadow(ByRef $oTable[, $iWidth = Null[, $iColor = Null[, $iLocation = Null]]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Shadow Width of the Table, set in Hundredths of a Millimeter (HMM).
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Table shadow Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Table Shadow. See constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ShadowFormat Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $iLocation
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  LibreOffice may change the shadow width +/- a Hundredth of a Millimeter (HMM).
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableShadow(ByRef $oTable, $iWidth = Null, $iColor = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tShdwFrmt
	Local $avShadow[3]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tShdwFrmt = $oTable.ShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.Location())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tShdwFrmt.ShadowWidth = $iWidth
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tShdwFrmt.Color = $iColor
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LOW_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oTable.ShadowFormat = $tShdwFrmt

	$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($oTable.ShadowFormat.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iColor)) ? ($iError) : (($oTable.ShadowFormat.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iLocation)) ? ($iError) : (($oTable.ShadowFormat.Location() = $iLocation) ? ($iError) : (BitOR($iError, 4)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_TableShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableStyleCurrent
; Description ...: Set or Retrieve the current Table style for a Table.
; Syntax ........: _LOWriter_TableStyleCurrent(ByRef $oDoc, ByRef $oTable[, $sTableStyle = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $sTableStyle         - [optional] a string value. Default is Null. The Table Style name to set the Table to.
; Return values .: Success: 1 or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sTableStyle not a String.
;                  @Error 1 @Extended 4 Return 0 = TableStyle called in $sTableStyle not found in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Table Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sTableStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Table Style successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning current TableStyle name as a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOWriter_TableStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableStyleCurrent(ByRef $oDoc, ByRef $oTable, $sTableStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sCurrStyle
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sTableStyle) Then
		$sCurrStyle = $oTable.TableTemplateName()
		If Not IsString($sCurrStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrStyle)
	EndIf

	If Not IsString($sTableStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not _LOWriter_TableStyleExists($oDoc, $sTableStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oTable.TableTemplateName = $sTableStyle
	$iError = (__LOWriter_TableStyleCompare($oDoc, $oTable.TableTemplateName(), $sTableStyle)) ? ($iError) : (BitOR($iError, 1))

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOWriter_TableStyleCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableStyleExists
; Description ...: Check whether a Document contains a specific Table Style by name.
; Syntax ........: _LOWriter_TableStyleExists(ByRef $oDoc, $sTableStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sTableStyle         - a string value. The Table Style Name to search for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sTableStyle not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Document contains the Table style called in $sTableStyle, True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableStyleExists(ByRef $oDoc, $sTableStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sTableStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oDoc.StyleFamilies.getByName("TableStyles").hasByName($sTableStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOWriter_TableStyleExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableStylesGetNames
; Description ...: Retrieve an array of all Table Style names available for a document.
; Syntax ........: _LOWriter_TableStylesGetNames(ByRef $oDoc[, $bUserOnly = False[, $bAppliedOnly = False[, $bDisplayName = False]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True only User-Created Table Styles are returned.
;                  $bAppliedOnly        - [optional] a boolean value. Default is False. If True only Applied Table Styles are returned.
;                  $bDisplayName        - [optional] a boolean value. Default is False. If True, the style name displayed in the UI (Display Name), instead of the programmatic style name, is returned. See remarks.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bUserOnly not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bAppliedOnly not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bDisplayName not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Table Style names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. An Array containing all Table Styles matching the called parameters. See remarks. @Extended contains the count of results returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If Only a Document object is input, all available Table styles will be returned.
;                  If Both $bUserOnly and $bAppliedOnly are called with True, only User-Created styles that are applied are returned.
;                  One Table style has a different internal name:
;                  - "Default Table Style" is internally called "Default Style".
;                  Previous to LibreOffice 25.2 either name would work when setting a Style, however after 25.2 only the internal, or programmatic style names, will work.
;                  Calling $bDisplayName with True will return a list of Style names, as the user sees them in the UI, in the same order as they are returned if $bDisplayName is False. It is best not to use these when setting Paragraph Styling.
; Related .......: _LOWriter_TableStyleCurrent, _LOWriter_TableStyleExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableStylesGetNames(ByRef $oDoc, $bUserOnly = False, $bAppliedOnly = False, $bDisplayName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asStyles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDisplayName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$asStyles = __LO_StylesGetNames($oDoc, "TableStyles", $bUserOnly, $bAppliedOnly, $bDisplayName)
	If Not IsArray($asStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asStyles), $asStyles)
EndFunc   ;==>_LOWriter_TableStylesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_TableWidth
; Description ...: Set or Retrieve the Width of a inserted table.
; Syntax ........: _LOWriter_TableWidth(ByRef $oTable[, $iWidth = Null[, $iRelativeWidth = Null]])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iWidth              - [optional] an integer value. Default is Null. The absolute table width in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iRelativeWidth      - [optional] an integer value. Default is Null. The width of the table relative to its environment, in percentage, without a percent sign. See Remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iRelativeWidth not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Unable to set $iWidth with Table orientation set to $LOW_ORIENT_HORI_FULL(6).
;                  @Error 3 @Extended 2 Return 0 = Unable to set $iRelativeWidth with orientation set to $LOW_ORIENT_HORI_FULL(6).
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iRelativeWidth
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters, the third element is a Boolean, If True, the relative width is used, else False means "plain" Width is used.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Relative Width and Width cannot be set until the Table Horizontal orientation is set to other than $LOW_ORIENT_HORI_FULL(6), which is LibeOffice's default setting.
;                  Width may change +/- a Hundredth of a Millimeter (HMM) once set due to Libre Office.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, _LOWriter_TableGetObjByName, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_TableWidth(ByRef $oTable, $iWidth = Null, $iRelativeWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avWidthProps[3]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iRelativeWidth) Then
		__LO_ArrayFill($avWidthProps, $oTable.Width(), $oTable.RelativeWidth(), $oTable.IsWidthRelative())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avWidthProps)
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; not an integer
		If ($oTable.HoriOrient() = $LOW_ORIENT_HORI_FULL) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Can't set Width/ Relative width with orientation set to Auto(6/Full)

		$oTable.Width = $iWidth
		$iError = (__LO_IntIsBetween($oTable.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRelativeWidth <> Null) Then
		If Not IsInt($iRelativeWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; not an integer
		If ($oTable.HoriOrient() = $LOW_ORIENT_HORI_FULL) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Can't set Width/ Relative width with orientation set to Auto(6/Full)

		$oTable.RelativeWidth = $iRelativeWidth
		$iError = (__LO_IntIsBetween($oTable.RelativeWidth(), $iRelativeWidth - 1, $iRelativeWidth + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_TableWidth
