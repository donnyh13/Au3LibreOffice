#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter)
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.0.3
; Description ...: Provides basic functionality through Autoit for interacting with Libre Office Writer.
; Author(s) .....: donnyh13, mLipok
; Sources .......: jguinch -- Printmgr.au3, used (_PrintMgr_EnumPrinter);
;					mLipok -- OOoCalc.au3, used (__OOoCalc_ComErrorHandler_UserFunction,_InternalComErrorHandler,
;						-- WriterDemo.au3, used _CreateStruct;
;					Andrew Pitonyak & Laurent Godard (VersionGet);
;					Leagnus & GMK -- OOoCalc.au3, used (SetPropertyValue)
; Dll ...........:
; Note...........: Tips/templates taken from OOoCalc UDF written by user GMK; also from Word UDF by user water.
;					I found the book by Andrew Pitonyak very helpful also, titled, "OpenOffice.org Macros Explained;
;						OOME Third Edition".
;					Of course, this UDF is written using the English version of LibreOffice, and may only work for the English
;						version of LibreOffice installations. Many functions in this UDF may or may not work with OpenOffice
;						Writer, however some settings are definitely for LibreOffice only.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_CellBackColor
; _LOWriter_CellBorderColor
; _LOWriter_CellBorderPadding
; _LOWriter_CellBorderStyle
; _LOWriter_CellBorderWidth
; _LOWriter_CellCreateTextCursor
; _LOWriter_CellFormula
; _LOWriter_CellGetDataType
; _LOWriter_CellGetError
; _LOWriter_CellGetName
; _LOWriter_CellProtect
; _LOWriter_CellString
; _LOWriter_CellValue
; _LOWriter_CellVertOrient
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBackColor
; Description ...: Set and Retrieve the Background color of a Cell or Cell Range.
; Syntax ........: _LOWriter_CellBackColor(Byref $oCell[, $iBackColor = Null[, $bBackTransparent = Null]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. Specify the Cell background color as a Long Integer. See Remarks.
;				   +						Set to $LOW_COLOR_OFF(-1) to disable Background color. Can also be one of the constants $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3
;                  $bBackTransparent    - [optional] a boolean value. Default is Null. If True, the background color is transparent.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell variable not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iBackColor not an Integer, set less than -1 or greater than 16777215.
;				   @Error 1 @Extended 3 Return 0 = $bBackTransparent not a Boolean and not set to Null.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iBackColor
;				   |								2 = Error setting $bBackTransparent
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it. $iBackColor is set using Long integer.
;					See _LOWriter_ConvertColorToLong, _LOWriter_ConvertColorFromLong.
; Related .......: _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBackColor(ByRef $oCell, $iBackColor = Null, $bBackTransparent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avColor[2]

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iBackColor, $bBackTransparent) Then
		__LOWriter_ArrayFill($avColor, $oCell.BackColor(), $oCell.BackTransparent())
		Return SetError($__LOW_STATUS_SUCCESS, 0, $avColor)
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iBackColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oCell.BackColor = $iBackColor
		If ($iBackColor = $LOW_COLOR_OFF) Then $oCell.BackTransparent = True
		$iError = ($oCell.BackColor() = $iBackColor) ? $iError : BitOR($iError, 1) ; Error setting color.
	EndIf

	If ($bBackTransparent <> Null) Then
		If Not IsBool($bBackTransparent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oCell.BackTransparent = $bBackTransparent
		$iError = ($oCell.BackTransparent() = $bBackTransparent) ? $iError : BitOR($iError, 2) ; Error setting BackTransparent.
	EndIf

	Return ($iError = 0) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0)
EndFunc   ;==>_LOWriter_CellBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBorderColor
; Description ...: Set the Cell or Cell Range Border Line Color. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_CellBorderColor(Byref $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. Sets the Top Border Line Color of the Cell in Long Color code format.
;				   +						A custom value or one of the constants $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. Sets the Bottom Border Line Color of the Cell in Long Color code format.
;				   +						A custom value or one of the constants $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. Sets the Left Border Line Color of the Cell in Long Color code format.
;				   +						A custom value or one of the constants $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. Sets the Right Border Line Color of the Cell in Long Color code format.
;				   +						A custom value or one of the constants $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
; Internal Remark: Error values for Initialization and Processing are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell Variable not Object type variable.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an integer, or set to less than 0 or higher than 16,777,215.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an integer, or set to less than 0 or higher than 16,777,215.
;				   @Error 1 @Extended 4 Return 0 = $iLeft not an integer, or set to less than 0 or higher than 16,777,215.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an integer, or set to less than 0 or higher than 16,777,215.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Color when Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Color when Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border Color when Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border Color when Border width not set.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_CellBorderWidth,
;					_LOWriter_CellBorderStyle, _LOWriter_CellBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBorderColor(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oCell, False, False, True, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)

EndFunc   ;==>_LOWriter_CellBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBorderPadding
; Description ...: Set or retrieve the Border Padding (spacing between the Cell text and border) settings.
; Syntax ........: _LOWriter_CellBorderPadding(Byref $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Distance between the Border and Cell text in Micrometers(uM).
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Distance between the Border and Cell text in Micrometers(uM).
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Distance between the Border and Cell text in Micrometers(uM).
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Distance between the Border and Cell text in Micrometers(uM).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell parameter not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $Left not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iTop border distance
;				   |								2 = Error setting $iBottom border distance
;				   |								4 = Error setting $iLeft border distance
;				   |								8 = Error setting $iRight border distance
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_CellBorderColor,
;					_LOWriter_CellBorderStyle, _LOWriter_CellBorderWidth
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBorderPadding(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[4]

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		__LOWriter_ArrayFill($aiBPadding, $oCell.TopBorderDistance(), $oCell.BottomBorderDistance(), $oCell.LeftBorderDistance(), $oCell.RightBorderDistance())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iTop <> Null) Then
		If Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oCell.TopBorderDistance = $iTop
		$iError = (__LOWriter_IntIsBetween($oCell.TopBorderDistance(), $iTop - 1, $iTop + 1)) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iBottom <> Null) Then
		If Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oCell.BottomBorderDistance = $iBottom
		$iError = (__LOWriter_IntIsBetween($oCell.BottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iLeft <> Null) Then
		If Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oCell.LeftBorderDistance = $iLeft
		$iError = (__LOWriter_IntIsBetween($oCell.LeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iRight <> Null) Then
		If Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oCell.RightBorderDistance = $iRight
		$iError = (__LOWriter_IntIsBetween($oCell.RightBorderDistance(), $iRight - 1, $iRight + 1)) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBorderStyle
; Description ...: Set or Retrieve the Cell or Cell Range Border Line style. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_CellBorderStyle(Byref $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Style of the Cell using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Style of the Cell using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Style of the Cell using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Style of the Cell using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3
; Internal Remark: Error values for Initialization and Processing are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell Variable not Object type variable.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an integer, or set to higher than 17 and not equal to 0x7FFF, or $iTop is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an integer, or set to higher than 17 and not equal to 0x7FFF, or $iBottom is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iLeft not an integer, or set to higher than 17 and not equal to 0x7FFF, or $iLeft is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an integer, or set to higher than 17 and not equal to 0x7FFF, or $iRight is set to less than 0 or not set to Null.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style when Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style when Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style when Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style when Border width not set.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_CellBorderWidth, _LOWriter_CellBorderColor, _LOWriter_CellBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBorderStyle(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oCell, False, True, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CellBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBorderWidth
; Description ...: Set or Retrieve the Cell or Cell Range Border Line Width. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_CellBorderWidth(Byref $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line width of the Cell in MicroMeters. Can be a custom value or one of the predefined constants listed $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Width of the Cell in MicroMeters. Can be a custom value or one of the predefined constants listed $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line width of the Cell in MicroMeters. Can be a custom value or one of the predefined constants listed $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Width of the Cell in MicroMeters. Can be a custom value or one of the predefined constants listed $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3
; Internal Remark: Error values for Initialization and Processing are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell Variable not Object type variable.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iLeft not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an integer, or set to less than 0 or not set to Null.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set them to 0
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_CellBorderStyle,
;					_LOWriter_CellBorderColor, _LOWriter_CellBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBorderWidth(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oCell, True, False, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CellBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellCreateTextCursor
; Description ...: Create a Text Cursor in a particular cell for inserting text etc.
; Syntax ........: _LOWriter_CellCreateTextCursor(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
; Return values .: Success: An Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. A Text Cursor Object located in the specified Cell.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellCreateTextCursor(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only create a Text Cursor for individual cells.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.Text.createTextCursor())
EndFunc   ;==>_LOWriter_CellCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellFormula
; Description ...: Set or retrieve a formula for a cell.
; Syntax ........: _LOWriter_CellFormula(Byref $oCell[, $sFormula = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $sFormula            - [optional] a string value. Default is Null. The Formula to set the Cell to.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFormula not a String and not set to Null keyword.
;				   @Error 1 @Extended 3 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Formula was successfully set.
;				   @Error 0 @Extended 1 Return String = Success. Current formula is returned in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Formula can only be set for an individual cell, not a range.
;					Setting the formula will overwrite any existing data in the cell.
; 					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					To retrieve the total of a formula, use _LOWriter_CellValue.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellFormula(ByRef $oCell, $sFormula = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFormula) And Not ($sFormula = Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Can only set/get formula value for individual cells.
	If ($sFormula = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oCell.getFormula())

	$oCell.setFormula($sFormula)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellGetDataType
; Description ...: Get the Data type of a specific cell, see remarks.
; Syntax ........: _LOWriter_CellGetDataType(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
; Return values .: Success: A Number.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return Number = Success. The Data Type in Number format, see constants, $LOW_CELL_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns the data type as one of the below constants, Note: If the data was entered by the keyboard, it is generally recognized as a string regardless of the data contained.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellGetDataType(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only get Data Type for individual cells

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.getType())
EndFunc   ;==>_LOWriter_CellGetDataType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellGetError
; Description ...: Get the formula error Value.
; Syntax ........: _LOWriter_CellGetError(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
; Return values .: Success: A Number.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return Number = Success. The Cell formula error code in Number format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Integer error value. If the cell is not a formula, the error value is zero.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellGetError(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only get Error for individual cells.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.getError())

EndFunc   ;==>_LOWriter_CellGetError

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellGetName
; Description ...: Retrieve the current Cell's name.
; Syntax ........: _LOWriter_CellGetName(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
; Return values .: Success: A String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. The Cell name in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellGetName(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only get Cell Name for individual cells.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.CellName())
EndFunc   ;==>_LOWriter_CellGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellProtect
; Description ...: Write-Protect a Cell
; Syntax ........: _LOWriter_CellProtect(Byref $oCell[, $bProtect = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $bProtect            - [optional] a boolean value. Default is Null. True = Protected from Writing, False = Unprotected. See remarks.
; Return values .: Success: 1 Or Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range. Can only set Write-Protect on individual cells.
;				   @Error 1 @Extended 3 Return 0 = $bProtect not a Boolean or not Null keyword.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Failed to set Write-Protect property.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully set Cell Protect setting.
;				   @Error 0 @Extended 0 Return Boolean = Success. $bProtect is set to Null, return will be the current setting of write-protection for the cell, a Boolean value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling $bProtect with Null keyword returns the current WriteProtection setting of the cell. (True or
;					False)
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellProtect(ByRef $oCell, $bProtect = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only set individual cell protect property.
	If ($bProtect = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.IsProtected())
	If Not IsBool($bProtect) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oCell.IsProtected = $bProtect
	Return ($oCell.IsProtected() = $bProtect) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)

EndFunc   ;==>_LOWriter_CellProtect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellString
; Description ...: Set or retrieve the current string for a cell.
; Syntax ........: _LOWriter_CellString(Byref $oCell[, $sString = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $sString             - [optional] a string value. Default is Null. The String of text to set the cell to.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sString not a String and not set to Null keyword.
;				   @Error 1 @Extended 3 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. String was successfully set.
;				   @Error 0 @Extended 1 Return String = Success. Current String is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: String can only be set for an individual cell, not a range.
;					Setting the String will overwrite any existing data in the cell.
; 					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellString(ByRef $oCell, $sString = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sString) And Not ($sString = Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Can only set/get a String for individual cells.

	If ($sString = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oCell.getString())

	$oCell.setString($sString)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellValue
; Description ...: Set or retrieve a Numerical value to a Cell
; Syntax ........: _LOWriter_CellValue(Byref $oCell[, $nValue = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $nValue              - [optional] a general number value. Default is Null. The value to set the cell to.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $nValue not a Number and not set to Null keyword.
;				   @Error 1 @Extended 3 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Value was successfully set.
;				   @Error 0 @Extended 1 Return String = Success. Current Value is returned in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Value can only be set for an individual cell, not a range.
;					Setting the Value will overwrite any existing data in the cell.
;					For a value cell the value is returned, for a string cell zero is returned and for a formula cell the result value of a formula is returned.
; 					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellValue(ByRef $oCell, $nValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsNumber($nValue) And Not ($nValue = Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Can only set/get individual cell values.

	If ($nValue = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oCell.getValue())

	$oCell.setValue($nValue)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellVertOrient
; Description ...: Set the Vertical Orientation of the Cell or Cell Range contents.
; Syntax ........: _LOWriter_CellVertOrient(Byref $oCell[, $iVertOrient = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition functions.
;                  $iVertOrient         - [optional] an integer value (0-3). Default is Null. A Vertical Orientation constant. $LOW_ORIENT_VERT_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iVertOrient not an integer, or less than 0 or greater than 3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Failed to set Cell Vertical Orientation property.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1  = Success. Succesfully set Vertical Orientation.
;				   @Error 0 @Extended 0 Return Integer = Success. $iVertOrient is set to Null, returning current Cell Vertical orientation, see constants $LOW_ORIENT_VERT_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $iVertOrient is set to Null the current setting is returned.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellVertOrient(ByRef $oCell, $iVertOrient = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	; 3 = Vert Orient Bottom, 1 = Vert orient Top

	If ($iVertOrient = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.VertOrient())
	If Not __LOWriter_IntIsBetween($iVertOrient, $LOW_ORIENT_VERT_NONE, $LOW_ORIENT_VERT_BOTTOM) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oCell.VertOrient = $iVertOrient

	Return ($oCell.VertOrient() = $iVertOrient) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)
EndFunc   ;==>_LOWriter_CellVertOrient
