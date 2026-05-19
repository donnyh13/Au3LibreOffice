#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Impress
#include "LibreOfficeImpress_Internal.au3"
#include "LibreOfficeImpress_Constants.au3"

; Other includes for Impress

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for manipulating a Text cursor, inserting or retrieving data or setting and retrieving text properties using an Impress Text Cursor.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_CursorCharEffect
; _LOImpress_CursorCharFont
; _LOImpress_CursorCharFontColor
; _LOImpress_CursorCharOverLine
; _LOImpress_CursorCharPosition
; _LOImpress_CursorCharScaling
; _LOImpress_CursorCharSpacing
; _LOImpress_CursorCharStrikeOut
; _LOImpress_CursorCharUnderLine
; _LOImpress_CursorGetString
; _LOImpress_CursorGoToRange
; _LOImpress_CursorInsertString
; _LOImpress_CursorIsCollapsed
; _LOImpress_CursorMove
; _LOImpress_CursorParAlignment
; _LOImpress_CursorParIndent
; _LOImpress_CursorParSpacing
; _LOImpress_CursorParTabStopCreate
; _LOImpress_CursorParTabStopDelete
; _LOImpress_CursorParTabStopMod
; _LOImpress_CursorParTabStopsGetList
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorCharEffect
; Description ...: Set or Retrieve the Font Effect settings.
; Syntax ........: _LOImpress_CursorCharEffect(ByRef $oTextCursor[, $iCase = Null[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iCase               - [optional] (0-4) Default is Null. The Character Case Style. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRelief             - [optional] (0-2) Default is Null. The Character Relief style. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOutline            - [optional] Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iCase not an Integer, less than 0 or greater than 4. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iRelief not an Integer, less than 0 or greater than 2. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $bOutline not a Boolean.
;                  @Error: 1, @Extended: 5 = $bShadow not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iCase
;                  |                               2 = Error setting $iRelief
;                  |                               4 = Error setting $bOutline
;                  |                               8 = Error setting $bShadow
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorCharEffect(ByRef $oTextCursor, $iCase = Null, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharEffect($oTextCursor, $iCase, $iRelief, $bOutline, $bShadow)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorCharEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorCharFont
; Description ...: Set and Retrieve the Font Settings.
; Syntax ........: _LOImpress_CursorCharFont(ByRef $oTextCursor[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $sFontName           - [optional] Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] Default is Null. The new Font size.
;                  $iPosture            - [optional] (0-5) Default is Null. The Font Italic setting. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] (0, 50-200) Default is Null. The Font Bold settings see Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $sFontName not a String.
;                  @Error: 1, @Extended: 3 = Font called in $sFontName not available.
;                  @Error: 1, @Extended: 4 = $nFontSize not a number.
;                  @Error: 1, @Extended: 5 = $iPosture not an Integer, less than 0 or greater than 5. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iWeight not an Integer, less than 50 but not equal to 0, or greater than 200. See Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $nFontSize
;                  |                               4 = Error setting $iPosture
;                  |                               8 = Error setting $iWeight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;                  LibreOffice accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorCharFont(ByRef $oTextCursor, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharFont($oTextCursor, $sFontName, $nFontSize, $iPosture, $iWeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorCharFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorCharFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting values.
; Syntax ........: _LOImpress_CursorCharFontColor(ByRef $oTextCursor[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iFontColor          - [optional] (-1-16777215) Default is Null. The font Color value, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - [optional] (0-100) Default is Null. Transparency percentage. 0 is visible, 100 is invisible. Available for LibreOffice 7.0 and up.
;                  $iHighlight          - [optional] (-1-16777215) Default is Null. The highlight Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters. If The current LibreOffice version is below 7.0 the $iTransparency parameter will return a Null value.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 3 = $iTransparency not an Integer, less than 0 or greater than 100%.
;                  @Error: 1, @Extended: 4 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $FontColor
;                  |                               2 = Error setting $iTransparency.
;                  |                               4 = Error setting $iHighlight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 7.0.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorCharFontColor(ByRef $oTextCursor, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharFontColor($oTextCursor, $iFontColor, $iTransparency, $iHighlight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorCharFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorCharOverLine
; Description ...: Set and retrieve the OverLine settings for a Paragraph.
; Syntax ........: _LOImpress_CursorCharOverLine(ByRef $oTextCursor[, $iOverLineStyle = Null[, $iOLColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iOverLineStyle      - [optional] (0-18) Default is Null. The style of the Overline line, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iOLColor            - [optional] (-1-16777215) Default is Null. The Overline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not Overlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iOverLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  @Error: 1, @Extended: 3 = $iOLColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $bWordOnly not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iOverLineStyle
;                  |                               2 = Error setting $iOLColor
;                  |                               4 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Overline line style uses the same constants as underline style.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorCharOverLine(ByRef $oTextCursor, $iOverLineStyle = Null, $iOLColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharOverLine($oTextCursor, $iOverLineStyle, $iOLColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorCharOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorCharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size for a Text Cursor.
; Syntax ........: _LOImpress_CursorCharPosition(ByRef $oTextCursor[, $iSuperScript = Null[, $iSubScript = Null[, $iRelativeSize = Null]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iSuperScript        - [optional] (-1-100) Default is Null. The Superscript percentage value. Call with -1 for Automatic SuperScript. See Remarks.
;                  $iSubScript          - [optional] (-1-100) Default is Null. Subscript percentage value. Call with -1 for Automatic SubScript. See Remarks.
;                  $iRelativeSize       - [optional] (1-100) Default is Null. The size percentage relative to current font size.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $oTextCursor does not support Character properties.
;                  @Error: 1, @Extended: 3 = $iSuperScript not an Integer, less than 0 or greater than 100, but not 14000.
;                  @Error: 1, @Extended: 4 = $iSubScript not an Integer, less than -100 or greater than 100, but not 14000 or -14000.
;                  @Error: 1, @Extended: 5 = $iRelativeSize not an Integer, less than 1 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iSuperScript
;                  |                               2 = Error setting $iSubScript
;                  |                               4 = Error setting $iRelativeSize.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Set either $iSubScript or $iSuperScript to 0 to return it to Normal setting.
;                  The way LibreOffice is set up Super/Subscript are set in the same setting, Superscript is a positive number from 1 to 100 (percentage), Subscript is a negative number set to -1 to -100 percentage. For the user's convenience this function automatically converts the positive numbers to negative, and back when setting or retrieving subscript values.
;                  Automatic Superscript has an Integer value of 14000, Auto Subscript has a Integer value of -14000. Being that there is no settable setting of Automatic Super/Sub Script, it has been chosen to use -1 to indicate an automatic Sub/SuperScript value.
;                  If you set both $iSuperScript and $iSubScript to -1 (Automatic), or both $iSuperScript and $iSubScript to any value, Subscript will be the result, as it is the last in the function to be set, and thus will overwrite any Superscript values.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorCharPosition(ByRef $oTextCursor, $iSuperScript = Null, $iSubScript = Null, $iRelativeSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharPosition($oTextCursor, $iSuperScript, $iSubScript, $iRelativeSize)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorCharPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorCharScaling
; Description ...: Set or retrieve the character Scale settings.
; Syntax ........: _LOImpress_CursorCharScaling(ByRef $oTextCursor[, $iScaleWidth = Null])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iScaleWidth         - [optional] (1-100) Default is Null. The percentage to horizontally stretch or compress the text. 100 is normal sizing.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current Scale Width value as an Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iScaleWidth not an Integer or less than 1% or greater than 100%.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Scale width.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iScaleWidth
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Fit to line seems to be unavailable in the API, and does not seem to work in LibreOffice anyway.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorCharScaling(ByRef $oTextCursor, $iScaleWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharScaling($oTextCursor, $iScaleWidth)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorCharScaling

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorCharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning) for a Text Cursor.
; Syntax ........: _LOImpress_CursorCharSpacing(ByRef $oTextCursor[, $bAutoKerning = Null[, $nKerning = Null]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $bAutoKerning        - [optional] Default is Null. If True, applies a spacing in between certain pairs of characters.
;                  $nKerning            - [optional] (-928.8-928.8) Default is Null. The kerning value of the characters. See Remarks. Values are in Printer's Points as set in the LibreOffice UI.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $bAutoKerning not a Boolean.
;                  @Error: 1, @Extended: 3 = $nKerning not a number, less than -928.8 or greater than 928.8 Points.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoKerning
;                  |                               2 = Error setting $nKerning.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User Display, however the internal setting is measured in Hundredths of a Millimeter (HMM). They will be automatically converted from Points to Hundredths of a Millimeter and back for retrieval of settings.
;                  The acceptable values are from -2 Pt to 928.8 Pt. The values can be directly converted easily, however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative Hundredths of a Millimeter internally from 928.9 up to 1000 Pt (Max setting).
;                  For example, 928.8Pt is the last correct value, which equals 32766 Hundredths of a Millimeter (HMM), after this LibreOffice reports the following: 928.9 Pt = -32766 HMM; 929 Pt = -32763 HMM; 929.1 = -32759; 1000 pt = -30258. Attempting to set Libre's kerning value to anything over 32768 HMM causes a COM exception, and attempting to set the kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the max settable kerning is -2.0 Pt to 928.8 Pt.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorCharSpacing(ByRef $oTextCursor, $bAutoKerning = Null, $nKerning = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharSpacing($oTextCursor, $bAutoKerning, $nKerning)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorCharSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorCharStrikeOut
; Description ...: Set or Retrieve the Strikeout settings for a Paragraph.
; Syntax ........: _LOImpress_CursorCharStrikeOut(ByRef $oTextCursor[, $iStrikeLineStyle = Null[, $bWordOnly = Null]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iStrikeLineStyle    - [optional] (0-6) Default is Null. The Strikeout Line Style, see constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bWordOnly           - [optional] Default is Null. If True, strike out is applied to words only, skipping whitespaces.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iStrikeLineStyle not an Integer, less than 0 or greater than 6. See constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $bWordOnly not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iStrikeLineStyle
;                  |                               2 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorCharStrikeOut(ByRef $oTextCursor, $iStrikeLineStyle = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharStrikeOut($oTextCursor, $iStrikeLineStyle, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorCharStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorCharUnderLine
; Description ...: Set and retrieve the Underline settings for a Paragraph.
; Syntax ........: _LOImpress_CursorCharUnderLine(ByRef $oTextCursor[, $iUnderLineStyle = Null[, $iULColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iUnderLineStyle     - [optional] (0-18) Default is Null. The Underline line style, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iULColor            - [optional] (-1-16777215) Default is Null. The underline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not underlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor an Object.
;                  @Error: 1, @Extended: 2 = $iUnderLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iULColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $bWordOnly not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iUnderLineStyle
;                  |                               2 = Error setting $iULColor
;                  |                               4 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorCharUnderLine(ByRef $oTextCursor, $iUnderLineStyle = Null, $iULColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharUnderLine($oTextCursor, $iUnderLineStyle, $iULColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorCharUnderLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorGetString
; Description ...: Retrieve the string of text currently selected by a cursor.
; Syntax ........: _LOImpress_CursorGetString(ByRef $oTextCursor)
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. The selected text as a String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice documentation states that when used in Libre Basic, GetString is limited to 64kb's in size. I do not know if the same limitation applies to any outside use of GetString (such as through Autoit).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorGetString(ByRef $oTextCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sCurrString

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sCurrString = $oTextCursor.getString()
	If Not IsString($sCurrString) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sCurrString)
EndFunc   ;==>_LOImpress_CursorGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorGoToRange
; Description ...: Moves a Text cursor to another Text Cursor Position or Range.
; Syntax ........: _LOImpress_CursorGoToRange(ByRef $oTextCursor, ByRef $oRange[, $bSelect = False])
; Parameters ....: $oTextCursor         - an object. A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $oRange              - an object. The Cursor to move cursor called in $oTextCursor to. A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $bSelect             - [optional] Default is False. If True, the selection is expanded or created from original cursor location to Range location.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Cursor successfully moved to $oRange position.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $oRange not an Object.
;                  @Error: 1, @Extended: 3 = $bSelect not a Boolean.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the Cursor being used as a range has anything selected, the selection will be selected in the Cursor called in $oTextCursor also.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorGoToRange(ByRef $oTextCursor, ByRef $oRange, $bSelect = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextCursor.gotoRange($oRange, $bSelect)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_CursorGoToRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorInsertString
; Description ...: Insert a string at a cursor position.
; Syntax ........: _LOImpress_CursorInsertString(ByRef $oCursor, $sString[, $bOverwrite = False])
; Parameters ....: $oCursor             - A Text Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $sString             - A String to insert.
;                  $bOverwrite          - [optional] Default is False. If True, and the cursor object has text selected, the selection is overwritten, else if False, the string is inserted to the left of the selection. If there are multiple selections, the string is inserted to the left of the last selection, and none are overwritten.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. String was successfully inserted.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCursor not an Object.
;                  @Error: 1, @Extended: 2 = $sString not a string..
;                  @Error: 1, @Extended: 3 = $bOverwrite not a Boolean.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Warning! For some reason this function doesn't seem to set the modified status to True. Changes could be inadvertently lost due to this, if the user closes without saving.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorInsertString(ByRef $oCursor, $sString, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCursor.Text.insertString($oCursor, $sString, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_CursorInsertString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorIsCollapsed
; Description ...: Query whether a Text cursor has any data selected or not.
; Syntax ........: _LOImpress_CursorIsCollapsed(ByRef $oTextCursor)
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning Boolean, True if Cursor has no data selected, else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query Cursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorIsCollapsed(ByRef $oTextCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bReturn = $oTextCursor.isCollapsed()
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOImpress_CursorIsCollapsed

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorMove
; Description ...: Move a Text Cursor object in a document. Also for Creating/Expanding selections.
; Syntax ........: _LOImpress_CursorMove(ByRef $oTextCursor, $iMove[, $iCount = 1[, $bSelect = False]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iMove               - (0-5) The movement command. See remarks. See Constants, $LOI_TEXTCUR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iCount              - [optional] Default is 1. Number of movements to make. See remarks.
;                  $bSelect             - [optional] Default is False. Whether to select data during this cursor movement. See remarks.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: ?, Return: Boolean = Success, Cursor object movement was processed successfully. Returning True if the full count of movements were successful, else False if none or only partially successful. @Extended set to number of successful movements. See Remarks
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iMove not an Integer, less than 0 or greater than 5. See Constants, $LOI_TEXTCUR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iCount not an Integer or less than 0.
;                  @Error: 1, @Extended: 4 = $bSelect not a Boolean.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only some movements accept movement amounts and selecting (such as $LOI_TEXTCUR_GO_RIGHT 2, True) etc. Also only some accept creating/ extending a selection of text/ data. They will be specified below.
;                  To Clear /Unselect a current selection, you can input a move such as $LOI_TEXTCUR_GO_RIGHT, 0, False.
;                  #Cursor Movement Constants which accept Number of Moves and Selecting:
;                  $LOI_TEXTCUR_GO_LEFT,
;                  $LOI_TEXTCUR_GO_RIGHT,
;                  #Cursor Movements which accept Selecting Only:
;                  $LOI_TEXTCUR_GOTO_START,
;                  $LOI_TEXTCUR_GOTO_END,
;                  #Cursor Movements which accept nothing and are done once per call:
;                  $LOI_TEXTCUR_COLLAPSE_TO_START,
;                  $LOI_TEXTCUR_COLLAPSE_TO_END
; Related .......: _LOImpress_CursorIsCollapsed
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorMove(ByRef $oTextCursor, $iMove, $iCount = 1, $bSelect = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCounted = 0
	Local $bMoved = False

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iMove, $LOI_TEXTCUR_COLLAPSE_TO_START, $LOI_TEXTCUR_GOTO_END) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iCount, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Switch $iMove
		Case $LOI_TEXTCUR_GO_LEFT
			$bMoved = $oTextCursor.goLeft($iCount, $bSelect)
			$iCounted = ($bMoved) ? ($iCount) : (0)

		Case $LOI_TEXTCUR_GO_RIGHT
			$bMoved = $oTextCursor.goRight($iCount, $bSelect)
			$iCounted = ($bMoved) ? ($iCount) : (0)

		Case $LOI_TEXTCUR_GOTO_START
			$oTextCursor.gotoStart($bSelect)
			$bMoved = (($oTextCursor.compareRegionStarts($oTextCursor.getStart(), $oTextCursor.Text.getStart()) = 0) ? (True) : (False))
			$iCounted = ($bMoved) ? ($iCount) : (0)

		Case $LOI_TEXTCUR_GOTO_END
			$oTextCursor.gotoEnd($bSelect)
			$bMoved = (($oTextCursor.compareRegionEnds($oTextCursor.getEnd(), $oTextCursor.Text.getEnd()) = 0) ? (True) : (False))
			$iCounted = ($bMoved) ? ($iCount) : (0)

		Case $LOI_TEXTCUR_COLLAPSE_TO_START
			$oTextCursor.collapseToStart()
			$bMoved = (($oTextCursor.compareRegionEnds($oTextCursor.getStart(), $oTextCursor.getEnd()) = 0) ? (True) : (False))
			$iCounted = ($bMoved) ? ($iCount) : (0)

		Case $LOI_TEXTCUR_COLLAPSE_TO_END
			$oTextCursor.collapseToEnd()
			$bMoved = (($oTextCursor.compareRegionStarts($oTextCursor.getStart(), $oTextCursor.getEnd()) = 0) ? (True) : (False))
			$iCounted = ($bMoved) ? ($iCount) : (0)
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)
EndFunc   ;==>_LOImpress_CursorMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorParAlignment
; Description ...: Set and Retrieve Paragraph Alignment settings.
; Syntax ........: _LOImpress_CursorParAlignment(ByRef $oTextCursor[, $iHorAlign = Null[, $iLastLineAlign = Null[, $iTxtDirection = Null]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iHorAlign           - [optional] (0-3) Default is Null. The Horizontal alignment of the paragraph. See Constants, $LOI_PAR_ALIGN_HOR_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iLastLineAlign      - [optional] (0-3) Default is Null. Specify the alignment for the last line in the paragraph. See Constants, $LOI_PAR_LAST_LINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iTxtDirection       - [optional] (0-5) Default is Null. The Text Writing Direction. See Constants, $LOI_PAR_TXT_DIR_* as defined in LibreOfficeImpress_Constants.au3. [LibreOffice Default is 4]
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iHorAlign not an Integer, less than 0 or greater than 3. See Constants, $LOI_PAR_ALIGN_HOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iLastLineAlign not an Integer, less than 0 or greater than 3. See Constants, $LOI_PAR_LAST_LINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iTxtDirection not an Integer, less than 0 or greater than 5. See Constants, $LOI_PAR_TXT_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iHorAlign
;                  |                               2 = Error setting $iLastLineALign
;                  |                               4 = Error setting $iTxtDirection
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iHorAlign must be set to $LOI_PAR_ALIGN_HOR_JUSTIFIED(2) before you can set $iLastLineAlign.
;                  $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Expand single word, Snap to grid, and Vertical align (Text-To-Text), seem to be unavailable in the API, and do not seem to work in LibreOffice.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorParAlignment(ByRef $oTextCursor, $iHorAlign = Null, $iLastLineAlign = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParAlignment($oTextCursor, $iHorAlign, $iLastLineAlign, $iTxtDirection)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorParAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorParIndent
; Description ...: Set or Retrieve Paragraph Indent settings.
; Syntax ........: _LOImpress_CursorParIndent(ByRef $oTextCursor[, $iBeforeTxt = Null[, $iAfterTxt = Null[, $iFirstLine = Null]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iBeforeTxt          - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM).
;                  $iAfterTxt           - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM)
;                  $iFirstLine          - [optional] (0-1162202) Default is Null. Indentation distance of the first line of a paragraph. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iBeforeText not an Integer, less than 0 or greater than 1162202.
;                  @Error: 1, @Extended: 3 = $iAfterText not an Integer, less than 0 or greater than 1162202.
;                  @Error: 1, @Extended: 4 = $iFirstLine not an Integer, less than 0 or greater than 1162202.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBeforeTxt
;                  |                               2 = Error setting $iAfterTxt
;                  |                               4 = Error setting $iFirstLine
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Auto indent first line does not seem to work in LibreOffice, and seems to be not available in the API.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorParIndent(ByRef $oTextCursor, $iBeforeTxt = Null, $iAfterTxt = Null, $iFirstLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParIndent($oTextCursor, $iBeforeTxt, $iAfterTxt, $iFirstLine)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorParIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorParSpacing
; Description ...: Set and Retrieve Line Spacing settings.
; Syntax ........: _LOImpress_CursorParSpacing(ByRef $oTextCursor[, $iAbovePar = Null[, $iBelowPar = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null]]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iAbovePar           - [optional] (0-100000) Default is Null. The Space above a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iBelowPar           - [optional] (0-100000) Default is Null. The Space Below a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iLineSpcMode        - [optional] (0-3) Default is Null. The line spacing type of the paragraph. See Constants, $LOI_PAR_LINE_SPC_MODE_* as defined in LibreOfficeImpress_Constants.au3, also notice min and max values for each.
;                  $iLineSpcHeight      - [optional] Default is Null. This value specifies the height in regard to Mode. See Remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iAbovePar not an Integer, less than 0 or greater than 100000.
;                  @Error: 1, @Extended: 3 = $iBelowPar not an Integer, less than 0 or greater than 100000.
;                  @Error: 1, @Extended: 4 = $iLineSpcMode not an Integer, less than 0 or greater than 3. See Constants, $LOI_PAR_LINE_SPC_MODE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iLineSpcHeight not an Integer.
;                  @Error: 1, @Extended: 6 = $iLineSpcMode set to 0(Proportional) and $iLineSpcHeight less than 6(%) or greater than 65535(%).
;                  @Error: 1, @Extended: 7 = $iLineSpcMode set to 1 or 2(Minimum, or Leading) and $iLineSpcHeight less than 0 or greater than 100000.
;                  @Error: 1, @Extended: 8 = $iLineSpcMode set to 3(Fixed) and $iLineSpcHeight less than 51 or greater than 100000.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaLineSpacing Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAbovePar
;                  |                               2 = Error setting $iBelowPar
;                  |                               4 = Error setting $iLineSpcMode
;                  |                               8 = Error setting $iLineSpcHeight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The settings in LibreOffice, (Single, 1.15, 1.5, Double), Use the Proportional mode, and are just varying percentages. e.g Single = 100, 1.15 = 115%, 1.5 = 150%, Double = 200%.
;                  $iLineSpcHeight depends on the $iLineSpcMode used, see constants for accepted Input values.
;                  $iAbovePar, $iBelowPar, $iLineSpcHeight may change +/- a Hundredth of a Millimeter (HMM) once set.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  The "Do not add space between paragraphs as the same style" setting seems to be not available to set or retrieve in the API, and seems to do nothing in LibreOffice anyway.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorParSpacing(ByRef $oTextCursor, $iAbovePar = Null, $iBelowPar = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParSpacing($oTextCursor, $iAbovePar, $iBelowPar, $iLineSpcMode, $iLineSpcHeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorParSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorParTabStopCreate
; Description ...: Create a new TabStop for a Paragraph.
; Syntax ........: _LOImpress_CursorParTabStopCreate(ByRef $oTextCursor, $iPosition[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iPosition           - The TabStop position to set the new TabStop to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Settings were successfully set. New TabStop position is returned.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iPosition not an Integer.
;                  @Error: 1, @Extended: 3 = Tab Stop position called in $iPosition already exists in this Paragraph.
;                  @Error: 1, @Extended: 4 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iDecChar not an Integer.
;                  @Error: 1, @Extended: 6 = $iFillChar not an Integer.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.style.TabStop" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Array Object.
;                  @Error: 3, @Extended: 2 = Error retrieving list of TabStop Positions.
;                  @Error: 3, @Extended: 3 = Failed to identify the new Tabstop once inserted.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iAlignment
;                  |                               4 = Error setting $iDecChar
;                  |                               8 = Error setting $iFillChar
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iPosition once set can vary +/- a Hundredth of a Millimeter (HMM). To ensure you can identify the tabstop to modify it again, This function returns the new TabStop position.
;                  Since $iPosition can fluctuate +/- a Hundredth of a Millimeter (HMM) when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;                  $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32. The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
;                  Call any optional parameter with Null keyword to skip it.
;                  $iNewTabStop position is still returned as even though some settings weren't successfully set, the new TabStop was still created.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorParTabStopCreate(ByRef $oTextCursor, $iPosition, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopCreate($oTextCursor, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorParTabStopCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorParTabStopDelete
; Description ...: Delete a TabStop from a Paragraph.
; Syntax ........: _LOImpress_CursorParTabStopDelete(ByRef $oTextCursor, $iTabStop)
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Returning True if TabStop was successfully deleted, else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iTabStop not an Integer.
;                  @Error: 1, @Extended: 3 = $iTabStop not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per paragraph.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorParTabStopDelete(ByRef $oTextCursor, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopDelete($oTextCursor, $iTabStop)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorParTabStopDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop.
; Syntax ........: _LOImpress_CursorParTabStopMod(ByRef $oTextCursor, $iTabStop[, $iPosition = Null[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - [optional] Default is Null. The New position to set the input position to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: ?, Return: 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iTabStop not an Integer.
;                  @Error: 1, @Extended: 3 = TabStop called in $iTabStop not found.
;                  @Error: 1, @Extended: 4 = $iPosition not an Integer.
;                  @Error: 1, @Extended: 5 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iDecChar not an Integer.
;                  @Error: 1, @Extended: 7 = $iFillChar not an Integer.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Requested TabStop Object.
;                  @Error: 3, @Extended: 3 = Paragraph already contains a TabStop at the length/Position specified in $iPosition.
;                  @Error: 3, @Extended: 4 = Error retrieving list of TabStop Positions.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iAlignment
;                  |                               4 = Error setting $iDecChar
;                  |                               8 = Error setting $iFillChar
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per Paragraph.
;                  $iPosition once set can vary +/- a Hundredth of a Millimeter (HMM). To ensure you can identify the tabstop to modify it again, This function returns the new TabStop position in @Extended when $iPosition is set, return value will be set to 2. See Return Values.
;                  Since $iPosition can fluctuate +/- a Hundredth of a Millimeter (HMM) when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;                  $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32. The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorParTabStopMod(ByRef $oTextCursor, $iTabStop, $iPosition = Null, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopMod($oTextCursor, $iTabStop, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorParTabStopMod

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorParTabStopsGetList
; Description ...: Retrieve an array of TabStops available in a Paragraph.
; Syntax ........: _LOImpress_CursorParTabStopsGetList(ByRef $oTextCursor)
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array of TabStops. @Extended set to number of results.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorParTabStopsGetList(ByRef $oTextCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopsGetList($oTextCursor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_CursorParTabStopsGetList
