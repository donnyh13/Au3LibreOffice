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
#include "LibreOfficeWriter_Char.au3"
#include "LibreOfficeWriter_Page.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Modifying and Applying Direct Character and Paragraph formatting in L.O. Writer.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_DirFrmtCharBorderColor
; _LOWriter_DirFrmtCharBorderPadding
; _LOWriter_DirFrmtCharBorderStyle
; _LOWriter_DirFrmtCharBorderWidth
; _LOWriter_DirFrmtCharEffect
; _LOWriter_DirFrmtCharFont
; _LOWriter_DirFrmtCharFontColor
; _LOWriter_DirFrmtCharOverLine
; _LOWriter_DirFrmtCharPosition
; _LOWriter_DirFrmtCharRotateScale
; _LOWriter_DirFrmtCharShadow
; _LOWriter_DirFrmtCharSpacing
; _LOWriter_DirFrmtCharStrikeOut
; _LOWriter_DirFrmtCharUnderLine
; _LOWriter_DirFrmtClear
; _LOWriter_DirFrmtGetCurStyles
; _LOWriter_DirFrmtParAlignment
; _LOWriter_DirFrmtParAreaColor
; _LOWriter_DirFrmtParAreaFillStyle
; _LOWriter_DirFrmtParAreaGradient
; _LOWriter_DirFrmtParAreaGradientMulticolor
; _LOWriter_DirFrmtParAreaTransparency
; _LOWriter_DirFrmtParBorderColor
; _LOWriter_DirFrmtParBorderPadding
; _LOWriter_DirFrmtParBorderStyle
; _LOWriter_DirFrmtParBorderWidth
; _LOWriter_DirFrmtParDropCaps
; _LOWriter_DirFrmtParHyphenation
; _LOWriter_DirFrmtParIndent
; _LOWriter_DirFrmtParOutLineAndList
; _LOWriter_DirFrmtParPageBreak
; _LOWriter_DirFrmtParShadow
; _LOWriter_DirFrmtParSpace
; _LOWriter_DirFrmtParTabStopCreate
; _LOWriter_DirFrmtParTabStopDelete
; _LOWriter_DirFrmtParTabStopMod
; _LOWriter_DirFrmtParTabStopsGetList
; _LOWriter_DirFrmtParTxtFlowOpt
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharBorderColor
; Description ...: Set and Retrieve the Character Style Border Line Color by Direct Formatting. LibreOffice 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharBorderColor(ByRef $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iTop                - [optional] (0-16777215) Default is Null. The Top Border Line Color of the Character Style as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] (0-16777215) Default is Null. The Bottom Border Line Color of the Character Style as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] (0-16777215) Default is Null. The Left Border Line Color of the Character Style as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] (0-16777215) Default is Null. The Right Border Line Color of the Character Style as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of border, Width, Style and Color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  @Error: 1, @Extended: 3 = $iTop not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $iBottom not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iLeft not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 6 = $iRight not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error: 3, @Extended: 2 = Cannot set Top Border Color when Top Border width not set.
;                  @Error: 3, @Extended: 3 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error: 3, @Extended: 4 = Cannot set Left Border Color when Left Border width not set.
;                  @Error: 3, @Extended: 5 = Cannot set Right Border Color when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 4.2.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  Border Width must be set first to be able to set Border Style and Color.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_DirFrmtCharBorderWidth, _LOWriter_DirFrmtCharBorderStyle, _LOWriter_DirFrmtCharBorderPadding, _LOWriter_DirFrmtParBorderColor, _LOWriter_CharStyleBorderColor, _LOWriter_ParStyleBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharBorderColor(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharTopBorder")
		$oSelection.setPropertyToDefault("CharBottomBorder") ; Resetting one truly resets all, but just to be sure, reset all.
		$oSelection.setPropertyToDefault("CharLeftBorder")
		$oSelection.setPropertyToDefault("CharRightBorder")
		If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_CharBorder($oSelection, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharBorderPadding
; Description ...: Set and retrieve the distance between the border and the characters by Direct Format. LibreOffice 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharBorderPadding(ByRef $oSelection[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iAll                - [optional] Default is Null. Set all four values to the same value. When used, all other parameters are ignored. In Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] Default is Null. The Top border distance in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] Default is Null. The Bottom border distance in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] Default is Null. The left border distance in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] Default is Null. The Right border distance in Hundredths of a Millimeter (HMM).
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of border padding, on all sides.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iAll not an Integer.
;                  @Error: 1, @Extended: 3 = $iTop not an Integer.
;                  @Error: 1, @Extended: 4 = $iBottom not an Integer.
;                  @Error: 1, @Extended: 5 = $Left not an Integer.
;                  @Error: 1, @Extended: 6 = $iRight not an Integer.
;                  @Error: 1, @Extended: 7 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAll border distance
;                  |                               2 = Error setting $iTop border distance
;                  |                               4 = Error setting $iBottom border distance
;                  |                               8 = Error setting $iLeft border distance
;                  |                               16 = Error setting $iRight border distance
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 4.2.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtCharBorderWidth, _LOWriter_DirFrmtCharBorderStyle, _LOWriter_DirFrmtCharBorderColor, _LOWriter_DirFrmtParBorderPadding, _LOWriter_CharStyleBorderPadding, _LOWriter_ParStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharBorderPadding(ByRef $oSelection, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If $bClearDirFrmt Then
		; Resetting any one of these settings causes all to reset; reset the "All" setting for quickness.
		$oSelection.setPropertyToDefault("CharBorderDistance")
		If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharBorderPadding($oSelection, $iAll, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharBorderStyle
; Description ...: Set or Retrieve the Character Style Border Line style by Direct Format. LibreOffice 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharBorderStyle(ByRef $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iTop                - [optional] (0x7FFF,0-17) Default is Null. The Top Border Line Style of the Characters. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] (0x7FFF,0-17) Default is Null. The Bottom Border Line Style of the Characters. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] (0x7FFF,0-17) Default is Null. The Left Border Line Style of the Characters. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] (0x7FFF,0-17) Default is Null. The Right Border Line Style of the Characters. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of border, Width, Style and Color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  @Error: 1, @Extended: 3 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error: 3, @Extended: 2 = Cannot set Top Border Style when Top Border width not set.
;                  @Error: 3, @Extended: 3 = Cannot set Bottom Border Style when Bottom Border width not set.
;                  @Error: 3, @Extended: 4 = Cannot set Left Border Style when Left Border width not set.
;                  @Error: 3, @Extended: 5 = Cannot set Right Border Style when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 4.2.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  Border Width must be set first to be able to set Border Style and Color.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOWriter_DirFrmtCharBorderWidth, _LOWriter_DirFrmtCharBorderColor, _LOWriter_DirFrmtCharBorderPadding, _LOWriter_DirFrmtParBorderStyle, _LOWriter_CharStyleBorderStyle, _LOWriter_ParStyleBorderStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharBorderStyle(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharTopBorder")
		$oSelection.setPropertyToDefault("CharBottomBorder") ; Resetting one truly resets all, but just to be sure, reset all.
		$oSelection.setPropertyToDefault("CharLeftBorder")
		$oSelection.setPropertyToDefault("CharRightBorder")
		If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOW_BORDER_STYLE_SOLID, $LOW_BORDER_STYLE_DASH_DOT_DOT, "", $LOW_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOW_BORDER_STYLE_SOLID, $LOW_BORDER_STYLE_DASH_DOT_DOT, "", $LOW_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOW_BORDER_STYLE_SOLID, $LOW_BORDER_STYLE_DASH_DOT_DOT, "", $LOW_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOW_BORDER_STYLE_SOLID, $LOW_BORDER_STYLE_DASH_DOT_DOT, "", $LOW_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_CharBorder($oSelection, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharBorderWidth
; Description ...: Set and Retrieve the Character Style Border Line Width by Direct Formatting. LibreOffice 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharBorderWidth(ByRef $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iTop                - [optional] Default is Null. The Top Border Line width of the Character Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of these constants, $LOW_BORDER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] Default is Null. The Bottom Border Line Width of the Character Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of these constants, $LOW_BORDER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] Default is Null. The Left Border Line width of the Character Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of these constants, $LOW_BORDER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] Default is Null. The Right Border Line Width of the Character Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of these constants, $LOW_BORDER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of border, Width, Style and Color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  @Error: 1, @Extended: 3 = $iTop not an Integer, or less than 0.
;                  @Error: 1, @Extended: 4 = $iBottom not an Integer, or less than 0.
;                  @Error: 1, @Extended: 5 = $iLeft not an Integer, or less than 0.
;                  @Error: 1, @Extended: 6 = $iRight not an Integer, or less than 0.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 4.2.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To "Turn Off" Borders, set them to 0
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtCharBorderStyle, _LOWriter_DirFrmtCharBorderColor, _LOWriter_DirFrmtCharBorderPadding, _LOWriter_DirFrmtParBorderWidth, _LOWriter_CharStyleBorderWidth, _LOWriter_ParStyleBorderWidth
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharBorderWidth(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharTopBorder")
		$oSelection.setPropertyToDefault("CharBottomBorder") ; Resetting one truly resets all, but just to be sure, reset all.
		$oSelection.setPropertyToDefault("CharLeftBorder")
		$oSelection.setPropertyToDefault("CharRightBorder")
		If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_CharBorder($oSelection, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharEffect
; Description ...: Set or Retrieve the Font Effect settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharEffect(ByRef $oSelection[, $iCase = Null[, $bHidden = Null[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iCase               - [optional] (0-4) Default is Null. The Character Case Style. See Constants, $LOW_CHAR_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bHidden             - [optional] Default is Null. If True, the Characters are hidden.
;                  $iRelief             - [optional] (0-2) Default is Null. The Character Relief style. See Constants, $LOW_CHAR_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bOutline            - [optional] Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] Default is Null. If True, the characters have a shadow.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iCase not an Integer, less than 0 or greater than 4. See Constants, $LOW_CHAR_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 3 = $bHidden not a Boolean.
;                  @Error: 1, @Extended: 4 = $iRelief not an Integer, less than 0 or greater than 2. See Constants, $LOW_CHAR_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 5 = $bOutline not a Boolean.
;                  @Error: 1, @Extended: 6 = $bShadow not a Boolean.
;                  @Error: 1, @Extended: 7 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iCase
;                  |                               2 = Error setting $bHidden
;                  |                               4 = Error setting $iRelief
;                  |                               8 = Error setting $bOutline
;                  |                               16 = Error setting $bShadow
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting.
; Related .......: _LOWriter_DirFrmtCharOverLine, _LOWriter_DirFrmtCharStrikeOut, _LOWriter_DirFrmtCharUnderLine, _LOWriter_CharStyleEffect, _LOWriter_ParStyleEffect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharEffect(ByRef $oSelection, $iCase = Null, $bHidden = Null, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If __LOWriter_AnyAreDefault($iCase, $bHidden, $iRelief, $bOutline, $bShadow) Then
		If ($iCase = Default) Then
			$oSelection.setPropertyToDefault("CharCaseMap")
			$iCase = Null
		EndIf

		If ($bHidden = Default) Then
			$oSelection.setPropertyToDefault("CharHidden")
			$bHidden = Null
		EndIf

		If ($iRelief = Default) Then
			$oSelection.setPropertyToDefault("CharRelief")
			$iRelief = Null
		EndIf

		If ($bOutline = Default) Then
			$oSelection.setPropertyToDefault("CharContoured")
			$bOutline = Null
		EndIf

		If ($bShadow = Default) Then
			$oSelection.setPropertyToDefault("CharShadowed")
			$bShadow = Null
		EndIf

		If __LO_VarsAreNull($iCase, $bHidden, $iRelief, $bOutline, $bShadow) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharEffect($oSelection, $iCase, $bHidden, $iRelief, $bOutline, $bShadow)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharFont
; Description ...: Set and Retrieve the Font Settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharFont(ByRef $oSelection[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $sFontName           - [optional] Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] Default is Null. The new Font size.
;                  $iPosture            - [optional] (0-5) Default is Null. Font Italic setting. See Constants, $LOW_CHAR_POSTURE_* as defined in LibreOfficeWriter_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] (0, 50-200) Default is Null. Font Bold settings, see Constants, $LOW_CHAR_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3. Also see remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $sFontName not a String.
;                  @Error: 1, @Extended: 3 = $sFontName not available in current document.
;                  @Error: 1, @Extended: 4 = $nFontSize not a Number.
;                  @Error: 1, @Extended: 5 = $iPosture not an Integer, less than 0 or greater than 5. See Constants, $LOW_CHAR_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iWeight less than 50 and not 0, or more than 200. See Constants, $LOW_CHAR_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 7 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $nFontSize
;                  |                               4 = Error setting $iPosture
;                  |                               8 = Error setting $iWeight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting.
;                  Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;                  LibreOffice Writer accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......: _LOWriter_FontsGetNames, _LOWriter_DirFrmtCharFontColor, _LOWriter_CharStyleFont, _LOWriter_ParStyleFont
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharFont(ByRef $oSelection, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If __LOWriter_AnyAreDefault($sFontName, $nFontSize, $iPosture, $iWeight) Then
		If ($sFontName = Default) Then
			$oSelection.setPropertyToDefault("CharFontName")
			$sFontName = Null
		EndIf

		If ($nFontSize = Default) Then
			$oSelection.setPropertyToDefault("CharHeight")
			$nFontSize = Null
		EndIf

		If ($iPosture = Default) Then
			$oSelection.setPropertyToDefault("CharPosture")
			$iPosture = Null
		EndIf

		If ($iWeight = Default) Then
			$oSelection.setPropertyToDefault("CharWeight")
			$iWeight = Null
		EndIf

		If __LO_VarsAreNull($sFontName, $nFontSize, $iPosture, $iWeight) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharFont($oSelection, $sFontName, $nFontSize, $iPosture, $iWeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharFontColor(ByRef $oSelection[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iFontColor          - [optional] (-1-16777215) Default is Null. The Font color, as a RGB Color Integer, Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - [optional] (0-100) Default is Null. Transparency percentage. 0 is visible, 100 is invisible. Available for LibreOffice 7.0 and up.
;                  $iHighlight          - [optional] (-1-16777215) Default is Null. The highlight color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters. If The current LibreOffice version is below 7.0 the $iTransparency parameter will return a Null value.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 3 = $iTransparency not an Integer, less than 0 or greater than 100%.
;                  @Error: 1, @Extended: 4 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
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
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting. Font Color and Transparency reset at the same time as the other, e.g., if you reset Font Color, it will reset Transparency.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_DirFrmtCharFont, _LOWriter_CharStyleFontColor, _LOWriter_ParStyleFontColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharFontColor(ByRef $oSelection, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If __LOWriter_AnyAreDefault($iFontColor, $iTransparency, $iHighlight) Then
		If ($iFontColor = Default) Then
			$oSelection.setPropertyToDefault("CharColor")
			$iFontColor = Null
		EndIf

		If ($iTransparency = Default) Then
			If Not __LO_VersionCheck(7.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

			$oSelection.setPropertyToDefault("CharTransparence")
			$iTransparency = Null
		EndIf

		If ($iHighlight = Default) Then
			If __LO_VersionCheck(4.2) Then $oSelection.setPropertyToDefault("CharHighlight")
			$oSelection.setPropertyToDefault("CharBackColor") ; Both may be used? not sure. Both do the same thing, so reset both to make sure.
			$iHighlight = Null
		EndIf

		If __LO_VarsAreNull($iFontColor, $iTransparency, $iHighlight) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharFontColor($oSelection, $iFontColor, $iTransparency, $iHighlight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharOverLine
; Description ...: Set and retrieve the OverLine settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharOverLine(ByRef $oSelection[, $iOverLineStyle = Null[, $iOLColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iOverLineStyle      - [optional] (0-18) Default is Null. The style of the Overline line, see constants, $LOW_CHAR_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iOLColor            - [optional] (-1-16777215) Default is Null. The color of the Overline, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not Overlined.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iOverLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOW_CHAR_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iOLColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $bWordOnly not a Boolean.
;                  @Error: 1, @Extended: 5 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iOverLineStyle
;                  |                               2 = Error setting $iOLColor
;                  |                               4 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting. Overline style, and Color all reset together.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_DirFrmtCharEffect, _LOWriter_DirFrmtCharStrikeOut, _LOWriter_DirFrmtCharUnderLine, _LOWriter_CharStyleOverLine, _LOWriter_ParStyleOverLine
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharOverLine(ByRef $oSelection, $iOverLineStyle = Null, $iOLColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If __LOWriter_AnyAreDefault($iOverLineStyle, $iOLColor, $bWordOnly) Then
		If ($iOverLineStyle = Default) Then
			$oSelection.setPropertyToDefault("CharOverline")
			$iOverLineStyle = Null
		EndIf

		If ($iOLColor = Default) Then
			$oSelection.setPropertyToDefault("CharOverlineHasColor")
			$oSelection.setPropertyToDefault("CharOverlineColor")
			$iOLColor = Null
		EndIf

		If ($bWordOnly = Default) Then
			$oSelection.setPropertyToDefault("CharWordMode")
			$bWordOnly = Null
		EndIf

		If __LO_VarsAreNull($iOverLineStyle, $iOLColor, $bWordOnly) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharOverLine($oSelection, $iOverLineStyle, $iOLColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharPosition(ByRef $oSelection[, $iSuperScript = Null[, $iSubScript = Null[, $iRelativeSize = Null[, $bClearDirFrmt = False]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iSuperScript        - [optional] (-1-100) Default is Null. The Superscript percentage value. Call with -1 for Automatic SuperScript. See Remarks.
;                  $iSubScript          - [optional] (-1-100) Default is Null. Subscript percentage value. Call with -1 for Automatic SubScript. See Remarks.
;                  $iRelativeSize       - [optional] (1-100) Default is Null. The size percentage relative to current font size.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Super/Sub Script settings.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iSuperScript not an Integer, less than -1 or greater than 100.
;                  @Error: 1, @Extended: 3 = $iSubScript not an Integer, less than -1 or greater than 100.
;                  @Error: 1, @Extended: 4 = $iRelativeSize not an Integer, less than 1 or greater than 100.
;                  @Error: 1, @Extended: 5 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iSuperScript
;                  |                               2 = Error setting $iSubScript
;                  |                               4 = Error setting $iRelativeSize.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Set either $iSubScript or $iSuperScript to 0 to return it to Normal setting.
;                  The way LibreOffice is set up Super/Subscript are set in the same setting, Superscript is a positive number from 1 to 100 (percentage), Subscript is a negative number set to -1 to -100 percentage. For the user's convenience this function automatically converts the positive numbers to negative, and back when setting or retrieving subscript values.
;                  Automatic Superscript has an Integer value of 14000, Auto Subscript has a Integer value of -14000. Being that there is no settable setting of Automatic Super/Sub Script, it has been chosen to use -1 to indicate an automatic Sub/SuperScript value.
;                  If you set both $iSuperScript and $iSubScript to -1 (Automatic), or both $iSuperScript and $iSubScript to any value, Subscript will be the result, as it is the last in the function to be set, and thus will overwrite any Superscript values.
; Related .......: _LOWriter_DirFrmtCharRotateScale, _LOWriter_DirFrmtCharSpacing, _LOWriter_CharStylePosition, _LOWriter_ParStylePosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharPosition(ByRef $oSelection, $iSuperScript = Null, $iSubScript = Null, $iRelativeSize = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharEscapement")
		If __LO_VarsAreNull($iSuperScript, $iSubScript, $iRelativeSize) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharPosition($oSelection, $iSuperScript, $iSubScript, $iRelativeSize)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharRotateScale
; Description ...: Set or retrieve the character rotational and Scale settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharRotateScale(ByRef $oSelection[, $iRotation = Null[, $iScaleWidth = Null[, $bRotateFitLine = Null]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iRotation           - [optional] (0, 90, 270) Default is Null. Degrees to rotate the text.
;                  $iScaleWidth         - [optional] (1-100) Default is Null. The percentage to horizontally stretch or compress the text. 100 is normal sizing.
;                  $bRotateFitLine      - [optional] Default is Null. If True, Stretches or compresses the selected text so that it fits between the line that is above the text and the line that is below the text.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iRotation not an Integer or not equal to 0, 90 or 270 degrees.
;                  @Error: 1, @Extended: 3 = $iScaleWidth not an Integer or less than 1% or greater than 100%.
;                  @Error: 1, @Extended: 4 = $bRotateFitLine not a Boolean.
;                  @Error: 1, @Extended: 5 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iRotation
;                  |                               2 = Error setting $iScaleWidth
;                  |                               4 = Error setting $bRotateFitLine
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting.
; Related .......: _LOWriter_DirFrmtCharPosition, _LOWriter_DirFrmtCharSpacing, _LOWriter_CharStyleRotateScale, _LOWriter_ParStyleRotateScale
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharRotateScale(ByRef $oSelection, $iRotation = Null, $iScaleWidth = Null, $bRotateFitLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If __LOWriter_AnyAreDefault($iRotation, $iScaleWidth, $bRotateFitLine) Then
		If ($iRotation = Default) Then
			$oSelection.setPropertyToDefault("CharRotation")
			$iRotation = Null
		EndIf

		If ($iScaleWidth = Default) Then
			$oSelection.setPropertyToDefault("CharScaleWidth")
			$iScaleWidth = Null
		EndIf

		If ($bRotateFitLine = Default) Then
			$oSelection.setPropertyToDefault("CharRotationIsFitToLine")
			$bRotateFitLine = Null
		EndIf

		If __LO_VarsAreNull($iRotation, $iScaleWidth, $bRotateFitLine) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	If __LO_VarsAreNull($iRotation, $iScaleWidth, $bRotateFitLine) Then
		$vReturn = __LOWriter_CharRotateScale($oSelection, $iRotation, $iScaleWidth, $bRotateFitLine)
		__LO_AddTo1DArray($vReturn, $oSelection.CharRotationIsFitToLine())

		Return SetError($__LO_STATUS_SUCCESS, 1, $vReturn)
	EndIf

	$vReturn = __LOWriter_CharRotateScale($oSelection, $iRotation, $iScaleWidth, $bRotateFitLine)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharRotateScale

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharShadow
; Description ...: Set and retrieve the Shadow for a Character Style by Direct Formatting. LibreOffice 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharShadow(ByRef $oSelection[, $iLocation = Null[, $iColor = Null[, $iWidth = Null[, $bClearDirFrmt = False]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iLocation           - [optional] (0-4) Default is Null. Location of the shadow compared to the characters. See Constants, $LOW_SHADOW_LOCATION_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iColor              - [optional] (0-16777215) Default is Null. Color of the shadow, as a RGB Color Integer. Can be a custom value or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. See remarks.
;                  $iWidth              - [optional] Default is Null. Width of the shadow, set in Hundredths of a Millimeter (HMM).
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Character Shadow, Width, Color and Location.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOW_SHADOW_LOCATION_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 5 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Shadow format Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Shadow format Object for Error checking.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLocation
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $iWidth
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 4.2.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  LibreOffice may adjust the set width +/- 1 Hundredth of a Millimeter (HMM) after setting.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert, _LOWriter_DirFrmtCharEffect, _LOWriter_DirFrmtCharFontColor, _LOWriter_DirFrmtParShadow, _LOWriter_CharStyleShadow, _LOWriter_ParStyleShadow
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharShadow(ByRef $oSelection, $iLocation = Null, $iColor = Null, $iWidth = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharShadowFormat")
		If __LO_VarsAreNull($iLocation, $iColor, $iWidth) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharShadow($oSelection, $iLocation, $iColor, $iWidth)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning)by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharSpacing(ByRef $oSelection[, $bAutoKerning = Null[, $nKerning = Null]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $bAutoKerning        - [optional] Default is Null. If True, applies a spacing in between certain pairs of characters.
;                  $nKerning            - [optional] (-2-928.8) Default is Null. The kerning value of the characters. See Remarks. Values are in Printer's Points as set in the LibreOffice UI.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $bAutoKerning not a Boolean.
;                  @Error: 1, @Extended: 3 = $nKerning not a number, less than -2 or greater than 928.8 Points.
;                  @Error: 1, @Extended: 4 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoKerning
;                  |                               2 = Error setting $nKerning.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting.
;                  When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User Display, however the internal setting is measured in Hundredths of a Millimeter (HMM). They will be automatically converted from Points to Hundredths of a Millimeter and back for retrieval of settings.
;                  The acceptable values are from -2 Pt to 928.8 Pt. the figures can be directly converted easily, however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative Hundredths of a Millimeter internally from 928.9 up to 1000 Pt (Max setting).
;                  For example, 928.8Pt is the last correct value, which equals 32766 Hundredths of a Millimeter (HMM), after this LibreOffice reports the following: ;928.9 Pt = -32766 (HMM); 929 Pt = -32763 (HMM); 929.1 = -32759; 1000 pt = -30258.
;                  Attempting to set LibreOffice's kerning value to anything over 32768 (HMM) causes a COM exception, and attempting to set the kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the max settable kerning is -2.0 Pt to 928.8 Pt.
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtCharRotateScale, _LOWriter_DirFrmtParSpace, _LOWriter_CharStyleSpacing, _LOWriter_ParStyleSpacing
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharSpacing(ByRef $oSelection, $bAutoKerning = Null, $nKerning = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If __LOWriter_AnyAreDefault($bAutoKerning, $nKerning) Then
		If ($bAutoKerning = Default) Then
			$oSelection.setPropertyToDefault("CharAutoKerning")
			$bAutoKerning = Null
		EndIf

		If ($nKerning = Default) Then
			$oSelection.setPropertyToDefault("CharKerning")
			$nKerning = Null
		EndIf
		If __LO_VarsAreNull($bAutoKerning, $nKerning) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharSpacing($oSelection, $bAutoKerning, $nKerning)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharStrikeOut
; Description ...: Set or Retrieve the StrikeOut settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharStrikeOut(ByRef $oSelection[, $iStrikeLineStyle = Null[, $bWordOnly = Null]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iStrikeLineStyle    - [optional] (0-6) Default is Null. The Strikeout Line Style, see constants, $LOW_CHAR_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bWordOnly           - [optional] Default is Null. If True, strikes out words only and skip whitespaces.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iStrikeLineStyle not an Integer, less than 0 or greater than 6. See constants, $LOW_CHAR_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 3 = $bWordOnly not a Boolean.
;                  @Error: 1, @Extended: 4 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iStrikeLineStyle
;                  |                               2 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting.
;                  Strikeout line style is converted to a single line in Ms word document format.
; Related .......: _LOWriter_DirFrmtCharEffect, _LOWriter_DirFrmtCharOverLine, _LOWriter_DirFrmtCharUnderLine, _LOWriter_CharStyleStrikeOut
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharStrikeOut(ByRef $oSelection, $iStrikeLineStyle = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If __LOWriter_AnyAreDefault($iStrikeLineStyle, $bWordOnly) Then
		If ($iStrikeLineStyle = Default) Then
			$oSelection.setPropertyToDefault("CharCrossedOut")
			$oSelection.setPropertyToDefault("CharStrikeout")
			$iStrikeLineStyle = Null
		EndIf

		If ($bWordOnly = Default) Then
			$oSelection.setPropertyToDefault("CharWordMode")
			$bWordOnly = Null
		EndIf

		If __LO_VarsAreNull($bWordOnly, $iStrikeLineStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharStrikeOut($oSelection, $iStrikeLineStyle, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharUnderLine
; Description ...: Set and retrieve the Underline settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharUnderLine(ByRef $oSelection[, $iUnderLineStyle = Null[, $iULColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
;                  $iUnderLineStyle     - [optional] (0-18) Default is Null. The style of the Underline line, see constants, $LOW_CHAR_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iULColor            - [optional] (-1-16777215) Default is Null. The color of the underline, as a RGB Color Integer. Can be a custom value or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not underlined.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iUnderLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOW_CHAR_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iULColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $bWordOnly not a Boolean.
;                  @Error: 1, @Extended: 5 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iUnderLineStyle
;                  |                               2 = Error setting $iULColor
;                  |                               4 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting. Underline style, and Color all reset together.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_DirFrmtCharEffect, _LOWriter_DirFrmtCharOverLine, _LOWriter_DirFrmtCharStrikeOut, _LOWriter_CharStyleUnderLine
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharUnderLine(ByRef $oSelection, $iUnderLineStyle = Null, $iULColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If __LOWriter_AnyAreDefault($iUnderLineStyle, $iULColor, $bWordOnly) Then
		If ($iUnderLineStyle = Default) Then
			$oSelection.setPropertyToDefault("CharUnderline")
			$iUnderLineStyle = Null
		EndIf

		If ($iULColor = Default) Then
			$oSelection.setPropertyToDefault("CharUnderlineHasColor")
			$oSelection.setPropertyToDefault("CharUnderlineColor")
			$iULColor = Null
		EndIf

		If ($bWordOnly = Default) Then
			$oSelection.setPropertyToDefault("CharWordMode")
			$bWordOnly = Null
		EndIf

		If __LO_VarsAreNull($iUnderLineStyle, $iULColor, $bWordOnly) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharUnderLine($oSelection, $iUnderLineStyle, $iULColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharUnderLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtClear
; Description ...: Clear any Direct formatting in a Cursor or Text Object.
; Syntax ........: _LOWriter_DirFrmtClear(ByRef $oDoc, ByRef $oSelection)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval function, Or A Paragraph Object, or other Object containing a selection of text.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Direct Formatting was successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not an Object.
;                  @Error: 1, @Extended: 3 = $oSelection does not support any of the following: "com.sun.star.text.Paragraph"; "TextPortion"; "TextCursor"; "TextViewCursor".
;                  @Error: 1, @Extended: 4 = $oSelection is a Table Cursor, which is not supported.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to determine $oSelection's cursor type.
;                  @Error: 3, @Extended: 2 = Failed to backup Viewcursor's position.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function causes the ViewCursor to select the data input in $oSelection, unless $oSelection is a ViewCursor object. After the formatting has been cleared the ViewCursor is returned to its previous position.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtClear(ByRef $oDoc, ByRef $oSelection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aArray[0]
	Local $oServiceManager, $oDispatcher, $oBackupSelection
	Local $iCursorType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oSelection)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Switch $iCursorType
		Case $LOW_CURTYPE_TEXT_CURSOR, $LOW_CURTYPE_PARAGRAPH, $LOW_CURTYPE_TEXT_PORTION
			; Backup the ViewCursor location and selection.
			$oBackupSelection = $oDoc.getCurrentSelection()
			If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oDoc.CurrentController.Select($oSelection)

			$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ResetAttributes", "", 0, $aArray)

			; Restore the ViewCursor to its previous location.
			$oDoc.CurrentController.Select($oBackupSelection)

		Case $LOW_CURTYPE_VIEW_CURSOR
			$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ResetAttributes", "", 0, $aArray)
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DirFrmtClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtGetCurStyles
; Description ...: Retrieve the current Styles set for a selection of text.
; Syntax ........: _LOWriter_DirFrmtGetCurStyles(ByRef $oSelection)
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions that has data selected. Or a paragraph or paragraph section.
; Return values .: Success: Array
;                  @Error: 0, @Extended: 0, Return: Array = Success. Returning a 4 element array in the following order: Paragraph Style Name, Character Style Name, Page Style Name, Numbering Style Name. See Remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection does not support Paragraph Properties service.
;                  @Error: 1, @Extended: 3 = $oSelection does not support Character Properties service.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Some of the returned style values may be blank if they are not set, particularly Numbering style.
; Related .......: _LOWriter_ParStyleCurrent, _LOWriter_CharStyleCurrent, _LOWriter_PageStyleCurrent, _LOWriter_NumStyleCurrent
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtGetCurStyles(ByRef $oSelection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asStyles[4]

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSelection.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSelection.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	__LO_ArrayFill($asStyles, $oSelection.ParaStyleName(), $oSelection.CharStyleName(), $oSelection.PageStyleName(), $oSelection.NumberingStyleName())

	Return SetError($__LO_STATUS_SUCCESS, 0, $asStyles)
EndFunc   ;==>_LOWriter_DirFrmtGetCurStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParAlignment
; Description ...: Set and Retrieve Alignment settings for a paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParAlignment(ByRef $oSelection[, $iHorAlign = Null[, $iVertAlign = Null[, $iLastLineAlign = Null[, $bExpandSingleWord = Null[, $bSnapToGrid = Null[, $iTxtDirection = Null]]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iHorAlign           - [optional] (0-3) Default is Null. The Horizontal alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iVertAlign          - [optional] (0-4) Default is Null. The Vertical alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLastLineAlign      - [optional] (0-3) Default is Null. Specify the alignment for the last line in the paragraph. See Constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $bExpandSingleWord   - [optional] Default is Null. If True, and the last line of a justified paragraph consists of one word, the word is stretched to the width of the paragraph.
;                  $bSnapToGrid         - [optional] Default is Null. If True, Aligns the paragraph to a text grid (if one is active).
;                  $iTxtDirection       - [optional] (0-5) Default is Null. The Text Writing Direction. See Constants, $LOW_PAR_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3. [LibreOffice Default is 4]
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iHorAlign not an Integer, less than 0 or greater than 3. See Constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iVertAlign not an Integer, less than 0 or greater than 4. See Constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iLastLineAlign not an Integer, less than 0 or greater than 3. See Constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 5 = $bExpandSingleWord not a Boolean.
;                  @Error: 1, @Extended: 6 = $bSnapToGrid not a Boolean.
;                  @Error: 1, @Extended: 7 = $iTxtDirection not an Integer, less than 0 or greater than 5. See constants, $LOW_PAR_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 8 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iHorAlign
;                  |                               2 = Error setting $iVertAlign
;                  |                               4 = Error setting $iLastLineALign
;                  |                               8 = Error setting $bExpandSIngleWord
;                  |                               16 = Error setting $bSnapToGrid
;                  |                               32 = Error setting $iTxtDirection
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting. $iHorAlign, $iLastLineAlign, and $bExpandSingleWord are all reset together.
;                  $iHorAlign must be set to $LOW_PAR_ALIGN_HOR_JUSTIFIED(2) before you can set $iLastLineAlign, and $iLastLineAlign must be set to $LOW_PAR_LAST_LINE_JUSTIFIED(2) before $bExpandSingleWord can be set.
; Related .......: _LOWriter_DirFrmtParIndent, _LOWriter_DirFrmtParSpace, _LOWriter_DirFrmtParTxtFlowOpt, _LOWriter_ParStyleAlignment
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParAlignment(ByRef $oSelection, $iHorAlign = Null, $iVertAlign = Null, $iLastLineAlign = Null, $bExpandSingleWord = Null, $bSnapToGrid = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If __LOWriter_AnyAreDefault($iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection) Then
		If ($iHorAlign = Default) Then
			$oSelection.setPropertyToDefault("ParaAdjust")
			$iHorAlign = Null
		EndIf

		If ($iVertAlign = Default) Then
			$oSelection.setPropertyToDefault("ParaVertAlignment")
			$iVertAlign = Null
		EndIf

		If ($iLastLineAlign = Default) Then
			$oSelection.setPropertyToDefault("ParaLastLineAdjust")
			$iLastLineAlign = Null
		EndIf

		If ($bExpandSingleWord = Default) Then
			$oSelection.setPropertyToDefault("ParaExpandSingleWord")
			$bExpandSingleWord = Null
		EndIf

		If ($bSnapToGrid = Default) Then
			$oSelection.setPropertyToDefault("SnapToGrid")
			$bSnapToGrid = Null
		EndIf

		If ($iTxtDirection = Default) Then
			$oSelection.setPropertyToDefault("WritingMode")
			$iTxtDirection = Null
		EndIf

		If __LO_VarsAreNull($iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParAlignment($oSelection, $iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParAreaColor
; Description ...: Set or Retrieve background color settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParAreaColor(ByRef $oSelection[, $iBackColor = Null[, $bClearDirFrmt = False]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iBackColor          - [optional] (-1-16777215) Default is Null. The background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) to turn Background color off.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Background color.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting as an Integer.
;                  @Error: 0, @Extended: 2, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 3 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current background color.
;                  @Error: 3, @Extended: 2 = Failed to retrieve TextParagraph Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_DirFrmtCharFontColor, _LOWriter_DirFrmtParAreaFillStyle, _LOWriter_DirFrmtParAreaGradient, _LOWriter_ParStyleAreaColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParAreaColor(ByRef $oSelection, $iBackColor = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oTxtPar

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTxtPar = $oSelection.TextParagraph()
	If Not IsObj($oTxtPar) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oTxtPar.setPropertyToDefault("FillColor")
		If __LO_VarsAreNull($iBackColor) Then Return SetError($__LO_STATUS_SUCCESS, 2, 2)
	EndIf

	$vReturn = __LOWriter_ParAreaColor($oTxtPar, $iBackColor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParAreaFillStyle
; Description ...: Retrieve what kind of background fill is active by Direct Formatting, if any.
; Syntax ........: _LOWriter_DirFrmtParAreaFillStyle(ByRef $oSelection)
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOW_AREA_FILL_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Fill Style.
;                  @Error: 3, @Extended: 2 = Failed to retrieve TextParagraph Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
;                  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  If an entire Paragraph is not selected in the selection, a wrong value may also result, and setting may not apply correctly.
; Related .......: _LOWriter_DirFrmtParAreaGradient, _LOWriter_DirFrmtParAreaColor, _LOWriter_ParStyleAreaFillStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParAreaFillStyle(ByRef $oSelection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oTxtPar

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTxtPar = $oSelection.TextParagraph()
	If Not IsObj($oTxtPar) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$vReturn = __LOWriter_ParAreaFillStyle($oTxtPar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParAreaGradient
; Description ...: Set or retrieve the Direct Formatting settings for Paragraph Background color Gradient.
; Syntax ........: _LOWriter_DirFrmtParAreaGradient(ByRef $oDoc, ByRef $oSelection[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null[, $bClearDirFrmt = False]]]]]]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $sGradientName       - [optional] Default is Null. A Preset Gradient Name. See remarks. See constants, $LOW_GRAD_NAME_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iType               - [optional] (-1-5) Default is Null. The gradient type to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iIncrement          - [optional] (0, 3-256) Default is Null. The number of steps of color change. 0 = Automatic.
;                  $iXCenter            - [optional] (0-100) Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] (0-100) Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] (0-359) Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] (0-100) Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] (0-16777215) Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] (0-16777215) Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] (0-100) Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] (0-100) Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Background color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. Gradient has been successfully turned off.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 3 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not an Object.
;                  @Error: 1, @Extended: 3 = $sGradientName not a String.
;                  @Error: 1, @Extended: 4 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error: 1, @Extended: 6 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 9 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 10 = $iFromColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 11 = $iToColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 12 = $iFromIntense not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 13 = $iToIntense not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 14 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillGradient" Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Error creating Gradient Name.
;                  @Error: 3, @Extended: 4 = Error setting Gradient Name.
;                  @Error: 3, @Extended: 5 = Failed to retrieve TextParagraph Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sGradientName
;                  |                               2 = Error setting $iType
;                  |                               4 = Error setting $iIncrement
;                  |                               8 = Error setting $iXCenter
;                  |                               16 = Error setting $iYCenter
;                  |                               32 = Error setting $iAngle
;                  |                               64 = Error setting $iTransitionStart
;                  |                               128 = Error setting $iFromColor
;                  |                               256 = Error setting $iToColor
;                  |                               512 = Error setting $iFromIntense
;                  |                               1024 = Error setting $iToIntense
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
;                  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  If an entire Paragraph is not selected in the selection, a wrong value may also result, and setting may not apply correctly.
; Related .......: _LOWriter_DirFrmtParAreaColor, _LOWriter_DirFrmtParAreaGradientMulticolor, _LOWriter_DirFrmtParAreaFillStyle, _LOWriter_ParStyleAreaGradient, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParAreaGradient(ByRef $oDoc, ByRef $oSelection, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oTxtPar

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

	$oTxtPar = $oSelection.TextParagraph()
	If Not IsObj($oTxtPar) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	If $bClearDirFrmt Then
		$oTxtPar.setPropertyToDefault("FillGradient")
		$oTxtPar.setPropertyToDefault("FillGradientName")
		$oTxtPar.setPropertyToDefault("FillGradientStepCount")
		If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then Return SetError($__LO_STATUS_SUCCESS, 0, 3)
	EndIf

	$vReturn = __LOWriter_ParAreaGradient($oDoc, $oTxtPar, $sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParAreaGradientMulticolor
; Description ...: Set or retrieve a Paragraph's Direct Formatting Multicolor Gradient settings.
; Syntax ........: _LOWriter_DirFrmtParAreaGradientMulticolor(ByRef $oSelection[, $avColorStops = Null])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $avColorStops        - [optional] Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: ?, Return: Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $avColorStops not an Array, or does not contain two columns.
;                  @Error: 1, @Extended: 3 = $avColorStops contains less than two rows.
;                  @Error: 1, @Extended: 4 = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error: 1, @Extended: 5 = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  @Error: 1, @Extended: 6 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve FillGradient Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Failed to retrieve StopColor Struct.
;                  @Error: 3, @Extended: 4 = Failed to retrieve TextParagraph Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current version less than 7.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, as a RGB Color Integer.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  If an entire Paragraph is not selected in the selection, a wrong value may also result, and setting may not apply correctly.
; Related .......: _LO_GradientMulticolorAdd, _LO_GradientMulticolorDelete, _LO_GradientMulticolorModify, _LOWriter_DirFrmtParAreaGradient, _LOWriter_ParStyleAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParAreaGradientMulticolor(ByRef $oSelection, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oTxtPar

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oTxtPar = $oSelection.TextParagraph()
	If Not IsObj($oTxtPar) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$vReturn = __LOWriter_ParAreaGradientMulticolor($oTxtPar, $avColorStops)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParAreaTransparency
; Description ...: Set or retrieve Transparency settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParAreaTransparency(ByRef $oSelection[, $iTransparency = Null[, $bClearDirFrmt = False]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iTransparency       - [optional] (0-100) Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Background color.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 3 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Transparency value.
;                  @Error: 3, @Extended: 2 = Failed to retrieve TextParagraph Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  If an entire Paragraph is not selected in the selection, a wrong value may also result, and setting may not apply correctly.
; Related .......: _LOWriter_DirFrmtParAreaColor, _LOWriter_ParStyleAreaTransparency
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParAreaTransparency(ByRef $oSelection, $iTransparency = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oTxtPar

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTxtPar = $oSelection.TextParagraph()
	If Not IsObj($oTxtPar) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oTxtPar.setPropertyToDefault("FillTransparence")
		If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParAreaTransparency($oTxtPar, $iTransparency)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParBorderColor
; Description ...: Set and Retrieve the Paragraph Style Border Line Color. LibreOffice Version 3.4 and Up.
; Syntax ........: _LOWriter_DirFrmtParBorderColor(ByRef $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iTop                - [optional] (0-16777215) Default is Null. The Top Border Line Color of the Paragraph Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] (0-16777215) Default is Null. The Bottom Border Line Color of the Paragraph Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] (0-16777215) Default is Null. The Left Border Line Color of the Paragraph Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] (0-16777215) Default is Null. The Right Border Line Color of the Paragraph Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of the Paragraph Border, Width, Style and Color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  @Error: 1, @Extended: 3 = $iTop not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $iBottom not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iLeft not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 6 = $iRight not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error: 3, @Extended: 2 = Cannot set Top Border Color when Top Border width not set.
;                  @Error: 3, @Extended: 3 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error: 3, @Extended: 4 = Cannot set Left Border Color when Left Border width not set.
;                  @Error: 3, @Extended: 5 = Cannot set Right Border Color when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.4.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Border Width must be set first to be able to set Border Style and Color.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_DirFrmtParBorderWidth, _LOWriter_DirFrmtParBorderStyle, _LOWriter_DirFrmtParBorderPadding, _LOWriter_DirFrmtCharBorderColor, _LOWriter_ParStyleBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParBorderColor(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("TopBorder")
		$oSelection.setPropertyToDefault("BottomBorder")
		$oSelection.setPropertyToDefault("LeftBorder")
		$oSelection.setPropertyToDefault("RightBorder")

		If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oSelection, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParBorderPadding
; Description ...: Set or retrieve the Border Padding (spacing between the Paragraph and border) settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParBorderPadding(ByRef $oSelection[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iAll                - [optional] Default is Null. Set all four padding distances to one distance in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] Default is Null. The Top Distance between the Border and Paragraph in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] Default is Null. The Bottom Distance between the Border and Paragraph in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] Default is Null. The Left Distance between the Border and Paragraph in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] Default is Null. The Right Distance between the Border and Paragraph in Hundredths of a Millimeter (HMM).
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Border padding related settings.
; Return values .: Success: Integer or Array, see Remarks.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iAll not an Integer.
;                  @Error: 1, @Extended: 3 = $iTop not an Integer.
;                  @Error: 1, @Extended: 4 = $iBottom not an Integer.
;                  @Error: 1, @Extended: 5 = $Left not an Integer.
;                  @Error: 1, @Extended: 6 = $iRight not an Integer.
;                  @Error: 1, @Extended: 7 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAll border distance
;                  |                               2 = Error setting $iTop border distance
;                  |                               4 = Error setting $iBottom border distance
;                  |                               8 = Error setting $iLeft border distance
;                  |                               16 = Error setting $iRight border distance
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtParBorderWidth, _LOWriter_DirFrmtParBorderStyle, _LOWriter_DirFrmtParBorderColor, _LOWriter_DirFrmtCharBorderPadding, _LOWriter_ParStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParBorderPadding(ByRef $oSelection, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("BorderDistance")
		If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParBorderPadding($oSelection, $iAll, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParBorderStyle
; Description ...: Set and retrieve the Paragraph Border Line style by Direct Formatting. LibreOffice Version 3.4 and Up.
; Syntax ........: _LOWriter_DirFrmtParBorderStyle(ByRef $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iTop                - [optional] (0x7FFF,0-17) Default is Null. The Top Border Line Style of the Paragraph Style. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] (0x7FFF,0-17) Default is Null. The Bottom Border Line Style of the Paragraph Style. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] (0x7FFF,0-17) Default is Null. The Left Border Line Style of the Paragraph Style. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] (0x7FFF,0-17) Default is Null. The Right Border Line Style of the Paragraph Style. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of the Paragraph Border, Width, Style and Color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  @Error: 1, @Extended: 3 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOW_BORDER_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error: 3, @Extended: 2 = Cannot set Top Border Style when Top Border width not set.
;                  @Error: 3, @Extended: 3 = Cannot set Bottom Border Style when Bottom Border width not set.
;                  @Error: 3, @Extended: 4 = Cannot set Left Border Style when Left Border width not set.
;                  @Error: 3, @Extended: 5 = Cannot set Right Border Style when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.4.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Border Width must be set first to be able to set Border Style and Color.
; Related .......: _LOWriter_DirFrmtParBorderWidth, _LOWriter_DirFrmtParBorderColor, _LOWriter_DirFrmtParBorderPadding, _LOWriter_DirFrmtCharBorderStyle, _LOWriter_ParStyleBorderStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParBorderStyle(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("TopBorder")
		$oSelection.setPropertyToDefault("BottomBorder")
		$oSelection.setPropertyToDefault("LeftBorder")
		$oSelection.setPropertyToDefault("RightBorder")

		If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOW_BORDER_STYLE_SOLID, $LOW_BORDER_STYLE_DASH_DOT_DOT, "", $LOW_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOW_BORDER_STYLE_SOLID, $LOW_BORDER_STYLE_DASH_DOT_DOT, "", $LOW_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOW_BORDER_STYLE_SOLID, $LOW_BORDER_STYLE_DASH_DOT_DOT, "", $LOW_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOW_BORDER_STYLE_SOLID, $LOW_BORDER_STYLE_DASH_DOT_DOT, "", $LOW_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oSelection, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParBorderWidth
; Description ...: Set and retrieve the Paragraph Border Line Width, or the Paragraph Connect Border option by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParBorderWidth(ByRef $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bConnectBorder = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iTop                - [optional] Default is Null. The Top Border Line width of the Paragraph in Hundredths of a Millimeter (HMM). Can be a custom value of one of the constants, $LOW_BORDER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3. LibreOffice Version 3.4 and Up.
;                  $iBottom             - [optional] Default is Null. The Bottom Border Line Width of the Paragraph in Hundredths of a Millimeter (HMM). Can be a custom value of one of the constants, $LOW_BORDER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3. LibreOffice Version 3.4 and Up.
;                  $iLeft               - [optional] Default is Null. The Left Border Line width of the Paragraph in Hundredths of a Millimeter (HMM). Can be a custom value of one of the constants, $LOW_BORDER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3. LibreOffice Version 3.4 and Up.
;                  $iRight              - [optional] Default is Null. The Right Border Line Width of the Paragraph in Hundredths of a Millimeter (HMM). Can be a custom value of one of the constants, $LOW_BORDER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3. LibreOffice Version 3.4 and Up.
;                  $bConnectBorder      - [optional] Default is Null. If True, borders set for a paragraph are merged with the next paragraph. Note: Borders are only merged if they are identical. LibreOffice Version 3.4 and Up.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of the Paragraph Border, Width, Style and Color. Doesn't clear $bConnectBorder. See Remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  @Error: 0, @Extended: 0, Return: 3 = Success. $bConnectBorder parameter was called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  @Error: 1, @Extended: 3 = $iTop not an Integer, or less than 0.
;                  @Error: 1, @Extended: 4 = $iBottom not an Integer, or less than 0.
;                  @Error: 1, @Extended: 5 = $iLeft not an Integer, or less than 0.
;                  @Error: 1, @Extended: 6 = $iRight not an Integer, or less than 0.
;                  @Error: 1, @Extended: 7 = $bConnectBorder not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  |                               16 = Error setting $bConnectBorder
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.4.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call $bConnectBorder Parameter with Default keyword to clear direct formatting for that setting.
;                  To "Turn Off" Borders, set them to 0
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtParBorderStyle, _LOWriter_DirFrmtParBorderColor, _LOWriter_DirFrmtParBorderPadding, _LOWriter_DirFrmtCharBorderWidth, _LOWriter_ParStyleBorderWidth
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParBorderWidth(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bConnectBorder = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("TopBorder")
		$oSelection.setPropertyToDefault("BottomBorder")
		$oSelection.setPropertyToDefault("LeftBorder")
		$oSelection.setPropertyToDefault("RightBorder")

		If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $bConnectBorder) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($bConnectBorder = Default) Then
		$oSelection.setPropertyToDefault("ParaIsConnectBorder")
		$bConnectBorder = Null
		If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $bConnectBorder) Then Return SetError($__LO_STATUS_SUCCESS, 0, 3)
	EndIf

	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($bConnectBorder <> Null) And Not IsBool($bConnectBorder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $bConnectBorder) Then
		$vReturn = __LOWriter_Border($oSelection, True, False, False, $iTop, $iBottom, $iLeft, $iRight)
		__LO_AddTo1DArray($vReturn, $oSelection.ParaIsConnectBorder())

		Return SetError($__LO_STATUS_SUCCESS, 1, $vReturn)

	ElseIf Not __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		$vReturn = __LOWriter_Border($oSelection, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

		If @error Then Return SetError(@error, @extended, $vReturn)
	EndIf
	If ($bConnectBorder <> Null) Then
		$oSelection.ParaIsConnectBorder = $bConnectBorder

		If ($oSelection.ParaIsConnectBorder() <> $bConnectBorder) Then
			If (@error = $__LO_STATUS_PROP_SETTING_ERROR) Then
				SetError(@error, BitOR(@extended, 16), $vReturn)

			Else
				SetError($__LO_STATUS_PROP_SETTING_ERROR, BitOR(0, 16), $vReturn)
			EndIf
		EndIf
	EndIf

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParDropCaps
; Description ...: Set or Retrieve DropCaps settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParDropCaps(ByRef $oDoc, ByRef $oSelection[, $iNumChar = Null[, $iLines = Null[, $iSpcTxt = Null[, $bWholeWord = Null[, $sCharStyle = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iNumChar            - [optional] (0-9) Default is Null. The number of characters to make into DropCaps.
;                  $iLines              - [optional] (0, 2-9) Default is Null. The number of lines to drop down.
;                  $iSpcTxt             - [optional] Default is Null. The distance between the drop cap and the following text. In Hundredths of a Millimeter (HMM).
;                  $bWholeWord          - [optional] Default is Null. If True, DropCap the whole first word. (Nullifys $iNumChar.)
;                  $sCharStyle          - [optional] Default is Null. The character style to use for the DropCaps. See Remarks.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of DropCaps and related settings.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not an Object.
;                  @Error: 1, @Extended: 3 = $iNumChar not an Integer, less than 0 or greater than 9.
;                  @Error: 1, @Extended: 4 = $iLines not an Integer, less than 0, equal to 1 or greater than 9
;                  @Error: 1, @Extended: 5 = $iSpaceTxt not an Integer, or less than 0.
;                  @Error: 1, @Extended: 6 = $bWholeWord not a Boolean.
;                  @Error: 1, @Extended: 7 = $sCharStyle not a String.
;                  @Error: 1, @Extended: 8 = Character Style called in $sCharStyle not found in current document.
;                  @Error: 1, @Extended: 9 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving DropCap Format Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iNumChar
;                  |                               2 = Error setting $iLines
;                  |                               4 = Error setting $iSpcTxt
;                  |                               8 = Error setting $bWholeWord
;                  |                               16 = Error setting $sCharStyle
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Set $iNumChars, $iLines, $iSpcTxt to 0 to disable DropCaps.
;                  I am unable to find a way to set Drop Caps character style to "None" as is available in the User Interface. When it is set to "None" LibreOffice returns a blank string ("") but setting it to a blank string throws a COM error/Exception. Consequently, you cannot set Character Style to "None", but you can still disable Drop Caps as noted above.
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtParOutLineAndList, _LOWriter_ParStyleDropCaps
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParDropCaps(ByRef $oDoc, ByRef $oSelection, $iNumChar = Null, $iLines = Null, $iSpcTxt = Null, $bWholeWord = Null, $sCharStyle = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("DropCapFormat")
		If __LO_VarsAreNull($iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParDropCaps($oDoc, $oSelection, $iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParDropCaps

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParHyphenation
; Description ...: Set or Retrieve Hyphenation settings for a paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParHyphenation(ByRef $oSelection[, $bAutoHyphen = Null[, $bHyphenNoCaps = Null[, $iMaxHyphens = Null[, $iMinLeadingChar = Null[, $iMinTrailingChar = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $bAutoHyphen         - [optional] Default is Null. If True, automatic hyphenation is applied.
;                  $bHyphenNoCaps       - [optional] Default is Null. Setting to True will disable hyphenation of words written in CAPS for this paragraph. LibreOffice 6.4 and up.
;                  $iMaxHyphens         - [optional] (0-99) Default is Null. The maximum number of consecutive hyphens.
;                  $iMinLeadingChar     - [optional] (2-9) Default is Null. Specifies the minimum number of characters to remain before the hyphen character (when hyphenation is applied).
;                  $iMinTrailingChar    - [optional] (2-9) Default is Null. Specifies the minimum number of characters to remain after the hyphen character (when hyphenation is applied).
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Hyphenation.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters. If the current LibreOffice Version is below 6.4, the $bHyphenNoCaps parameter will return a Null value.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $bAutoHyphen not a Boolean.
;                  @Error: 1, @Extended: 3 = $bHyphenNoCaps not a Boolean.
;                  @Error: 1, @Extended: 4 = $iMaxHyphens not an Integer, less than 0 or greater than 99.
;                  @Error: 1, @Extended: 5 = $iMinLeadingChar not an Integer, less than 2 or greater than 9.
;                  @Error: 1, @Extended: 6 = $iMinTrailingChar not an Integer, less than 2 or greater than 9.
;                  @Error: 1, @Extended: 7 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoHyphen
;                  |                               2 = Error setting $bHyphenNoCaps
;                  |                               4 = Error setting $iMaxHyphens
;                  |                               8 = Error setting $iMinLeadingChar
;                  |                               16 = Error setting $iMinTrailingChar
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 6.4.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  $bAutoHyphen needs to be set to True for the rest of the settings to be activated, but they will be still successfully be set regardless.
; Related .......: _LOWriter_DirFrmtParPageBreak, _LOWriter_DirFrmtParTxtFlowOpt, _LOWriter_ParStyleHyphenation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParHyphenation(ByRef $oSelection, $bAutoHyphen = Null, $bHyphenNoCaps = Null, $iMaxHyphens = Null, $iMinLeadingChar = Null, $iMinTrailingChar = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("ParaIsHyphenation") ; Resetting one resets all.
		If __LO_VarsAreNull($bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParHyphenation($oSelection, $bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParHyphenation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParIndent
; Description ...: Set or Retrieve Indent settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParIndent(ByRef $oSelection[, $iBeforeTxt = Null[, $iAfterTxt = Null[, $iFirstLine = Null[, $bAutoFirstLine = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iBeforeTxt          - [optional] (-9998989-17094) Default is Null. The amount of space that you want to indent the paragraph from the page margin. If you want the paragraph to extend into the page margin, enter a negative number. Set in Hundredths of a Millimeter (HMM).
;                  $iAfterTxt           - [optional] (-9998989-17094) Default is Null. The amount of space that you want to indent the paragraph from the page margin. If you want the paragraph to extend into the page margin, enter a negative number. Set in Hundredths of a Millimeter (HMM).
;                  $iFirstLine          - [optional] (-57785-17904) Default is Null. Indents the first line of a paragraph by the amount that you enter. Set in Hundredths of a Millimeter (HMM).
;                  $bAutoFirstLine      - [optional] Default is Null. If True, the first line will be indented automatically.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Indent related settings.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iBeforeText not an Integer, less than -9998989 or greater than 17094.
;                  @Error: 1, @Extended: 3 = $iAfterText not an Integer, less than -9998989 or greater than 17094.
;                  @Error: 1, @Extended: 4 = $iFirstLine not an Integer, less than -57785 or greater than 17094.
;                  @Error: 1, @Extended: 5 = $bAutoFirstLine not a Boolean.
;                  @Error: 1, @Extended: 6 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBeforeTxt
;                  |                               2 = Error setting $iAfterTxt
;                  |                               4 = Error setting $iFirstLine
;                  |                               8 = Error setting $bAutoFirstLine
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  $iFirstLine Indent cannot be set if $bAutoFirstLine is set to True.
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtParAlignment, _LOWriter_DirFrmtParSpace, _LOWriter_ParStyleIndent
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParIndent(ByRef $oSelection, $iBeforeTxt = Null, $iAfterTxt = Null, $iFirstLine = Null, $bAutoFirstLine = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("ParaLeftMargin") ; Resetting one resets all -- but just in case reset the rest.
		$oSelection.setPropertyToDefault("ParaRightMargin")
		$oSelection.setPropertyToDefault("ParaFirstLineIndent")
		$oSelection.setPropertyToDefault("ParaIsAutoFirstLineIndent")
		If __LO_VarsAreNull($iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParIndent($oSelection, $iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParOutLineAndList
; Description ...: Set and Retrieve the Outline and List settings for a paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParOutLineAndList(ByRef $oDoc, ByRef $oSelection[, $iOutline = Null[, $sNumStyle = Null[, $bParLineCount = Null[, $iLineCountVal = Null]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iOutline            - [optional] (0-10) Default is Null. The Outline Level, see Constants, $LOW_PAR_OUTLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sNumStyle           - [optional] Default is Null. The name of the Numbering Style for the Paragraph numbering. Call with "" for None.
;                  $bParLineCount       - [optional] Default is Null. If True, the paragraph is included in the line numbering.
;                  $iLineCountVal       - [optional] Default is Null. The start value for numbering if a new numbering starts at this paragraph. Call with 0 for no line numbering restart.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not an Object.
;                  @Error: 1, @Extended: 3 = $iOutline not an Integer, less than 0 or greater than 10. See Constants, $LOW_PAR_OUTLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 4 = $sNumStyle not a String.
;                  @Error: 1, @Extended: 5 = Numbering Style called in $sNumStyle not found in document.
;                  @Error: 1, @Extended: 6 = $bParLineCount not a Boolean.
;                  @Error: 1, @Extended: 7 = $iLineCountVal not an Integer, or less than 0.
;                  @Error: 1, @Extended: 8 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iOutline
;                  |                               2 = Error setting $sNumStyle
;                  |                               4 = Error setting $bParLineCount
;                  |                               8 = Error setting $iLineCountVal
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting. $iOutline, $bParLineCount, and $iLineCountVal all are reset together.
;                  In LibreOffice User Interface (UI), there are two options available when applying direct formatting, "Restart numbering at this paragraph", and "start value", these are too glitchy to make available, I am able to set "Restart numbering at this paragraph" to True, but I cannot set it to False, and I am unable to clear either setting once applied, so for those reasons I am not including it in this UDF.
; Related .......: _LOWriter_ParStyleOutLineAndList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParOutLineAndList(ByRef $oDoc, ByRef $oSelection, $iOutline = Null, $sNumStyle = Null, $bParLineCount = Null, $iLineCountVal = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If __LOWriter_AnyAreDefault($iOutline, $sNumStyle, $bParLineCount, $iLineCountVal) Then
		If ($iOutline = Default) Then
			$oSelection.setPropertyToDefault("OutlineLevel")
			$iOutline = Null
		EndIf

		If ($sNumStyle = Default) Then
			$oSelection.NumberingStyleName = "" ; Set to no numbering style first in order to reset successfully.
			$oSelection.setPropertyToDefault("NumberingStyleName")
			$sNumStyle = Null
		EndIf

		If ($bParLineCount = Default) Then
			$oSelection.setPropertyToDefault("ParaLineNumberCount")
			$bParLineCount = Null
		EndIf

		If ($iLineCountVal = Default) Then
			$oSelection.setPropertyToDefault("ParaLineNumberStartValue")
			$iLineCountVal = Null
		EndIf

		If __LO_VarsAreNull($iOutline, $sNumStyle, $bParLineCount, $iLineCountVal) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParOutLineAndList($oDoc, $oSelection, $iOutline, $sNumStyle, $bParLineCount, $iLineCountVal)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParOutLineAndList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParPageBreak
; Description ...: Set or Retrieve Page Break Settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParPageBreak(ByRef $oDoc, ByRef $oSelection[, $iBreakType = Null[, $sPageStyle = Null[, $iPgNumOffSet = Null[, $bClearDirFrmt = False]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iBreakType          - [optional] (0-6) Default is Null. The Page Break Type. See Constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPageStyle          - [optional] Default is Null. Creates a page break before the paragraph it belongs to and assigns the new page style to use. Note: If you set this parameter, to remove the page break setting you must set this to "".
;                  $iPgNumOffSet        - [optional] Default is Null. If a page break property is set at a paragraph, this property contains the new value for the page number.
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Page Break, including Type, Number offset and Page Style. See Remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not an Object.
;                  @Error: 1, @Extended: 3 = $iBreakType not an Integer, less than 0 or greater than 6. See Constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 4 = $sPageStyle not a String.
;                  @Error: 1, @Extended: 5 = Page Style called in $sPageStyle not found in document.
;                  @Error: 1, @Extended: 6 = $iPgNumOffSet not an Integer, or less than 0.
;                  @Error: 1, @Extended: 7 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBreakType
;                  |                               2 = Error setting $sPageStyle
;                  |                               4 = Error setting $iPgNumOffSet
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Clearing directly formatted page breaks may fail, If the cursor selection contains more than one paragraph that has more than one type of page break, it may fail to literally reset it to the paragraph style's original settings even though it returns a success, you will need to reset each paragraph one at a time if this is the case.
;                  Break Type must be set before Page Style will be able to be set, and page style needs set before $iPgNumOffSet can be set.
;                  LibreOffice doesn't directly show in its User interface options for Break type constants #3 and #6 (Column both) and (Page both), but doesn't throw an error when being set to either one, so they are included here, though I'm not sure if they will work correctly.
; Related .......: _LOWriter_DirFrmtParTxtFlowOpt, _LOWriter_DirFrmtParHyphenation, _LOWriter_ParStylePageBreak
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParPageBreak(ByRef $oDoc, ByRef $oSelection, $iBreakType = Null, $sPageStyle = Null, $iPgNumOffSet = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If $bClearDirFrmt Then
		$oSelection.PageDescName = ""
		$oSelection.BreakType = $LOW_BREAK_NONE
		$oSelection.setPropertyToDefault("BreakType")
		$oSelection.setPropertyToDefault("PageDescName")
		$oSelection.setPropertyToDefault("PageNumberOffset")

		If __LO_VarsAreNull($iBreakType, $iPgNumOffSet, $sPageStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParPageBreak($oDoc, $oSelection, $iBreakType, $sPageStyle, $iPgNumOffSet)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParShadow
; Description ...: Set or Retrieve the Shadow settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParShadow(ByRef $oSelection[, $iLocation = Null[, $iColor = Null[, $iWidth = Null[, $bClearDirFrmt = False]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iLocation           - [optional] (0-4) Default is Null. The location of the shadow compared to the paragraph. See Constants, $LOW_SHADOW_LOCATION_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iColor              - [optional] (0-16777215) Default is Null. The color of the shadow, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iWidth              - [optional] Default is Null. The width of the shadow set in Hundredths of a Millimeter (HMM).
;                  $bClearDirFrmt       - [optional] Default is False. If True, clears ALL direct formatting of Shadow Width, Color and Location.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. $bClearDirFrmt was called with True, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOW_SHADOW_LOCATION_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $iWidth not an Integer, or less than 0.
;                  @Error: 1, @Extended: 5 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Shadow Format Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Shadow Format Object for Error checking.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLocation
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $iWidth
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  LibreOffice may change the shadow width +/- a Hundredth of a Millimeter (HMM).
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert, _LOWriter_DirFrmtCharShadow, _LOWriter_ParStyleShadow
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParShadow(ByRef $oSelection, $iLocation = Null, $iColor = Null, $iWidth = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("ParaShadowFormat")
		If __LO_VarsAreNull($iLocation, $iColor, $iWidth) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParShadow($oSelection, $iLocation, $iColor, $iWidth)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParSpace
; Description ...: Set and Retrieve Line Spacing settings for a paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParSpace(ByRef $oSelection[, $iAbovePar = Null[, $iBelowPar = Null[, $bAddSpace = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null[, $bPageLineSpc = Null]]]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iAbovePar           - [optional] (0-10008) Default is Null. The Space above a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iBelowPar           - [optional] (0-10008) Default is Null. The Space Below a paragraph, in Hundredths of a Millimeter (HMM).
;                  $bAddSpace           - [optional] Default is Null. If True, the top and bottom margins of the paragraph should not be applied when the previous and next paragraphs have the same style. LibreOffice 3.6 and Up.
;                  $iLineSpcMode        - [optional] (0-3) Default is Null. The type of the line spacing of a paragraph. See Constants, $LOW_PAR_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3. Also notice min and max values for each.
;                  $iLineSpcHeight      - [optional] Default is Null. This value specifies the spacing of the lines. See Remarks for Minimum and Max values.
;                  $bPageLineSpc        - [optional] Default is Null. If True, register mode is applied to the paragraph. See Remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters. If the LibreOffice version is below 3.6, the $bAddSpace parameter will return a Null value.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iAbovePar not an Integer, less than 0 or greater than 10008.
;                  @Error: 1, @Extended: 3 = $iBelowPar not an Integer, less than 0 or greater than 10008.
;                  @Error: 1, @Extended: 4 = $bAddSpc not a Boolean.
;                  @Error: 1, @Extended: 5 = $iLineSpcMode Not an Integer, less than 0 or greater than 3. See Constants, $LOW_PAR_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iLineSpcHeight not an Integer.
;                  @Error: 1, @Extended: 7 = $iLineSpcMode set to 0 (Proportional) and $iLineSpcHeight less than 6(%) or greater than 65535(%).
;                  @Error: 1, @Extended: 8 = $iLineSpcMode set to 1 or 2 (Minimum, or Leading) and $iLineSpcHeight less than 0 or greater than 10008.
;                  @Error: 1, @Extended: 9 = $iLineSpcMode set to 3(Fixed) and $iLineSpcHeight less than 51 or greater than 10008.
;                  @Error: 1, @Extended: 10 = $bPageLineSpc not a Boolean.
;                  @Error: 1, @Extended: 11 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaLineSpacing Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAbovePar
;                  |                               2 = Error setting $iBelowPar
;                  |                               4 = Error setting $bAddSpace
;                  |                               8 = Error setting $iLineSpcMode
;                  |                               16 = Error setting $iLineSpcHeight
;                  |                               32 = Error setting bPageLineSpc
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting. $iAbovePar, $iBelowPar, and $bAddSpace are all reset together, $iLineSpace Mode / Height also reset together.
;                  $bPageLineSpc(Register mode) is only used if the register mode property of the page style is switched on. $bPageLineSpc(Register Mode) Aligns the baseline of each line of text to a vertical document grid, so that each line is the same height.
;                  The settings in LibreOffice, (Single,1.15, 1.5, Double,) Use the Proportional mode, and are just varying percentages. e.g Single = 100, 1.15 = 115%, 1.5 = 150%, Double = 200%.
;                  $iLineSpcHeight depends on the $iLineSpcMode used, see constants for accepted Input values.
;                  $iAbovePar, $iBelowPar, $iLineSpcHeight may change +/- a Hundredth of a Millimeter (HMM) once set.
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtParIndent, _LOWriter_DirFrmtParAlignment, _LOWriter_ParStyleSpace
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParSpace(ByRef $oSelection, $iAbovePar = Null, $iBelowPar = Null, $bAddSpace = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null, $bPageLineSpc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

	If __LOWriter_AnyAreDefault($iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc) Then
		If ($iAbovePar = Default) Then
			$oSelection.setPropertyToDefault("ParaTopMargin")
			$iAbovePar = Null
		EndIf

		If ($iBelowPar = Default) Then
			$oSelection.setPropertyToDefault("ParaBottomMargin")
			$iBelowPar = Null
		EndIf

		If ($bAddSpace = Default) Then
			If Not __LO_VersionCheck(3.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

			$oSelection.setPropertyToDefault("ParaContextMargin")
			$bAddSpace = Null
		EndIf

		If ($iLineSpcMode = Default) Or ($iLineSpcHeight = Default) Then
			$oSelection.setPropertyToDefault("ParaLineSpacing")
			$iLineSpcMode = Null
		EndIf

		If ($iLineSpcHeight = Default) Or ($iLineSpcHeight = Default) Then
			$oSelection.setPropertyToDefault("ParaLineSpacing")
			$iLineSpcHeight = Null
		EndIf

		If ($bPageLineSpc = Default) Then
			$oSelection.setPropertyToDefault("ParaRegisterModeActive")
			$bPageLineSpc = Null
		EndIf

		If __LO_VarsAreNull($iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParSpace($oSelection, $iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParSpace

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTabStopCreate
; Description ...: Create a new TabStop for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTabStopCreate(ByRef $oSelection, $iPosition[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iPosition           - The TabStop position to set the new TabStop to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOW_PAR_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value (See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOW_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc value (see AutoIt function) of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Settings were successfully set. New TabStop position is returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iPosition not an Integer.
;                  @Error: 1, @Extended: 3 = Tab Stop position called in $iPosition already exists in this Paragraph.
;                  @Error: 1, @Extended: 4 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOW_PAR_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iDecChar not an Integer.
;                  @Error: 1, @Extended: 6 = $iFillChar not an Integer.
;                  @Error: 1, @Extended: 7 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.style.TabStop" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Array Object.
;                  @Error: 3, @Extended: 2 = Error retrieving list of TabStop Positions.
;                  @Error: 3, @Extended: 3 = Failed to identify the new Tabstop position once inserted.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iAlignment
;                  |                               4 = Error setting $iDecChar
;                  |                               8 = Error setting $iFillChar
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  $iPosition once set can vary +/- a Hundredth of a Millimeter (HMM). To ensure you can identify the tabstop to modify it again, this function returns the new TabStop position in @Extended when $iPosition is set, return value will be set to 2. See Return Values.
;                  Since $iPosition can fluctuate +/- a Hundredth of a Millimeter (HMM) when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;                  $iFillChar, LibreOffice's Default value, "None" is in reality a space character which is Asc value 32. The other values offered by LibreOffice are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can also enter a custom ASC value. See ASC AutoIt Function and "ASCII Character Codes" in the AutoIt help file.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  $iNewTabStop position is still returned even though some settings weren't successfully set, the new TabStop was still created.
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtParTabStopDelete, _LOWriter_DirFrmtParTabStopMod, _LOWriter_DirFrmtParTabStopsGetList, _LOWriter_ParStyleTabStopCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTabStopCreate(ByRef $oSelection, $iPosition, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iPosition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$iPosition = __LOWriter_ParTabStopCreate($oSelection, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $iPosition)
EndFunc   ;==>_LOWriter_DirFrmtParTabStopCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTabStopDelete
; Description ...: Delete a TabStop from a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTabStopDelete(ByRef $oDoc, ByRef $oSelection, $iTabStop)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iTabStop            - The TabStop position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Returning True if the TabStop was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not an Object.
;                  @Error: 1, @Extended: 3 = $iTabStop not an Integer.
;                  @Error: 1, @Extended: 4 = $iTabStop not found in this Paragraph.
;                  @Error: 1, @Extended: 5 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
;                  @Error: 3, @Extended: 2 = Failed to identify and delete TabStop in Paragraph.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essentially the "length" of a TabStop from the edge of a page margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per paragraph.
; Related .......: _LOWriter_DirFrmtParTabStopCreate, _LOWriter_DirFrmtParTabStopsGetList, _LOWriter_ParStyleTabStopDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTabStopDelete(ByRef $oDoc, ByRef $oSelection, $iTabStop)
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_ParTabStopDelete($oSelection, $oDoc, $iTabStop)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParTabStopDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop in a Paragraph from Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTabStopMod(ByRef $oSelection, $iTabStop[, $iPosition = Null[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $iTabStop            - The TabStop position of the TabStop to modify. See Remarks.
;                  $iPosition           - [optional] Default is Null. The New position to set the TabStop called in $iTabStop to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOW_PAR_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOW_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: ?, Return: 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
;                  @Error: 0, @Extended: 0, Return: 3 = Success. $iTabStop parameter was called with Default, and rest of parameters were called with Null. Direct formatting inserted TabStops have been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $iTabStop not an Integer.
;                  @Error: 1, @Extended: 3 = TabStop called in $iTabStop not found in this Paragraph or selection.
;                  @Error: 1, @Extended: 4 = $iPosition not an Integer.
;                  @Error: 1, @Extended: 5 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOW_PAR_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iDecChar not an Integer.
;                  @Error: 1, @Extended: 7 = $iFillChar not an Integer.
;                  @Error: 1, @Extended: 8 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Requested TabStop Object.
;                  @Error: 3, @Extended: 3 = Paragraph style already contains a TabStop at the length/Position specified in $iPosition.
;                  @Error: 3, @Extended: 4 = Error retrieving list of TabStop Positions.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iAlignment
;                  |                               4 = Error setting $iDecChar
;                  |                               8 = Error setting $iFillChar
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a $iTabStop with Default keyword to clear all direct formatting created TabStops.
;                  $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per paragraph.
;                  $iPosition once set can vary +/- a Hundredth of a Millimeter (HMM). To ensure you can identify the tabstop to modify it again, This function returns the new TabStop position in @Extended when $iPosition is set, return value will be set to 2. See Return Values.
;                  Since $iPosition can fluctuate +/- a Hundredth of a Millimeter (HMM) when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;                  $iFillChar, LibreOffice's Default value, "None" is in reality a space character which is Asc value 32. The other values offered by LibreOffice are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
; Related .......: _LO_UnitConvert, _LOWriter_DirFrmtParTabStopCreate, _LOWriter_DirFrmtParTabStopDelete, _LOWriter_DirFrmtParTabStopsGetList, _LOWriter_ParStyleTabStopMod
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTabStopMod(ByRef $oSelection, $iTabStop, $iPosition = Null, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If ($iTabStop = Default) Then
		$oSelection.setPropertyToDefault("ParaTabStops")

		Return SetError($__LO_STATUS_SUCCESS, 0, 3)
	EndIf

	$vReturn = __LOWriter_ParTabStopMod($oSelection, $iTabStop, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParTabStopMod

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTabStopsGetList
; Description ...: Retrieve an array of TabStops available in a Paragraph from Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTabStopsGetList(ByRef $oSelection)
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array of TabStops. @Extended set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
; Related .......: _LOWriter_DirFrmtParTabStopCreate, _LOWriter_DirFrmtParTabStopDelete, _LOWriter_DirFrmtParTabStopMod, _LOWriter_ParStyleTabStopsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTabStopsGetList(ByRef $oSelection)
	Local $aiTabList

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$aiTabList = __LOWriter_ParTabStopsGetList($oSelection)

	Return SetError(@error, @extended, $aiTabList)
EndFunc   ;==>_LOWriter_DirFrmtParTabStopsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTxtFlowOpt
; Description ...: Set and Retrieve Text Flow settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTxtFlowOpt(ByRef $oSelection[, $bParSplit = Null[, $bKeepTogether = Null[, $iParOrphans = Null[, $iParWidows = Null]]]])
; Parameters ....: $oSelection          - A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object/Object Section returned from _LOWriter_CursorParObjCreateList or _LOWriter_CursorParObjSectionsGet function.
;                  $bParSplit           - [optional] Default is Null. If False, prevents the paragraph from getting split between two pages or columns
;                  $bKeepTogether       - [optional] Default is Null. If True, prevents page or column breaks between this and the following paragraph
;                  $iParOrphans         - [optional] (0, 2-9) Default is Null. Specifies the minimum number of lines of the paragraph that have to be at bottom of a page if the paragraph is spread over more than one page. 0 = disabled.
;                  $iParWidows          - [optional] (0, 2-9) Default is Null. Specifies the minimum number of lines of the paragraph that have to be at top of a page if the paragraph is spread over more than one page. 0 = disabled.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. One or more parameter(s) were called with Default, and rest of parameters were called with Null. Direct formatting has been successfully cleared.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSelection not an Object.
;                  @Error: 1, @Extended: 2 = $bParSplit not a Boolean.
;                  @Error: 1, @Extended: 3 = $bKeepTogether not a Boolean.
;                  @Error: 1, @Extended: 4 = $iParOrphans not an Integer, less than 0, equal to 1, or greater than 9.
;                  @Error: 1, @Extended: 5 = $iParWidows not an Integer, less than 0, equal to 1, or greater than 9.
;                  @Error: 1, @Extended: 6 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bParSplit
;                  |                               2 = Error setting $bKeepTogether
;                  |                               4 = Error setting $iParOrphans
;                  |                               8 = Error setting $iParWidows
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is messy to deal with both by proxy (such as by AutoIt automation) and directly in the document, and is generally not recommended to use. Character and Paragraph styles are generally recommended instead.
;                  Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of 0, False, Null, etc.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Call a Parameter with Default keyword to clear direct formatting for that setting. Resetting Orphan or Widow will reset $bParSplit to False if it was set to True.
;                  If you do not set ParSplit to True, the rest of the settings will still show to have been set, but will not become active until $bParSplit is set to True.
; Related .......: _LOWriter_ParStyleHyphenation, _LOWriter_ParStylePageBreak, _LOWriter_ParStyleTxtFlowOpt
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTxtFlowOpt(ByRef $oSelection, $bParSplit = Null, $bKeepTogether = Null, $iParOrphans = Null, $iParWidows = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If __LOWriter_AnyAreDefault($bParSplit, $bKeepTogether, $iParOrphans, $iParWidows) Then
		If ($bParSplit = Default) Then
			$oSelection.setPropertyToDefault("ParaSplit")
			$bParSplit = Null
		EndIf

		If ($bKeepTogether = Default) Then
			$oSelection.setPropertyToDefault("ParaKeepTogether")
			$bKeepTogether = Null
		EndIf

		If ($iParOrphans = Default) Then
			$oSelection.setPropertyToDefault("ParaOrphans")
			$iParOrphans = Null
		EndIf

		If ($iParWidows = Default) Then
			$oSelection.setPropertyToDefault("ParaWidows")
			$iParWidows = Null
		EndIf

		If __LO_VarsAreNull($bParSplit, $bKeepTogether, $iParOrphans, $iParWidows) Then Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParTxtFlowOpt($oSelection, $bParSplit, $bKeepTogether, $iParOrphans, $iParWidows)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParTxtFlowOpt
