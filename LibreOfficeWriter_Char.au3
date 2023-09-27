;~ #AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

#include "LibreOfficeWriter_Font.au3"

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter)
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.0.3
; Description ...: Provides basic functionality through Autoit for interacting with Libre Office Writer.
; Author(s) .....: donnyh13
; Sources . . . .:  jguinch -- Printmgr.au3, used (_PrintMgr_EnumPrinter);
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
;_LOWriter_CharStyleBorderColor
;_LOWriter_CharStyleBorderPadding
;_LOWriter_CharStyleBorderStyle
;_LOWriter_CharStyleBorderWidth
;_LOWriter_CharStyleCreate
;_LOWriter_CharStyleDelete
;_LOWriter_CharStyleEffect
;_LOWriter_CharStyleExists
;_LOWriter_CharStyleFont
;_LOWriter_CharStyleFontColor
;_LOWriter_CharStyleGetObj
;_LOWriter_CharStyleOrganizer
;_LOWriter_CharStyleOverLine
;_LOWriter_CharStylePosition
;_LOWriter_CharStyleRotateScale
;_LOWriter_CharStyleSet
;_LOWriter_CharStylesGetNames
;_LOWriter_CharStyleShadow
;_LOWriter_CharStyleSpacing
;_LOWriter_CharStyleStrikeOut
;_LOWriter_CharStyleUnderLine
;_LOWriter_FontExists
;_LOWriter_FontsList
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleBorderColor
; Description ...: Set and Retrieve the Character Style Border Line Color. Libre Office 4.2 and Up.
; Syntax ........: _LOWriter_CharStyleBorderColor(Byref $oCharStyle[, $iTop = Null[, $iBottom = Null[,$iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Color of the
;				   +						Character Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Color of the
;				   +						Character Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Color of the
;				   +						Character Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Color of the
;				   +						Character Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value.
; Internal Remark: Certain Error values are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0 or higher than 16,777,215 or not
;				   +								set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0 or higher than 16,777,215 or
;				   +								not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0 or higher than 16,777,215 or not
;				   +								set to Null.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0 or higher than 16,777,215 or
;				   +								not set to Null.
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
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Color Constants: $LOW_COLOR_BLACK(0),
;					$LOW_COLOR_WHITE(16777215),
;					$LOW_COLOR_LGRAY(11711154),
;					$LOW_COLOR_GRAY(8421504),
;					$LOW_COLOR_DKGRAY(3355443),
;					$LOW_COLOR_YELLOW(16776960),
;					$LOW_COLOR_GOLD(16760576),
;					$LOW_COLOR_ORANGE(16744448),
;					$LOW_COLOR_BRICK(16728064),
;					$LOW_COLOR_RED(16711680),
;					$LOW_COLOR_MAGENTA(12517441),
;					$LOW_COLOR_PURPLE(8388736),
;					$LOW_COLOR_INDIGO(5582989),
;					$LOW_COLOR_BLUE(2777241),
;					$LOW_COLOR_TEAL(1410150),
;					$LOW_COLOR_GREEN(43315),
;					$LOW_COLOR_LIME(8508442),
;					$LOW_COLOR_BROWN(9127187).
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_ConvertColorFromLong,
;					_LOWriter_ConvertColorToLong, _LOWriter_CharStyleBorderWidth, _LOWriter_CharStyleBorderStyle,
;					_LOWriter_CharStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleBorderColor(ByRef $oCharStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_CharBorder($oCharStyle, False, False, True, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleBorderPadding
; Description ...: Set and retrieve the distance between the border and the characters. Libre Office 4.2 and Up.
; Syntax ........: _LOWriter_CharStyleBorderPadding(Byref $oCharStyle[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four values to the same value.
;				   +											When used, all other parameters are ignored.  In MicroMeters.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top border distance in
;				   +							MicroMeters.
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom border distance in
;				   +							MicroMeters.
;                  $iLeft               - [optional] an integer value. Default is Null. Set the left border distance in
;				   +							MicroMeters.
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right border distance in
;				   +							MicroMeters.
; Return values .:  Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iAll not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $Left not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iRight not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16
;				   |								1 = Error setting $iAll border distance
;				   |								2 = Error setting $iTop border distance
;				   |								4 = Error setting $iBottom border distance
;				   |								8 = Error setting $iLeft border distance
;				   |								16 = Error setting $iRight border distance
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					All distance values are set in MicroMeters. Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer, _LOWriter_CharStyleBorderWidth, _LOWriter_CharStyleBorderStyle,
;					_LOWriter_CharStyleBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleBorderPadding(ByRef $oCharStyle, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharBorderPadding($oCharStyle, $iAll, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleBorderStyle
; Description ...: Set or Retrieve the Character Style Border Line style. Libre Office 4.2 and Up.
; Syntax ........: _LOWriter_CharStyleBorderStyle(Byref $oCharStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Style of the
;				   +							Character Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Style of the
;				   +							Character Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Style of the
;				   +							Character Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Style of the
;				   +							Character Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
; Internal Remark: Error values for Initialization and Processing, except for Processing @Extended 1, are passed from the
;					internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to higher than 17 and not equal to 0x7FFF,
;				   +									Or $iTop is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to higher than 17 and not equal to
;				   +								0x7FFF, Or $iBottom is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to higher than 17 and not equal to 0x7FFF,
;				   +									Or $iLeft is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to higher than 17 and not equal to
;				   +									0x7FFF, Or $iRight is set to less than 0 or not set to Null.
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
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					 Call any optional parameter with Null keyword to skip it.
; Style Constants: $LOW_BORDERSTYLE_NONE(0x7FFF) No border line,
;					$LOW_BORDERSTYLE_SOLID(0) Solid border line,
;					$LOW_BORDERSTYLE_DOTTED(1) Dotted border line,
;					$LOW_BORDERSTYLE_DASHED(2) Dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE(3) Double border line,
;					$LOW_BORDERSTYLE_THINTHICK_SMALLGAP(4) Double border line with a thin line outside and a thick line inside
;						separated by a small gap,
;					$LOW_BORDERSTYLE_THINTHICK_MEDIUMGAP(5) Double border line with a thin line outside and a thick line inside
;						separated by a medium gap,
;						$LOW_BORDERSTYLE_THINTHICK_LARGEGAP(6) Double border line with a thin line outside and a thick line
;						inside separated by a large gap,
;					$LOW_BORDERSTYLE_THICKTHIN_SMALLGAP(7) Double border line with a thick line outside and a thin line inside
;						separated by a small gap,
;					$LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP(8) Double border line with a thick line outside and a thin line inside
;						separated by a medium gap,
;					$LOW_BORDERSTYLE_THICKTHIN_LARGEGAP(9) Double border line with a thick line outside and a thin line inside
;						separated by a large gap,
;					$LOW_BORDERSTYLE_EMBOSSED(10) 3D embossed border line,
;					$LOW_BORDERSTYLE_ENGRAVED(11) 3D engraved border line,
;					$LOW_BORDERSTYLE_OUTSET(12) Outset border line,
;					$LOW_BORDERSTYLE_INSET(13) Inset border line,
;					$LOW_BORDERSTYLE_FINE_DASHED(14) Finely dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE_THIN(15) Double border line consisting of two fixed thin lines separated by a
;						variable gap,
;					$LOW_BORDERSTYLE_DASH_DOT(16) Line consisting of a repetition of one dash and one dot,
;					$LOW_BORDERSTYLE_DASH_DOT_DOT(17) Line consisting of a repetition of one dash and 2 dots.
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_CharStyleBorderWidth,
;					_LOWriter_CharStyleBorderColor, _LOWriter_CharStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleBorderStyle(ByRef $oCharStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_CharBorder($oCharStyle, False, True, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleBorderWidth
; Description ...: Set and Retrieve the Character Style Border Line Width. Libre Office 4.2 and Up.
; Syntax ........: _LOWriter_CharStyleBorderWidth(Byref $oCharStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line width of the
;				   +							Character Style in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Width of the
;				   +							Character Style in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line width of the
;				   +							Character Style in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Width of the
;				   +							Character Style in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null.
; Internal Remark: Certain Error values are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0 or not set to Null.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;													settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set them to 0
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Width Constants:  $LOW_BORDERWIDTH_HAIRLINE(2),
;					$LOW_BORDERWIDTH_VERY_THIN(18),
;					$LOW_BORDERWIDTH_THIN(26),
;					$LOW_BORDERWIDTH_MEDIUM(53),
;					$LOW_BORDERWIDTH_THICK(79),
;					$LOW_BORDERWIDTH_EXTRA_THICK(159)
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer, _LOWriter_CharStyleBorderColor, _LOWriter_CharStyleBorderStyle,
;					_LOWriter_CharStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleBorderWidth(ByRef $oCharStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_CharBorder($oCharStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleCreate
; Description ...: Create a new Character Style in a Document.
; Syntax ........: _LOWriter_CharStyleCreate(Byref $oDoc, $sCharStyle)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sCharStyle          - a string value. The Name of the New Character Style to Create.
; Return values .:  Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCharStyle not a String.
;				   @Error 1 @Extended 3 Return 0 = $sCharStyle name already exists in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Retrieving "CharacterStyle" Object.
;				   @Error 2 @Extended 2 Return 0 = Error Creating "com.sun.star.style.CharacterStyle" Object.
;				   @Error 2 @Extended 3 Return 0 = Error Retrieving Created Character Style Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error creating new Character Style by Name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. New Character Style successfully created. Returning Character
;				   +												Style Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_CharStyleDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleCreate(ByRef $oDoc, $sCharStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCharStyles, $oStyle, $oCharStyle

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oCharStyles = $oDoc.StyleFamilies().getByName("CharacterStyles")
	If Not IsObj($oCharStyles) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	If _LOWriter_CharStyleExists($oDoc, $sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oStyle = $oDoc.createInstance("com.sun.star.style.CharacterStyle")
	If Not IsObj($oStyle) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oCharStyles.insertByName($sCharStyle, $oStyle)

	If Not $oCharStyles.hasByName($sCharStyle) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	$oCharStyle = $oCharStyles.getByName($sCharStyle)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCharStyle)
EndFunc   ;==>_LOWriter_CharStyleCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleDelete
; Description ...: Delete a User-Created Character Style from a Document.
; Syntax ........: _LOWriter_CharStyleDelete(Byref $oDoc, $oCharStyle[, $bForceDelete = False[, $sReplacementStyle = ""]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function. Must be a User-Created Style, not a built-in Style native
;				   +						to Libre-Office.
;                  $bForceDelete        - [optional] a boolean value. Default is False. If True Character style will be deleted
;				   +						regardless of whether it is in use or not.
;                  $sReplacementStyle   - [optional] a string value. Default is "". The Character style to use instead of the
;				   +						one being deleted if the Character style being deleted was already applied to text
;				   +						in the document. "" = No Character Style
; Return values .:  Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 4 Return 0 = $bForceDelete not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sReplacementStyle not a String.
;				   @Error 1 @Extended 6 Return 0 = $sReplacementStyle doesn't exist in Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving "CharacterStyles" Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Character Style Name.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = $sCharStyle is not a User-Created Character Style and cannot be deleted.
;				   @Error 3 @Extended 2 Return 0 = $sCharStyle is in use and $bForceDelete is false.
;				   @Error 3 @Extended 3 Return 0 = $sCharStyle still exists after deletion attempt.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Character Style called in $oCharStyle was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleDelete(ByRef $oDoc, ByRef $oCharStyle, $bForceDelete = False, $sReplacementStyle = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCharStyles
	Local $sCharStyle

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bForceDelete) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplacementStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($sReplacementStyle <> "") And Not _LOWriter_CharStyleExists($oDoc, $sReplacementStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$oCharStyles = $oDoc.StyleFamilies().getByName("CharacterStyles")
	If Not IsObj($oCharStyles) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$sCharStyle = $oCharStyle.Name()
	If Not IsString($sCharStyle) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	If Not $oCharStyle.isUserDefined() Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	If $oCharStyle.isInUse() And Not ($bForceDelete) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0) ; If Style is in use return an error unless force delete is true.

	If ($sReplacementStyle <> "") Then $oCharStyle.setParentStyle($sReplacementStyle)
	;If User has called a specific style to replace this style, set it to that.

	$oCharStyles.removeByName($sCharStyle)
	Return ($oCharStyles.hasByName($sCharStyle)) ? SetError($__LOW_STATUS_PROCESSING_ERROR, 3, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CharStyleDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleEffect
; Description ...: Set or Retrieve the Font Effect settings for a Character Style.
; Syntax ........: _LOWriter_CharStyleEffect(Byref $oCharStyle[, $iRelief = Null[, $iCase = Null[, $bHidden = Null[, $bOutline = Null[, $bShadow = Null]]]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $iRelief             - [optional] an integer value. Default is Null. The Character Relief style. See Constants
;				   +									below.
;                  $iCase               - [optional] an integer value. Default is Null. The Character Case Style. See Constants
;				   +									below.
;                  $bHidden             - [optional] a boolean value. Default is Null. Whether the Characters are hidden or not.
;                  $bOutline            - [optional] a boolean value. Default is Null. Whether the characters have an outline
;				   +									around the outside.
;                  $bShadow             - [optional] a boolean value. Default is Null. Whether the characters have a shadow.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iRelief not an integer or less than 0 or greater than 2. See Constants.
;				   @Error 1 @Extended 5 Return 0 = $iCase not an integer or less than 0 or greater than 4. See Constants.
;				   @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bOutline not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $bShadow not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4,8, 16
;				   |								1 = Error setting $iRelief
;				   |								2 = Error setting $iCase
;				   |								4 = Error setting $bHidden
;				   |								8 = Error setting $bOutline
;				   |								16 = Error setting $bShadow
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Relief Constants: $LOW_RELIEF_NONE(0); no relief is used.
;						$LOW_RELIEF_EMBOSSED(1); the font relief is embossed.
;						$LOW_RELIEF_ENGRAVED(2); the font relief is engraved.
; Case Constants : 	$LOW_CASEMAP_NONE(0); The case of the characters is unchanged.
;						$LOW_CASEMAP_UPPER(1); All characters are put in upper case.
;						$LOW_CASEMAP_LOWER(2); All characters are put in lower case.
;						$LOW_CASEMAP_TITLE(3); The first character of each word is put in upper case.
;						$LOW_CASEMAP_SM_CAPS(4); All characters are put in upper case, but with a smaller font height.
; Related .......: _LOWriter_CharStyleGetObj,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleEffect(ByRef $oCharStyle, $iRelief = Null, $iCase = Null, $bHidden = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharEffect($oCharStyle, $iRelief, $iCase, $bHidden, $bOutline, $bShadow)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleExists
; Description ...: Check whether a document contains a Character Style by Name.
; Syntax ........: _LOWriter_CharStyleExists(Byref $oDoc, $sCharStyle)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sCharStyle          - a string value. The Character Style to check.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object,
;				   @Error 1 @Extended 2 Return 0 = $sPageStyle not a String
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean  = Success. If Character Style exists then True is returned, if not,
;				   +									False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleExists(ByRef $oDoc, $sCharStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $oDoc.StyleFamilies.getByName("CharacterStyles").hasByName($sCharStyle) Then Return SetError($__LOW_STATUS_SUCCESS, 0, True)

	Return SetError($__LOW_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOWriter_CharStyleExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleFont
; Description ...: Set and Retrieve the Font Settings for a Character Style.
; Syntax ........: _LOWriter_CharStyleFont(Byref $oDoc, $oCharStyle[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $sFontName           - [optional] a string value. Default is Null. The Font Name to change to.
;                  $nFontSize           - [optional] a general number value. Default is Null. The new Font size.
;                  $iPosture            - [optional] an integer value. Default is Null. Italic setting. See Constants below. Also
;				   +								see remarks.
;                  $iWeight             - [optional] an integer value. Default is Null. Bold settings see Constants below.
;				   +								Also see remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc parameter not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 4 Return 0 = $sFontName not available in current document.
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $sFontName not a String.
;				   @Error 1 @Extended 7 Return 0 = $nFontSize not a Number.
;				   @Error 1 @Extended 8 Return 0 = $iPosture not an Integer, less than 0 or greater than 5. See Constants.
;				   @Error 1 @Extended 9 Return 0 = $iWeight less than 50 and not 0, or more than 200. See Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4,8
;				   |								1 = Error setting $sFontName
;				   |								2 = Error setting $nFontSize
;				   |								4 = Error setting $iPosture
;				   |								8 = Error setting $iWeight
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted,
;					such as oblique, ultra Bold etc. Libre Writer accepts only the predefined weight values, any other values
;					are changed automatically to an acceptable value, which could trigger a settings error.
; Weight Constants : $LOW_WEIGHT_DONT_KNOW(0); The font weight is not specified/known.
;						$LOW_WEIGHT_THIN(50); specifies a 50% font weight.
;						$LOW_WEIGHT_ULTRA_LIGHT(60); specifies a 60% font weight.
;						$LOW_WEIGHT_LIGHT(75); specifies a 75% font weight.
;						$LOW_WEIGHT_SEMI_LIGHT(90); specifies a 90% font weight.
;						$LOW_WEIGHT_NORMAL(100); specifies a normal font weight.
;						$LOW_WEIGHT_SEMI_BOLD(110); specifies a 110% font weight.
;						$LOW_WEIGHT_BOLD(150); specifies a 150% font weight.
;						$LOW_WEIGHT_ULTRA_BOLD(175); specifies a 175% font weight.
;						$LOW_WEIGHT_BLACK(200); specifies a 200% font weight.
; Slant/Posture Constants : $LOW_POSTURE_NONE(0); specifies a font without slant.
;							$LOW_POSTURE_OBLIQUE(1); specifies an oblique font (slant not designed into the font).
;							$LOW_POSTURE_ITALIC(2); specifies an italic font (slant designed into the font).
;							$LOW_POSTURE_DontKnow(3); specifies a font with an unknown slant.
;							$LOW_POSTURE_REV_OBLIQUE(4); specifies a reverse oblique font (slant not designed into the font).
;							$LOW_POSTURE_REV_ITALIC(5); specifies a reverse italic font (slant designed into the font).
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_FontsList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleFont(ByRef $oDoc, ByRef $oCharStyle, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($sFontName <> Null) And Not _LOWriter_FontExists($oDoc, $sFontName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_CharFont($oCharStyle, $sFontName, $nFontSize, $iPosture, $iWeight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting of a Character style.
; Syntax ........: _LOWriter_CharStyleFontColor(Byref $oCharStyle[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $iFontColor          - [optional] an integer value. Default is Null. the desired Color value in Long Integer
;				   +								format, to make the font, can be one of the constants listed below or a
;				   +								custom value. Set to $LOW_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - [optional] an integer value. Default is Null. Transparency percentage. 0 is not
;				   +								visible, 100 is fully visible. Available for Libre Office 7.0 and up.
;                  $iHighlight          - [optional] an integer value. Default is Null. A Color value in Long Integer format,
;				   +								to highlight the text in, can be one of the constants listed below or a
;				   +								custom value. Set to $LOW_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iFontColor not an integer, less than -1 or greater than 16777215.
;				   @Error 1 @Extended 5 Return 0 = $iTransparency not an Integer, or less than 0 or greater than 100%.
;				   @Error 1 @Extended 6 Return 0 = $iHighlight not an integer, less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $FontColor
;				   |								2 = Error setting $iTransparency.
;				   |								4 = Error setting $iHighlight
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 7.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 or 3 Element Array with values in order of function
;				   +								parameters. If The current Libre Office version is below 7.0 the returned
;				   +								array will contain 2 elements, because $iTransparency is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Note: When setting transparency, the value of font color value is changed.
; Color Constants: $LOW_COLOR_OFF(-1)
;					$LOW_COLOR_BLACK(0),
;					$LOW_COLOR_WHITE(16777215),
;					$LOW_COLOR_LGRAY(11711154),
;					$LOW_COLOR_GRAY(8421504),
;					$LOW_COLOR_DKGRAY(3355443),
;					$LOW_COLOR_YELLOW(16776960),
;					$LOW_COLOR_GOLD(16760576),
;					$LOW_COLOR_ORANGE(16744448),
;					$LOW_COLOR_BRICK(16728064),
;					$LOW_COLOR_RED(16711680),
;					$LOW_COLOR_MAGENTA(12517441),
;					$LOW_COLOR_PURPLE(8388736),
;					$LOW_COLOR_INDIGO(5582989),
;					$LOW_COLOR_BLUE(2777241),
;					$LOW_COLOR_TEAL(1410150),
;					$LOW_COLOR_GREEN(43315),
;					$LOW_COLOR_LIME(8508442),
;					$LOW_COLOR_BROWN(9127187).
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_ConvertColorFromLong,
;					_LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleFontColor(ByRef $oCharStyle, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharFontColor($oCharStyle, $iFontColor, $iTransparency, $iHighlight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleGetObj
; Description ...: Retrieve a Character Style Object for use with other CharStyle functions.
; Syntax ........: _LOWriter_CharStyleGetObj(Byref $oDoc, $sCharStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sCharStyle           - a string value. The Character Style name to retrieve the Object for.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCharStyle not a String.
;				   @Error 1 @Extended 3 Return 0 = Character Style defined in $sCharStyle not found in Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Character Style Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved and returned requested Character Style
;				   +										Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_CharStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleGetObj(ByRef $oDoc, $sCharStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCharStyle

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_CharStyleExists($oDoc, $sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oCharStyle = $oDoc.StyleFamilies().getByName("CharacterStyles").getByName($sCharStyle)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCharStyle)
EndFunc   ;==>_LOWriter_CharStyleGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleOrganizer
; Description ...: Set or retrieve the Organizer settings of a Character Style.
; Syntax ........: _LOWriter_CharStyleOrganizer(Byref $oDoc, $oCharStyle[, $sNewCharStyleName = Null[, $sParentStyle = Null[, $bHidden = Null]]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $sNewCharStyleName   - [optional] a string value. Default is Null. The new name to set $sChrStyle Character
;				   +						style to.
;                  $sParentStyle        - [optional] a string value. Default is Null. Set an existing  Character style
;				   +						(or an Empty String ("") = - None -) to apply its settings to the current style.
;				   +						Use the other settings to modify the inherited style settings.
;                  $bHidden             - [optional] a boolean value. Default is Null. Whether to hide the style in the UI.
;				   +						Libre Office version 4.0 and up only.
; Return values .:   Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc parameter not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 4 Return 0 = $sNewCharStyleName not a String.
;				   @Error 1 @Extended 5 Return 0 = $sNewCharStyleName already exists in document.
;				   @Error 1 @Extended 6 Return 0 = $sParentStyle not a String.
;				   @Error 1 @Extended 7 Return 0 = $sParentStyle Doesn't exist in this Document.
;				   @Error 1 @Extended 8 Return 0 = $bHidden not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $sNewParStyleName
;				   |								2 = Error setting $sParentStyle
;				   |								4 = Error setting $bHidden
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 or 3 Element Array with values in order of function
;				   +								parameters. If the Libre Office version is below 4.0, the Array will contain
;				   +								2 elements because $bHidden is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Call this function with only the required parameters (or with all other parameters set
;					to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleOrganizer(ByRef $oDoc, ByRef $oCharStyle, $sNewCharStyleName = Null, $sParentStyle = Null, $bHidden = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOrganizer[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($sNewCharStyleName, $sParentStyle, $bHidden) Then
		If __LOWriter_VersionCheck(4.0) Then
			__LOWriter_ArrayFill($avOrganizer, $oCharStyle.Name(), __LOWriter_CharStyleNameToggle($oCharStyle.ParentStyle(), True), _
					$oCharStyle.Hidden())
		Else
			__LOWriter_ArrayFill($avOrganizer, $oCharStyle.Name(), __LOWriter_CharStyleNameToggle($oCharStyle.ParentStyle(), True))
		EndIf

		Return SetError($__LOW_STATUS_SUCCESS, 1, $avOrganizer)
	EndIf

	If ($sNewCharStyleName <> Null) Then
		If Not IsString($sNewCharStyleName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If _LOWriter_CharStyleExists($oDoc, $sNewCharStyleName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oCharStyle.Name = $sNewCharStyleName
		$iError = ($oCharStyle.Name() = $sNewCharStyleName) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sParentStyle <> Null) Then
		If Not IsString($sParentStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If ($sParentStyle <> "") Then
			If Not _LOWriter_CharStyleExists($oDoc, $sParentStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
			$sParentStyle = __LOWriter_CharStyleNameToggle($sParentStyle)
		EndIf
		$oCharStyle.ParentStyle = $sParentStyle
		$iError = ($oCharStyle.ParentStyle() = $sParentStyle) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$oCharStyle.Hidden = $bHidden
		$iError = ($oCharStyle.Hidden() = $bHidden) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CharStyleOrganizer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleOverLine
; Description ...: Set and retrieve the OverLine settings for a Character style.
; Syntax ........: _LOWriter_CharStyleOverLine(Byref $oCharStyle[, $bWordOnly = Null[, $iOverLineStyle = Null[, $bOLHasColor = Null[, $iOLColor = Null]]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not Overlined.
;                  $iOverLineStyle      - [optional] an integer value. Default is Null. The style of the Overline line, see
;				   +									constants listed below. See Remarks.
;                  $bOLHasColor         - [optional] a boolean value. Default is Null. Whether the Overline is colored, must
;				   +						be set to true in order to set the Overline color.
;                  $iOLColor            - [optional] an integer value. Default is Null. The color of the Overline, set in Long
;				   +						integer format. Can be one of the constants below or a custom value. Set to
;				   +						$LOW_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iOverLineStyle not an Integer, or less than 0 or greater than 18. Check
;				   +									the Constants list.
;				   @Error 1 @Extended 6 Return 0 = $bOLHasColor not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iOLColor not an Integer, or less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $iOverLineStyle
;				   |								4 = Error setting $OLHasColor
;				   |								8 = Error setting $iOLColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Note: OverLine line style uses the same constants as underline style.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Note: $bOLHasColor must be set to true in order to set the Overline color.
; OverLine line style Constants: $LOW_UNDERLINE_NONE(0),
;									$LOW_UNDERLINE_SINGLE(1),
;									$LOW_UNDERLINE_DOUBLE(2),
;									$LOW_UNDERLINE_DOTTED(3),
;									$LOW_UNDERLINE_DONT_KNOW(4),
;									$LOW_UNDERLINE_DASH(5),
;									$LOW_UNDERLINE_LONG_DASH(6),
;									$LOW_UNDERLINE_DASH_DOT(7),
;									$LOW_UNDERLINE_DASH_DOT_DOT(8),
;									$LOW_UNDERLINE_SML_WAVE(9),
;									$LOW_UNDERLINE_WAVE(10),
;									$LOW_UNDERLINE_DBL_WAVE(11),
;									$LOW_UNDERLINE_BOLD(12),
;									$LOW_UNDERLINE_BOLD_DOTTED(13),
;									$LOW_UNDERLINE_BOLD_DASH(14),
;									$LOW_UNDERLINE_BOLD_LONG_DASH(15),
;									$LOW_UNDERLINE_BOLD_DASH_DOT(16),
;									$LOW_UNDERLINE_BOLD_DASH_DOT_DOT(17),
;									$LOW_UNDERLINE_BOLD_WAVE(18)
; Color Constants: $LOW_COLOR_OFF(-1),
;					$LOW_COLOR_BLACK(0),
;					$LOW_COLOR_WHITE(16777215),
;					$LOW_COLOR_LGRAY(11711154),
;					$LOW_COLOR_GRAY(8421504),
;					$LOW_COLOR_DKGRAY(3355443),
;					$LOW_COLOR_YELLOW(16776960),
;					$LOW_COLOR_GOLD(16760576),
;					$LOW_COLOR_ORANGE(16744448),
;					$LOW_COLOR_BRICK(16728064),
;					$LOW_COLOR_RED(16711680),
;					$LOW_COLOR_MAGENTA(12517441),
;					$LOW_COLOR_PURPLE(8388736),
;					$LOW_COLOR_INDIGO(5582989),
;					$LOW_COLOR_BLUE(2777241),
;					$LOW_COLOR_TEAL(1410150),
;					$LOW_COLOR_GREEN(43315),
;					$LOW_COLOR_LIME(8508442),
;					$LOW_COLOR_BROWN(9127187).
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_ConvertColorFromLong,
;					_LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleOverLine(ByRef $oCharStyle, $bWordOnly = Null, $iOverLineStyle = Null, $bOLHasColor = Null, $iOLColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharOverLine($oCharStyle, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStylePosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size.
; Syntax ........: _LOWriter_CharStylePosition(Byref $oCharStyle[, $bAutoSuper = Null[, $iSuperScript = Null[, $bAutoSub = Null[, $iSubScript = Null[, $iRelativeSize = Null]]]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $bAutoSuper          - [optional] a boolean value. Default is Null. Whether to active automatically sizing
;				   +									for SuperScript.
;                  $iSuperScript        - [optional] an integer value. Default is Null. SuperScript percentage value. See
;				   +									Remarks.
;                  $bAutoSub            - [optional] a boolean value. Default is Null. Whether to active automatically sizing
;				   +									for SubScript.
;                  $iSubScript          - [optional] an integer value. Default is Null. SubScript percentage value. See Remarks.
;                  $iRelativeSize       - [optional] an integer value. Default is Null. 1-100 percentage relative to current
;				   +									font size.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoSuper not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bAutoSub not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iSuperScript not an integer, or less than 0, higher than 100 and Not 14000.
;				   @Error 1 @Extended 7 Return 0 = $iSubScript not an integer, or less than -100, higher than 100 and Not 14000.
;				   @Error 1 @Extended 8 Return 0 = $iRelativeSize not an integer, or less than 1, higher than 100.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $iSuperScript
;				   |								2 = Error setting $iSubScript
;				   |								4 = Error setting $iRelativeSize.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Set either $iSubScript or $iSuperScript to 0 to return it to Normal setting.
;					The way LibreOffice is set up Super/SubScript are set in the same setting, Super is a positive number from
;						1 to 100 (percentage), SubScript is a negative number set to 1 to 100 percentage. For the user's
;						convenience this function accepts both positive and negative numbers for SubScript, if a positive number
;						is called for SubScript, it is automatically set to a negative. Automatic Superscript has a integer
;						value of 14000, Auto SubScript has a integer value of -14000. There is no settable setting of Automatic
;						Super/Sub Script, though one exists, it is read-only in LibreOffice, consequently I have made two
;						separate parameters to be able to determine if the user wants to automatically set SuperScript or
;						SubScript. If you set both Auto SuperScript to True and Auto SubScript to True, or $iSuperScript
;						to an integer and $iSubScript to an integer, Subscript will be set as it is the last in the
;						line to be set in this function, and thus will over-write any SuperScript settings.
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStylePosition(ByRef $oCharStyle, $bAutoSuper = Null, $iSuperScript = Null, $bAutoSub = Null, $iSubScript = Null, $iRelativeSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharPosition($oCharStyle, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStylePosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleRotateScale
; Description ...: Set or retrieve the character rotational and Scale settings for a Character Style.
; Syntax ........: _LOWriter_CharStyleRotateScale(Byref $oCharStyle[, $iRotation = Null[, $iScaleWidth = Null]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $iRotation           - [optional] an integer value. Default is Null. Degrees to rotate the text. Accepts
;				   +								only 0, 90, and 270 degrees.
;                  $iScaleWidth         - [optional] an integer value. Default is Null. The percentage to  horizontally
;				   +					stretch or compress the text. Must be above 1. Max 100. 100 is normal sizing.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iRotation not an Integer or not equal to 0, 90 or 270 degrees.
;				   @Error 1 @Extended 5 Return 0 = $iScaleWidth not an Integer or less than 1% or greater than 100%.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $iRotation
;				   |								2 = Error setting $iScaleWidth
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleRotateScale(ByRef $oCharStyle, $iRotation = Null, $iScaleWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharRotateScale($oCharStyle, $iRotation, $iScaleWidth)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleRotateScale

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleSet
; Description ...: Set a Character style for a section of text by Cursor or paragraph Object.
; Syntax ........: _LOWriter_CharStyleSet(Byref $oDoc, Byref $oObj, $sCharStyle)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oObj                - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object returned from
;				   +						_LOWriter_ParObjCreateList function.
;                  $sCharStyle          - a string value. The Character Style name.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oObj does not support Character Properties Service.
;				   @Error 1 @Extended 4 Return 0 = $sCharStyle not a String.
;				   @Error 1 @Extended 5 Return 0 = Character Style defined in $sCharStyle doesn't exist in Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Error setting Character Style.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Character Style successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor,
;					_LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_CharStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleSet(ByRef $oDoc, ByRef $oObj, $sCharStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oObj.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not _LOWriter_CharStyleExists($oDoc, $sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	$sCharStyle = __LOWriter_CharStyleNameToggle($sCharStyle)
	$oObj.CharStyleName = $sCharStyle
	Return ($oObj.CharStyleName() = $sCharStyle) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)
EndFunc   ;==>_LOWriter_CharStyleSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStylesGetNames
; Description ...: Retrieve a list of all Character Style names available for a document.
; Syntax ........: _LOWriter_CharStylesGetNames(Byref $oDoc[, $bUserOnly = False[, $bAppliedOnly = False]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True only User-Created Character
;				   +						Styles are returned.
;                  $bAppliedOnly        - [optional] a boolean value. Default is False. If True only Applied Character Styles
;				   +						are returned.
; Return values .:  Success: Integer or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bUserOnly not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bAppliedOnly not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Character Styles Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. No Character Styles found according to parameters.
;				   @Error 0 @Extended ? Return Array = Success. An Array containing all Character Styles matching the
;				   +		input parameters. @Extended contains the count of results returned. If Only a Document object is
;				   +		input, all available Character styles will be returned. Else if $bUserOnly is set to True, only
;				   +		User-Created Character Styles are returned. Else if $bAppliedOnly, only Applied Character Styles
;				   +		are returned. If Both are true then only User-Created Character styles that are applied are
;				   +		returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Five Character styles have two separate names, Footnote Characters is also internally called
;					"Footnote Symbol"; Bullets, is internally called "Bullet Symbol"; Endnote Characters is internally called
;					"Endnote Symbol"; Quotation is internally called "Citation"; and "No Character Style is internally called
;					"Standard". Either name works when setting a Character Style, but on certain functions that return a
;					Character Style Name, you may see one of these alternative names.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStylesGetNames(ByRef $oDoc, $bUserOnly = False, $bAppliedOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oStyles
	Local $aStyles[0]
	Local $iCount = 0
	Local $sExecute = ""

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oStyles = $oDoc.StyleFamilies.getByName("CharacterStyles")
	If Not IsObj($oStyles) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	ReDim $aStyles[$oStyles.getCount()]

	If Not $bUserOnly And Not $bAppliedOnly Then
		For $i = 0 To $oStyles.getCount() - 1
			$aStyles[$i] = $oStyles.getByIndex($i).DisplayName
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
		Return SetError($__LOW_STATUS_SUCCESS, $i, $aStyles)
	EndIf

	$sExecute = ($bUserOnly) ? "($oStyles.getByIndex($i).isUserDefined())" : $sExecute
	$sExecute = ($bUserOnly And $bAppliedOnly) ? ($sExecute & " And ") : $sExecute
	$sExecute = ($bAppliedOnly) ? ($sExecute & "($oStyles.getByIndex($i).isInUse())") : $sExecute

	For $i = 0 To $oStyles.getCount() - 1
		If Execute($sExecute) Then
			$aStyles[$iCount] = $oStyles.getByIndex($i).DisplayName()
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next
	ReDim $aStyles[$iCount]

	Return ($iCount = 0) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_SUCCESS, $iCount, $aStyles)
EndFunc   ;==>_LOWriter_CharStylesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleShadow
; Description ...: Set and retrieve the Shadow for a Character Style. Libre Office 4.2 and Up.
; Syntax ........: _LOWriter_CharStyleShadow(Byref $oCharStyle[, $iWidth = Null[, $iColor = Null[, $bTransparent = Null[, $iLocation = Null]]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $iWidth              - [optional] an integer value. Default is Null. Width of the shadow, set in Micrometers.
;                  $iColor              - [optional] an integer value. Default is Null. Color of the shadow. See Remarks and
;				   +							Constants below.
;                  $bTransparent        - [optional] a boolean value. Default is Null. Whether the shadow is transparent or not.
;                  $iLocation           - [optional] an integer value. Default is Null. Location of the shadow compared to the
;				   +									characters. See Constants listed below.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iColor not an Integer, or less than 0 or greater than 16777215 micrometers.
;				   @Error 1 @Extended 6 Return 0 = $bTransparent not a boolean.
;				   @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, or less than 0 or greater than 4. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shadow format Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Shadow format Object for Error Checking.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $iWidth
;				   |								2 = Error setting $iColor
;				   |								4 = Error setting $bTransparent
;				   |								8 = Error setting $iLocation
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Note: LibreOffice may adjust the set width +/- 1 Micrometer after setting.
;					Color is set in Long Integer format. You can use one of the below listed constants or a custom one.
;Shadow Location Constants: $LOW_SHADOW_NONE(0) = No shadow.
;							$LOW_SHADOW_TOP_LEFT(1) = Shadow is located along the upper and left sides.
;							$LOW_SHADOW_TOP_RIGHT(2) = Shadow is located along the upper and right sides.
;							$LOW_SHADOW_BOTTOM_LEFT(3) = Shadow is located along the lower and left sides.
;							$LOW_SHADOW_BOTTOM_RIGHT(4) = Shadow is located along the lower and right sides.
; Color Constants: $LOW_COLOR_BLACK(0),
;					$LOW_COLOR_WHITE(16777215),
;					$LOW_COLOR_LGRAY(11711154),
;					$LOW_COLOR_GRAY(8421504),
;					$LOW_COLOR_DKGRAY(3355443),
;					$LOW_COLOR_YELLOW(16776960),
;					$LOW_COLOR_GOLD(16760576),
;					$LOW_COLOR_ORANGE(16744448),
;					$LOW_COLOR_BRICK(16728064),
;					$LOW_COLOR_RED(16711680),
;					$LOW_COLOR_MAGENTA(12517441),
;					$LOW_COLOR_PURPLE(8388736),
;					$LOW_COLOR_INDIGO(5582989),
;					$LOW_COLOR_BLUE(2777241),
;					$LOW_COLOR_TEAL(1410150),
;					$LOW_COLOR_GREEN(43315),
;					$LOW_COLOR_LIME(8508442),
;					$LOW_COLOR_BROWN(9127187).
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_ConvertColorFromLong,
;					 _LOWriter_ConvertColorToLong, _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleShadow(ByRef $oCharStyle, $iWidth = Null, $iColor = Null, $bTransparent = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharShadow($oCharStyle, $iWidth, $iColor, $bTransparent, $iLocation)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning) for a Character style.
; Syntax ........: _LOWriter_CharStyleSpacing(Byref $oDoc, $sCharStyle[, $bAutoKerning = Null[, $nKerning = Null]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $bAutoKerning        - [optional] a boolean value. Default is Null. True applies a spacing in between
;				   +						certain pairs of characters. False = disabled.
;                  $nKerning            - [optional] a general number value. Default is Null. The kerning value of the
;				   +								characters. Min is -2 Pt. Max is 928.8 Pt. See Remarks. Values are in
;				   +								Printer's Points as set in the Libre Office UI.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoKerning not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $nKerning not a number, or less than -2 or greater than 928.8 Points.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bAutoKerning
;				   |								2 = Error setting $nKerning.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User
;						Display, however the internal setting is measured in MicroMeters. They will be automatically converted
;						from Points to MicroMeters and back for retrieval of settings.
;						The acceptable values are from -2 Pt to  928.8 Pt. the figures can be directly converted easily,
;						however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative
;						MicroMeters internally from 928.9 up to 1000 Pt (Max setting). For example, 928.8Pt is the last
;						correct value, which equals 32766 uM (MicroMeters), after this LibreOffice reports the following:
;						;928.9 Pt = -32766 uM; 929 Pt = -32763 uM; 929.1 = -32759; 1000 pt = -30258. Attempting to set Libre's
;						kerning value to anything over 32768 uM causes a COM exception, and attempting to set the kerning to
;						any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the
;						max settable kerning is -2.0 Pt  to 928.8 Pt.
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleSpacing(ByRef $oCharStyle, $bAutoKerning = Null, $nKerning = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharSpacing($oCharStyle, $bAutoKerning, $nKerning)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleStrikeOut
; Description ...: Set or Retrieve the StrikeOut settings for a Character style.
; Syntax ........: _LOWriter_CharStyleStrikeOut(Byref $oCharStyle[, $bWordOnly = Null[, $bStrikeOut = Null[, $iStrikeLineStyle = Null]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. Whether to strike out words only and skip
;				   +							whitespaces. True = skip whitespaces.
;                  $bStrikeOut          - [optional] a boolean value. Default is Null. True = strikeout, False = no strikeout.
;                  $iStrikeLineStyle    - [optional] an integer value. Default is Null. The Strikeout Line Style, see constants
;				   +								below.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bStrikeOut not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iStrikeLineStyle not an Integer, or less than 0 or greater than 8. Check
;				   +									the Constants list.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $bStrikeOut
;				   |								4 = Error setting $iStrikeLineStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Note Strikeout converted to single line in Ms word document format.
; Strikeout Line Style Constants : $LOW_STRIKEOUT_NONE(0); specifies not to strike out the characters.
;					$LOW_STRIKEOUT_SINGLE(1); specifies to strike out the characters with a single line
;					$LOW_STRIKEOUT_DOUBLE(2); specifies to strike out the characters with a double line.
;					$LOW_STRIKEOUT_DONT_KNOW(3); The strikeout mode is not specified.
;					$LOW_STRIKEOUT_BOLD(4); specifies to strike out the characters with a bold line.
;					$LOW_STRIKEOUT_SLASH(5); specifies to strike out the characters with slashes.
;					$LOW_STRIKEOUT_X(6); specifies to strike out the characters with X's.
; Related .......: _LOWriter_CharStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleStrikeOut(ByRef $oCharStyle, $bWordOnly = Null, $bStrikeOut = Null, $iStrikeLineStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharStrikeOut($oCharStyle, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CharStyleUnderLine
; Description ...: Set and retrieve the UnderLine settings for a Character style.
; Syntax ........: _LOWriter_CharStyleUnderLine(Byref $oCharStyle[, $bWordOnly = Null[, $iUnderLineStyle = Null[, $bULHasColor = Null[, $iULColor = Null]]]])
; Parameters ....: $oCharStyle           - [in/out] an object. A Character Style object returned by previous CharStyle Create or
;				   +						Object Retrieval function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not underlined.
;                  $iUnderLineStyle     - [optional] an integer value. Default is Null. The style of the Underline line, see
;				   +									constants listed below.
;                  $bULHasColor         - [optional] a boolean value. Default is Null. Whether the underline is colored, must
;				   +						be set to true in order to set the underline color.
;                  $iULColor            - [optional] an integer value. Default is Null. The color of the underline, set in Long
;				   +						integer format. Can be one of the constants below or a custom value. Set to
;				   +						$LOW_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCharStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCharStyle not a Character Style Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iUnderLineStyle not an Integer, or less than 0 or greater than 18. Check
;				   +									the Constants list.
;				   @Error 1 @Extended 6 Return 0 = $bULHasColor not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iULColor not an Integer, or less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   +								1 = Error setting $bWordOnly
;				   |								2 = Error setting $iUnderLineStyle
;				   |								4 = Error setting $ULHasColor
;				   |								8 = Error setting $iULColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Note: $bULHasColor must be set to true in order to set the underline color.
; UnderLine line style Constants: $LOW_UNDERLINE_NONE(0),
;									$LOW_UNDERLINE_SINGLE(1),
;									$LOW_UNDERLINE_DOUBLE(2),
;									$LOW_UNDERLINE_DOTTED(3),
;									$LOW_UNDERLINE_DONT_KNOW(4),
;									$LOW_UNDERLINE_DASH(5),
;									$LOW_UNDERLINE_LONG_DASH(6),
;									$LOW_UNDERLINE_DASH_DOT(7),
;									$LOW_UNDERLINE_DASH_DOT_DOT(8),
;									$LOW_UNDERLINE_SML_WAVE(9),
;									$LOW_UNDERLINE_WAVE(10),
;									$LOW_UNDERLINE_DBL_WAVE(11),
;									$LOW_UNDERLINE_BOLD(12),
;									$LOW_UNDERLINE_BOLD_DOTTED(13),
;									$LOW_UNDERLINE_BOLD_DASH(14),
;									$LOW_UNDERLINE_BOLD_LONG_DASH(15),
;									$LOW_UNDERLINE_BOLD_DASH_DOT(16),
;									$LOW_UNDERLINE_BOLD_DASH_DOT_DOT(17),
;									$LOW_UNDERLINE_BOLD_WAVE(18)
; Color Constants: $LOW_COLOR_OFF(-1),
;					$LOW_COLOR_BLACK(0),
;					$LOW_COLOR_WHITE(16777215),
;					$LOW_COLOR_LGRAY(11711154),
;					$LOW_COLOR_GRAY(8421504),
;					$LOW_COLOR_DKGRAY(3355443),
;					$LOW_COLOR_YELLOW(16776960),
;					$LOW_COLOR_GOLD(16760576),
;					$LOW_COLOR_ORANGE(16744448),
;					$LOW_COLOR_BRICK(16728064),
;					$LOW_COLOR_RED(16711680),
;					$LOW_COLOR_MAGENTA(12517441),
;					$LOW_COLOR_PURPLE(8388736),
;					$LOW_COLOR_INDIGO(5582989),
;					$LOW_COLOR_BLUE(2777241),
;					$LOW_COLOR_TEAL(1410150),
;					$LOW_COLOR_GREEN(43315),
;					$LOW_COLOR_LIME(8508442),
;					$LOW_COLOR_BROWN(9127187).
; Related .......: _LOWriter_CharStyleGetObj, _LOWriter_CharStyleCreate, _LOWriter_ConvertColorFromLong,
;					_LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CharStyleUnderLine(ByRef $oCharStyle, $bWordOnly = Null, $iUnderLineStyle = Null, $bULHasColor = Null, $iULColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCharStyle.supportsService("com.sun.star.style.CharacterStyle") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharUnderLine($oCharStyle, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CharStyleUnderLine