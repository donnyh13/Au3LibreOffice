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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Removing, applying etc. L.O. Calc Page Styles.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_PageStyleBackColor
; _LOCalc_PageStyleBorderColor
; _LOCalc_PageStyleBorderPadding
; _LOCalc_PageStyleBorderStyle
; _LOCalc_PageStyleBorderWidth
; _LOCalc_PageStyleCreate
; _LOCalc_PageStyleCurrent
; _LOCalc_PageStyleDelete
; _LOCalc_PageStyleExists
; _LOCalc_PageStyleFooter
; _LOCalc_PageStyleFooterBackColor
; _LOCalc_PageStyleFooterBorderColor
; _LOCalc_PageStyleFooterBorderPadding
; _LOCalc_PageStyleFooterBorderStyle
; _LOCalc_PageStyleFooterBorderWidth
; _LOCalc_PageStyleFooterCreateTextCursor
; _LOCalc_PageStyleFooterObj
; _LOCalc_PageStyleFooterShadow
; _LOCalc_PageStyleGetObj
; _LOCalc_PageStyleHeader
; _LOCalc_PageStyleHeaderBackColor
; _LOCalc_PageStyleHeaderBorderColor
; _LOCalc_PageStyleHeaderBorderPadding
; _LOCalc_PageStyleHeaderBorderStyle
; _LOCalc_PageStyleHeaderBorderWidth
; _LOCalc_PageStyleHeaderCreateTextCursor
; _LOCalc_PageStyleHeaderObj
; _LOCalc_PageStyleHeaderShadow
; _LOCalc_PageStyleLayout
; _LOCalc_PageStyleMargins
; _LOCalc_PageStyleOrganizer
; _LOCalc_PageStylePaperFormat
; _LOCalc_PageStylesGetNames
; _LOCalc_PageStyleShadow
; _LOCalc_PageStyleSheetPageOrder
; _LOCalc_PageStyleSheetPrint
; _LOCalc_PageStyleSheetScale
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleBackColor
; Description ...: Set or Retrieve background color settings for a Page style.
; Syntax ........: _LOCalc_PageStyleBackColor(ByRef $oPageStyle[, $iBackColor = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current background color.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleBackColor(ByRef $oPageStyle, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iColor

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iBackColor) Then
		$iColor = $oPageStyle.BackColor()
		If Not IsInt($iColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPageStyle.BackColor = $iBackColor
	$iError = ($oPageStyle.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleBorderColor
; Description ...: Set the Page Style Border Line Color. Libre Office Version 3.6 and Up.
; Syntax ........: _LOCalc_PageStyleBorderColor(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. The Top Border Line Color of the Page, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. The Bottom Border Line Color of the Page, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. The Left Border Line Color of the Page, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. The Right Border Line Color of the Page, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Cannot set Top Border Color when Top Border width not set.
;                  @Error 3 @Extended 3 Return 0 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error 3 @Extended 4 Return 0 = Cannot set Left Border Color when Left Border width not set.
;                  @Error 3 @Extended 5 Return 0 = Cannot set Right Border Color when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOCalc_PageStyleBorderWidth, _LOCalc_PageStyleBorderStyle, _LOCalc_PageStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleBorderColor(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOCalc_PageStyleBorder($oPageStyle, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_PageStyleBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleBorderPadding
; Description ...: Set or retrieve the Page Style Border Padding settings.
; Syntax ........: _LOCalc_PageStyleBorderPadding(ByRef $oPageStyle[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The Top Distance between the Border and Page contents in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Distance between the Border and Page contents in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Distance between the Border and Page contents in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The Right Distance between the Border and Page contents in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iAll not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iBottom not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $Left not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iRight not an Integer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAll border distance
;                  |                               2 = Error setting $iTop border distance
;                  |                               4 = Error setting $iBottom border distance
;                  |                               8 = Error setting $iLeft border distance
;                  |                               16 = Error setting $iRight border distance
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert, _LOCalc_PageStyleBorderWidth, _LOCalc_PageStyleBorderStyle, _LOCalc_PageStyleBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleBorderPadding(ByRef $oPageStyle, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oPageStyle.BorderDistance(), $oPageStyle.TopBorderDistance(), _
				$oPageStyle.BottomBorderDistance(), $oPageStyle.LeftBorderDistance(), $oPageStyle.RightBorderDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LO_IntIsBetween($iAll, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.BorderDistance = $iAll
		$iError = (__LO_IntIsBetween($oPageStyle.BorderDistance(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.TopBorderDistance = $iTop
		$iError = (__LO_IntIsBetween($oPageStyle.TopBorderDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.BottomBorderDistance = $iBottom
		$iError = (__LO_IntIsBetween($oPageStyle.BottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.LeftBorderDistance = $iLeft
		$iError = (__LO_IntIsBetween($oPageStyle.LeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.RightBorderDistance = $iRight
		$iError = (__LO_IntIsBetween($oPageStyle.RightBorderDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleBorderStyle
; Description ...: Set or Retrieve the Page Style Border Line style. Libre Office Version 3.6 and Up.
; Syntax ........: _LOCalc_PageStyleBorderStyle(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. The Top Border Line Style of the Page. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. The Bottom Border Line Style of the Page. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. The Left Border Line Style of the Page. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. The Right Border Line Style of the Page. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Cannot set Top Border Style when Top Border width not set.
;                  @Error 3 @Extended 3 Return 0 = Cannot set Bottom Border Style when Bottom Border width not set.
;                  @Error 3 @Extended 4 Return 0 = Cannot set Left Border Style when Left Border width not set.
;                  @Error 3 @Extended 5 Return 0 = Cannot set Right Border Style when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LOCalc_PageStyleBorderWidth, _LOCalc_PageStyleBorderColor, _LOCalc_PageStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleBorderStyle(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOCalc_PageStyleBorder($oPageStyle, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_PageStyleBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleBorderWidth
; Description ...: Set or Retrieve the Page Style Border Line Width. Libre Office Version 3.6 and Up.
; Syntax ........: _LOCalc_PageStyleBorderWidth(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value. Default is Null. The Top Border Line width of the Page in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Border Line Width of the Page in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Border Line width of the Page in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. The Right Border Line Width of the Page in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, or less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set Width to 0
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert, _LOCalc_PageStyleBorderStyle, _LOCalc_PageStyleBorderColor, _LOCalc_PageStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleBorderWidth(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOCalc_PageStyleBorder($oPageStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_PageStyleBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleCreate
; Description ...: Create a new Page Style in a Document.
; Syntax ........: _LOCalc_PageStyleCreate(ByRef $oDoc, $sPageStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sPageStyle          - a string value. The Name of the new Page Style to create.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sPageStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = Page Style name called in $sPageStyle already exists in document.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating "com.sun.star.style.PageStyle" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Retrieving "PageStyle" Object.
;                  @Error 3 @Extended 2 Return 0 = Error creating new Page Style by name.
;                  @Error 3 @Extended 3 Return 0 = Error Retrieving Created Page Style Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. New page Style successfully created. Returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_PageStyleDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleCreate(ByRef $oDoc, $sPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageStyles, $oStyle, $oPageStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oPageStyles = $oDoc.StyleFamilies().getByName("PageStyles")
	If Not IsObj($oPageStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If _LOCalc_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oStyle = $oDoc.createInstance("com.sun.star.style.PageStyle")
	If Not IsObj($oStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oPageStyles.insertByName($sPageStyle, $oStyle)

	If Not $oPageStyles.hasByName($sPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oPageStyle = $oPageStyles.getByName($sPageStyle)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPageStyle)
EndFunc   ;==>_LOCalc_PageStyleCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleCurrent
; Description ...: Set or Retrieve the current Page style for a Sheet.
; Syntax ........: _LOCalc_PageStyleCurrent(ByRef $oDoc, ByRef $oSheet[, $sPageStyle = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sPageStyle          - [optional] a string value. Default is Null. The Page Style name to set the Page to.
; Return values .: Success: 1 or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSheet is not a Sheet Object.
;                  @Error 1 @Extended 4 Return 0 = $sPageStyle not a String.
;                  @Error 1 @Extended 5 Return 0 = Page Style called in $sPageStyle doesn't exist in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sPageStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning current Page Style set for this Sheet.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOCalc_PageStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleCurrent(ByRef $oDoc, ByRef $oSheet, $sPageStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sCurrStyle
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSheet.supportsService("com.sun.star.sheet.Spreadsheet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($sPageStyle) Then
		$sCurrStyle = $oSheet.PageStyle()
		If Not IsString($sCurrStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrStyle)
	EndIf

	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not _LOCalc_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSheet.PageStyle = $sPageStyle
	$iError = ($oSheet.PageStyle() = $sPageStyle) ? ($iError) : (BitOR($iError, 1))

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOCalc_PageStyleCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleDelete
; Description ...: Delete a User-Created Page Style from a Document.
; Syntax ........: _LOCalc_PageStyleDelete(ByRef $oDoc, $oPageStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function. Must be User-Created, not a built-in Style native to LibreOffice.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving "PageStyles" Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Page Style Name.
;                  @Error 3 @Extended 3 Return 0 = $oPageStyle is not a User-Created Page Style and cannot be deleted.
;                  @Error 3 @Extended 4 Return 0 = $oPageStyle is in use and cannot be deleted.
;                  @Error 3 @Extended 5 Return 0 = $oPageStyle still exists after deletion attempt.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Page Style called in $oPageStyle was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleDelete(ByRef $oDoc, ByRef $oPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageStyles
	Local $sPageStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPageStyles = $oDoc.StyleFamilies().getByName("PageStyles")
	If Not IsObj($oPageStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sPageStyle = $oPageStyle.Name()
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oPageStyle.isUserDefined() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If $oPageStyle.isInUse() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; If Style is in use return an error.

	$oPageStyles.removeByName($sPageStyle)

	Return ($oPageStyles.hasByName($sPageStyle)) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleExists
; Description ...: Check whether a document contains the requested Page Style by Name.
; Syntax ........: _LOCalc_PageStyleExists(ByRef $oDoc, $sPageStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sPageStyle          - a string value. The Page Style Name to search for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object,
;                  @Error 1 @Extended 2 Return 0 = $sPageStyle not a String
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If Page Style name exists, then True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleExists(ByRef $oDoc, $sPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oDoc.StyleFamilies.getByName("PageStyles").hasByName($sPageStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOCalc_PageStyleExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleFooter
; Description ...: Modify or retrieve Footer settings for a page style.
; Syntax ........: _LOCalc_PageStyleFooter(ByRef $oPageStyle[, $bFooterOn = Null[, $bSameLeftRight = Null[, $bSameOnFirst = Null[, $iLeftMargin = Null[, $iRightMargin = Null[, $iSpacing = Null[, $iHeight = Null[, $bAutoHeight = Null]]]]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $bFooterOn           - [optional] a boolean value. Default is Null. If True, adds a footer to the page style.
;                  $bSameLeftRight      - [optional] a boolean value. Default is Null. If True, Even and odd pages share the same content.
;                  $bSameOnFirst        - [optional] a boolean value. Default is Null. If True, First and even/odd pages share the same content. LibreOffice 4.0 and up.
;                  $iLeftMargin         - [optional] an integer value. Default is Null. The amount of space to leave between the left edge of the page and the left edge of the footer. Set in Hundredths of a Millimeter (HMM).
;                  $iRightMargin        - [optional] an integer value. Default is Null. The amount of space to leave between the right edge of the page and the right edge of the footer. Set in Hundredths of a Millimeter (HMM).
;                  $iSpacing            - [optional] an integer value. Default is Null. The amount of space that you want to maintain between the bottom edge of the document text and the top edge of the footer. Set in Hundredths of a Millimeter (HMM).
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the footer. Set in Hundredths of a Millimeter (HMM).
;                  $bAutoHeight         - [optional] a boolean value. Default is Null. If True, automatically adjusts the height of the footer to fit the contents.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $bFooterOn not a Boolean value.
;                  @Error 1 @Extended 4 Return 0 = $bSameLeftRight not a Boolean value.
;                  @Error 1 @Extended 5 Return 0 = $bSameOnFirst not a Boolean value.
;                  @Error 1 @Extended 6 Return 0 = $iLeftMargin not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iRightMargin not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iSpacing not an Integer.
;                  @Error 1 @Extended 9 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 10 Return 0 = $bAutoHeight not a Boolean value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bFooterOn
;                  |                               2 = Error setting $bSameLeftRight
;                  |                               4 = Error setting $bSameOnFirst
;                  |                               8 = Error setting $iLeftMargin
;                  |                               16 = Error setting $iRightMargin
;                  |                               32 = Error setting $iSpacing
;                  |                               64 = Error setting $iHeight
;                  |                               128 = Error setting $bAutoHeight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 or 8 Element Array with values in order of function parameters. If Libre Office version is less than 4.0, then the Array returned will contain 7 elements, because $bSameOnFirst will not be available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleFooter(ByRef $oPageStyle, $bFooterOn = Null, $bSameLeftRight = Null, $bSameOnFirst = Null, $iLeftMargin = Null, $iRightMargin = Null, $iSpacing = Null, $iHeight = Null, $bAutoHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFooter[7]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bFooterOn, $bSameLeftRight, $bSameOnFirst, $iLeftMargin, $iRightMargin, $iSpacing, $iHeight, $bAutoHeight) Then
		If __LO_VersionCheck(4.0) Then
			__LO_ArrayFill($avFooter, $oPageStyle.FooterIsOn(), $oPageStyle.FooterIsShared(), $oPageStyle.FirstPageFooterIsShared(), $oPageStyle.FooterLeftMargin(), _
					$oPageStyle.FooterRightMargin(), $oPageStyle.FooterBodyDistance(), $oPageStyle.FooterHeight(), $oPageStyle.FooterIsDynamicHeight())

		Else
			__LO_ArrayFill($avFooter, $oPageStyle.FooterIsOn(), $oPageStyle.FooterIsShared(), $oPageStyle.FooterLeftMargin(), _
					$oPageStyle.FooterRightMargin(), $oPageStyle.FooterBodyDistance(), $oPageStyle.FooterHeight(), $oPageStyle.FooterIsDynamicHeight())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFooter)
	EndIf

	If ($bFooterOn <> Null) Then
		If Not IsBool($bFooterOn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.FooterIsOn = $bFooterOn
		$iError = ($oPageStyle.FooterIsOn() = $bFooterOn) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bSameLeftRight <> Null) Then
		If Not IsBool($bSameLeftRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.FooterIsShared = $bSameLeftRight
		$iError = ($oPageStyle.FooterIsShared() = $bSameLeftRight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bSameOnFirst <> Null) Then
		If Not IsBool($bSameOnFirst) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not __LO_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oPageStyle.FirstPageFooterIsShared = $bSameOnFirst
		$iError = ($oPageStyle.FirstPageFooterIsShared() = $bSameOnFirst) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeftMargin <> Null) Then
		If Not IsInt($iLeftMargin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.FooterLeftMargin = $iLeftMargin
		$iError = (__LO_IntIsBetween($oPageStyle.FooterLeftMargin(), $iLeftMargin - 1, $iLeftMargin + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRightMargin <> Null) Then
		If Not IsInt($iRightMargin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.FooterRightMargin = $iRightMargin
		$iError = (__LO_IntIsBetween($oPageStyle.FooterRightMargin(), $iRightMargin - 1, $iRightMargin + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iSpacing <> Null) Then
		If Not IsInt($iSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPageStyle.FooterBodyDistance = $iSpacing
		$iError = (__LO_IntIsBetween($oPageStyle.FooterBodyDistance(), $iSpacing - 1, $iSpacing + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iHeight <> Null) Then
		If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oPageStyle.FooterHeight = $iHeight
		$iError = (__LO_IntIsBetween($oPageStyle.FooterHeight(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bAutoHeight <> Null) Then
		If Not IsBool($bAutoHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oPageStyle.FooterIsDynamicHeight = $bAutoHeight
		$iError = ($oPageStyle.FooterIsDynamicHeight() = $bAutoHeight) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleFooterBackColor
; Description ...: Set or Retrieve background color settings for a Page style Footer.
; Syntax ........: _LOCalc_PageStyleFooterBackColor(ByRef $oPageStyle[, $iBackColor = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Footers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current background color.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleFooterBackColor(ByRef $oPageStyle, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iColor

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iBackColor) Then
		$iColor = $oPageStyle.FooterBackColor()
		If Not IsInt($iColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPageStyle.FooterBackColor = $iBackColor
	$iError = ($oPageStyle.FooterBackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleFooterBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleFooterBorderColor
; Description ...: Set and Retrieve the Page Style Footer Border Line Color.
; Syntax ........: _LOCalc_PageStyleFooterBorderColor(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. The Top Border Line Color of the Page Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. The Bottom Border Line Color of the Page Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. The Left Border Line Color of the Page Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. The Right Border Line Color of the Page Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Footers are not enabled for this Page Style.
;                  @Error 3 @Extended 3 Return 0 = Cannot set Top Border Color when Top Border width not set.
;                  @Error 3 @Extended 4 Return 0 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error 3 @Extended 5 Return 0 = Cannot set Left Border Color when Left Border width not set.
;                  @Error 3 @Extended 6 Return 0 = Cannot set Right Border Color when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOCalc_PageStyleFooterBorderWidth, _LOCalc_PageStyleFooterBorderStyle, _LOCalc_PageStyleFooterBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleFooterBorderColor(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOCalc_PageStyleFooterBorder($oPageStyle, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_PageStyleFooterBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleFooterBorderPadding
; Description ...: Set or retrieve the Footer Border Padding settings.
; Syntax ........: _LOCalc_PageStyleFooterBorderPadding(ByRef $oPageStyle[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The Top Distance between the Border and Page contents in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Distance between the Border and Page contents in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Distance between the Border and Page contents in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The Right Distance between the Border and Page contents in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array, see Remarks.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iAll not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iBottom not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $Left not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iRight not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Footers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAll border distance
;                  |                               2 = Error setting $iTop border distance
;                  |                               4 = Error setting $iBottom border distance
;                  |                               8 = Error setting $iLeft border distance
;                  |                               16 = Error setting $iRight border distance
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert, _LOCalc_PageStyleFooterBorderWidth, _LOCalc_PageStyleFooterBorderStyle, _LOCalc_PageStyleFooterBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleFooterBorderPadding(ByRef $oPageStyle, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oPageStyle.FooterBorderDistance(), $oPageStyle.FooterTopBorderDistance(), _
				$oPageStyle.FooterBottomBorderDistance(), $oPageStyle.FooterLeftBorderDistance(), $oPageStyle.FooterRightBorderDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not (IsInt($iAll) Or ($iAll > 0)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.FooterBorderDistance = $iAll
		$iError = (__LO_IntIsBetween($oPageStyle.FooterBorderDistance(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTop <> Null) Then
		If Not (IsInt($iTop) Or ($iTop > 0)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.FooterTopBorderDistance = $iTop
		$iError = (__LO_IntIsBetween($oPageStyle.FooterTopBorderDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iBottom <> Null) Then
		If Not (IsInt($iBottom) Or ($iBottom > 0)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.FooterBottomBorderDistance = $iBottom
		$iError = (__LO_IntIsBetween($oPageStyle.FooterBottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeft <> Null) Then
		If Not (IsInt($iLeft) Or ($iLeft > 0)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.FooterLeftBorderDistance = $iLeft
		$iError = (__LO_IntIsBetween($oPageStyle.FooterLeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRight <> Null) Then
		If Not (IsInt($iRight) Or ($iRight > 0)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.FooterRightBorderDistance = $iRight
		$iError = (__LO_IntIsBetween($oPageStyle.FooterRightBorderDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleFooterBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleFooterBorderStyle
; Description ...: Set and retrieve the Page Style Footer Border Line style.
; Syntax ........: _LOCalc_PageStyleFooterBorderStyle(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. The Top Border Line Style of the Page Style. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. The Bottom Border Line Style of the Page Style. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. The Left Border Line Style of the Page Style. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. The Right Border Line Style of the Page Style. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Footers are not enabled for this Page Style.
;                  @Error 3 @Extended 3 Return 0 = Cannot set Top Border Style when Top Border width not set.
;                  @Error 3 @Extended 4 Return 0 = Cannot set Bottom Border Style when Bottom Border width not set.
;                  @Error 3 @Extended 5 Return 0 = Cannot set Left Border Style when Left Border width not set.
;                  @Error 3 @Extended 6 Return 0 = Cannot set Right Border Style when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LOCalc_PageStyleFooterBorderWidth, _LOCalc_PageStyleFooterBorderColor, _LOCalc_PageStyleFooterBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleFooterBorderStyle(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOCalc_PageStyleFooterBorder($oPageStyle, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_PageStyleFooterBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleFooterBorderWidth
; Description ...: Set and retrieve the Page Style Footer Border Line Width.
; Syntax ........: _LOCalc_PageStyleFooterBorderWidth(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value. Default is Null. The Top Border Line width of the Page Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Border Line Width of the Page Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Border Line width of the Page Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. The Right Border Line Width of the Page Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, or less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Footers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set Width to 0.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert, _LOCalc_PageStyleFooterBorderStyle, _LOCalc_PageStyleFooterBorderColor, _LOCalc_PageStyleFooterBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleFooterBorderWidth(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOCalc_PageStyleFooterBorder($oPageStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_PageStyleFooterBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleFooterCreateTextCursor
; Description ...: Create a Text Cursor in a Footer area.
; Syntax ........: _LOCalc_PageStyleFooterCreateTextCursor(ByRef $oFooter[, $bAtEnd = False[, $bLeftArea = False[, $bCenterArea = False[, $bRightArea = False]]]])
; Parameters ....: $oFooter             - [in/out] an object. A Footer Object from a previous _LOCalc_PageStyleFooterObj function.
;                  $bAtEnd              - [optional] a boolean value. Default is False. If True, the Text Cursor is created at the end of the Text, else it will be created at the beginning.
;                  $bLeftArea           - [optional] a boolean value. Default is False. If True, the Text Cursor will be created in the Left Area of the Footer.
;                  $bCenterArea         - [optional] a boolean value. Default is False. If True, the Text Cursor will be created in the Center Area of the Footer.
;                  $bRightArea          - [optional] a boolean value. Default is False. If True, the Text Cursor will be created in the Right Area of the Footer.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFooter not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFooter not a Header/Footer Object.
;                  @Error 1 @Extended 3 Return 0 = $bAtEnd not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bLeftArea not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bCenterArea not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bRightArea not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = Either more than one of the following are called with True, or all are False, $bLeftArea, $bCenterArea, $bRightArea.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Text Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully created a Text Cursor in the requested Footer area, returning the Text Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You can only create a Text Cursor in one area at a time per function call.
; Related .......: _LOCalc_TextCursorMove, _LOCalc_PageStyleFooterObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleFooterCreateTextCursor(ByRef $oFooter, $bAtEnd = False, $bLeftArea = False, $bCenterArea = False, $bRightArea = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oFooter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oFooter.supportsService("com.sun.star.sheet.HeaderFooterContent")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAtEnd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bLeftArea) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bCenterArea) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bRightArea) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($bLeftArea + $bCenterArea + $bRightArea <> 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0) ; More than one True, or all False.

	If $bLeftArea Then
		$oTextCursor = $oFooter.LeftText.createTextCursor()

	ElseIf $bCenterArea Then
		$oTextCursor = $oFooter.CenterText.createTextCursor()

	ElseIf $bRightArea Then
		$oTextCursor = $oFooter.RightText.createTextCursor()
	EndIf

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $bAtEnd Then
		$oTextCursor.gotoEnd(False)

	Else
		$oTextCursor.gotoStart(False)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOCalc_PageStyleFooterCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleFooterObj
; Description ...: Set or Retrieve the Object for the Page Style Footer Object. See Remarks.
; Syntax ........: _LOCalc_PageStyleFooterObj(ByRef $oPageStyle[, $oFirstPage = Null[, $oRightPage = Null[, $oLeftPage = Null]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $oFirstPage          - [optional] an object. Default is Null. Set or Retrieve the First Page Footer Object. Call with Default Keyword to retrieve the Object, else Call with the modified Object to set the new content.
;                  $oRightPage          - [optional] an object. Default is Null. Set or Retrieve the Right Page Footer Object. Call with Default Keyword to retrieve the Object, else Call with the modified Object to set the new content.
;                  $oLeftPage           - [optional] an object. Default is Null. Set or Retrieve the Left Page Footer Object. Call with Default Keyword to retrieve the Object, else Call with the modified Object to set the new content.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFirstPage not a keyword (Null or Default) and not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oRightPage not a keyword (Null or Default) and not an Object.
;                  @Error 1 @Extended 4 Return 0 = $oLeftPage not a keyword (Null or Default) and not an Object.
;                  @Error 1 @Extended 5 Return 0 = None of the Parameters ($oFirstPage, $oRightPage, $oLeftPage) are called with other than Null.
;                  @Error 1 @Extended 6 Return 0 = Object called in $oFirstPage not a Header/Footer Object.
;                  @Error 1 @Extended 7 Return 0 = Object called in $oRightPage not a Header/Footer Object.
;                  @Error 1 @Extended 8 Return 0 = Object called in $oLeftPage not a Header/Footer Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve First Page Footer Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Right Page Footer Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Left Page Footer Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Object = Success. One of the three parameters ($oFirstPage, $oRightPage, $oLeftPage) was called with Default keyword, returning the specified Footer Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office Calc Footers are set up in a way that you must retrieve the Object for the one you wish to modify the content of, modify it, and then re-insert the modified object.
;                  To modify the Header/Footer content, retrieve the desired Page (First, Left/Right) by calling the respective parameter with Default keyword, then create a text cursor in the desired area by calling _LOCalc_PageStyleFooterCreateTextCursor with the Header/Footer Object, insert, delete, etc. the content, then call this function again with the Object you retrieved from it earlier, in the appropriate parameter.
;                  The Object returned is interchangeable among the others. i.e. if you want identical content on all pages, you can set all three using the same Object. You can also use this method to copy header/footer content from one page Style to another, even in other documents.
;                  Only one Object can be retrieved at once. But you could set and retrieve another object at one time, such as set First page and retrieve Left Page. If more than one parameter is called with Default, the first parameter called with Default is retrieved.
;                  If Same Content on Left and Right is True, enter Content using RightPage Object.
; Related .......: _LOCalc_PageStyleHeaderObj, _LOCalc_PageStyleFooterCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleFooterObj(ByRef $oPageStyle, $oFirstPage = Null, $oRightPage = Null, $oLeftPage = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFooter

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsKeyword($oFirstPage) And Not IsObj($oFirstPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsKeyword($oRightPage) And Not IsObj($oRightPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsKeyword($oLeftPage) And Not IsObj($oLeftPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If __LO_VarsAreNull($oFirstPage, $oRightPage, $oLeftPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If ($oFirstPage = Default) Then
		$oFooter = $oPageStyle.FirstPageFooterContent()
		If Not IsObj($oFooter) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ElseIf IsObj($oFirstPage) Then
		If Not ($oFirstPage.supportsService("com.sun.star.sheet.HeaderFooterContent")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.FirstPageFooterContent = $oFirstPage
	EndIf

	If ($oRightPage = Default) Then
		If Not IsObj($oFooter) Then ; Only retrieve the Object if I haven't retrieved one already.
			$oFooter = $oPageStyle.RightPageFooterContent()
			If Not IsObj($oFooter) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndIf

	ElseIf IsObj($oRightPage) Then
		If Not ($oRightPage.supportsService("com.sun.star.sheet.HeaderFooterContent")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.RightPageFooterContent = $oRightPage
	EndIf

	If ($oLeftPage = Default) Then
		If Not IsObj($oFooter) Then ; Only retrieve the Object if I haven't retrieved one already.
			$oFooter = $oPageStyle.LeftPageFooterContent()
			If Not IsObj($oFooter) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		EndIf

	ElseIf IsObj($oLeftPage) Then
		If Not ($oLeftPage.supportsService("com.sun.star.sheet.HeaderFooterContent")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPageStyle.LeftPageFooterContent = $oLeftPage
	EndIf

	Return IsObj($oFooter) ? (SetError($__LO_STATUS_SUCCESS, 0, $oFooter)) : SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_PageStyleFooterObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleFooterShadow
; Description ...: Set or Retrieve the shadow settings for a Page Style Footer.
; Syntax ........: _LOCalc_PageStyleFooterShadow(ByRef $oPageStyle[, $iWidth = Null[, $iColor = Null[, $iLocation = Null]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Shadow Width of the footer, set in Hundredths of a Millimeter (HMM).
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Color of the Footer shadow, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Footer Shadow. See Constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Footers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving ShadowFormat Object.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving ShadowFormat Object for Error checking.
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
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleFooterShadow(ByRef $oPageStyle, $iWidth = Null, $iColor = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tShdwFrmt
	Local $iError = 0
	Local $avShadow[4]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tShdwFrmt = $oPageStyle.FooterShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.Location())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Or ($iWidth < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tShdwFrmt.ShadowWidth = $iWidth
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tShdwFrmt.Color = $iColor
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOC_SHADOW_NONE, $LOC_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oPageStyle.FooterShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oPageStyle.FooterShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($tShdwFrmt.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iColor)) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iLocation)) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 4)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleFooterShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleGetObj
; Description ...: Retrieve a Page Style Object for use with other Page Style functions.
; Syntax ........: _LOCalc_PageStyleGetObj(ByRef $oDoc, $sPageStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sPageStyle          - a string value. The Page Style name to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sPageStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = Page Style called in $sPageStyle not found in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Page Style Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Page Style successfully retrieved, returning Page Style Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_PageStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleGetObj(ByRef $oDoc, $sPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOCalc_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPageStyle = $oDoc.StyleFamilies().getByName("PageStyles").getByName($sPageStyle)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPageStyle)
EndFunc   ;==>_LOCalc_PageStyleGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleHeader
; Description ...: Modify or retrieve Header settings for a page style.
; Syntax ........: _LOCalc_PageStyleHeader(ByRef $oPageStyle[, $bHeaderOn = Null[, $bSameLeftRight = Null[, $bSameOnFirst = Null[, $iLeftMargin = Null[, $iRightMargin = Null[, $iSpacing = Null[, $iHeight = Null[, $bAutoHeight = Null]]]]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $bHeaderOn           - [optional] a boolean value. Default is Null. If True, adds a Header to the page style.
;                  $bSameLeftRight      - [optional] a boolean value. Default is Null. If True, Even and odd pages share the same content.
;                  $bSameOnFirst        - [optional] a boolean value. Default is Null. If True, First and even/odd pages share the same content. LibreOffice 4.0 and up.
;                  $iLeftMargin         - [optional] an integer value. Default is Null. The amount of space to leave between the left edge of the page and the left edge of the Header. Set in Hundredths of a Millimeter (HMM).
;                  $iRightMargin        - [optional] an integer value. Default is Null. The amount of space to leave between the right edge of the page and the right edge of the Header. Set in Hundredths of a Millimeter (HMM).
;                  $iSpacing            - [optional] an integer value. Default is Null. The amount of space to maintain between the Top edge of the document text and the bottom edge of the Header. Set in Hundredths of a Millimeter (HMM).
;                  $iHeight             - [optional] an integer value. Default is Null. The height for the Header. Set in Hundredths of a Millimeter (HMM).
;                  $bAutoHeight         - [optional] a boolean value. Default is Null. If True, Automatically adjusts the height of the Header to fit the contents.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $bHeaderOn not a Boolean value.
;                  @Error 1 @Extended 4 Return 0 = $bSameLeftRight not a Boolean value.
;                  @Error 1 @Extended 5 Return 0 = $bSameOnFirst not a Boolean value.
;                  @Error 1 @Extended 6 Return 0 = $iLeftMargin not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iRightMargin not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iSpacing not an Integer.
;                  @Error 1 @Extended 9 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 10 Return 0 = $bAutoHeight not a Boolean value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHeaderOn
;                  |                               2 = Error setting $bSameLeftRight
;                  |                               4 = Error setting $bSameOnFirst
;                  |                               8 = Error setting $iLeftMargin
;                  |                               16 = Error setting $iRightMargin
;                  |                               32 = Error setting $iSpacing
;                  |                               64 = Error setting $iHeight
;                  |                               128 = Error setting $bAutoHeight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 or 8 Element Array with values in order of function parameters. If Libre Office version is less than 4.0, then the Array returned will contain 7 elements, because $bSameOnFirst will not be available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleHeader(ByRef $oPageStyle, $bHeaderOn = Null, $bSameLeftRight = Null, $bSameOnFirst = Null, $iLeftMargin = Null, $iRightMargin = Null, $iSpacing = Null, $iHeight = Null, $bAutoHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHeader[7]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bHeaderOn, $bSameLeftRight, $bSameOnFirst, $iLeftMargin, $iRightMargin, $iSpacing, $iHeight, $bAutoHeight) Then
		If __LO_VersionCheck(4.0) Then
			__LO_ArrayFill($avHeader, $oPageStyle.HeaderIsOn(), $oPageStyle.HeaderIsShared(), $oPageStyle.FirstPageHeaderIsShared(), $oPageStyle.HeaderLeftMargin(), _
					$oPageStyle.HeaderRightMargin(), $oPageStyle.HeaderBodyDistance(), $oPageStyle.HeaderHeight(), $oPageStyle.HeaderIsDynamicHeight())

		Else
			__LO_ArrayFill($avHeader, $oPageStyle.HeaderIsOn(), $oPageStyle.HeaderIsShared(), $oPageStyle.HeaderLeftMargin(), _
					$oPageStyle.HeaderRightMargin(), $oPageStyle.HeaderBodyDistance(), $oPageStyle.HeaderHeight(), $oPageStyle.HeaderIsDynamicHeight())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avHeader)
	EndIf

	If ($bHeaderOn <> Null) Then
		If Not IsBool($bHeaderOn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.HeaderIsOn = $bHeaderOn
		$iError = ($oPageStyle.HeaderIsOn() = $bHeaderOn) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bSameLeftRight <> Null) Then
		If Not IsBool($bSameLeftRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.HeaderIsShared = $bSameLeftRight
		$iError = ($oPageStyle.HeaderIsShared() = $bSameLeftRight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bSameOnFirst <> Null) Then
		If Not IsBool($bSameOnFirst) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not __LO_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oPageStyle.FirstPageHeaderIsShared = $bSameOnFirst
		$iError = ($oPageStyle.FirstPageHeaderIsShared() = $bSameOnFirst) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeftMargin <> Null) Then
		If Not IsInt($iLeftMargin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.HeaderLeftMargin = $iLeftMargin
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderLeftMargin(), $iLeftMargin - 1, $iLeftMargin + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRightMargin <> Null) Then
		If Not IsInt($iRightMargin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.HeaderRightMargin = $iRightMargin
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderRightMargin(), $iRightMargin - 1, $iRightMargin + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iSpacing <> Null) Then
		If Not IsInt($iSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPageStyle.HeaderBodyDistance = $iSpacing
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderBodyDistance(), $iSpacing - 1, $iSpacing + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iHeight <> Null) Then
		If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oPageStyle.HeaderHeight = $iHeight
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderHeight(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bAutoHeight <> Null) Then
		If Not IsBool($bAutoHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oPageStyle.HeaderIsDynamicHeight = $bAutoHeight
		$iError = ($oPageStyle.HeaderIsDynamicHeight() = $bAutoHeight) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleHeader

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleHeaderBackColor
; Description ...: Set or Retrieve background color settings for a Page style header.
; Syntax ........: _LOCalc_PageStyleHeaderBackColor(ByRef $oPageStyle[, $iBackColor = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Headers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current background color.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleHeaderBackColor(ByRef $oPageStyle, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iColor

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iBackColor) Then
		$iColor = $oPageStyle.HeaderBackColor()
		If Not IsInt($iColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPageStyle.HeaderBackColor = $iBackColor
	$iError = ($oPageStyle.HeaderBackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleHeaderBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleHeaderBorderColor
; Description ...: Set and Retrieve the Page Style Header Border Line Color.
; Syntax ........: _LOCalc_PageStyleHeaderBorderColor(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. The Top Border Line Color of the Page Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. The Bottom Border Line Color of the Page Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. The Left Border Line Color of the Page Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. The Right Border Line Color of the Page Style, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Headers are not enabled for this Page Style.
;                  @Error 3 @Extended 3 Return 0 = Cannot set Top Border Color when Top Border width not set.
;                  @Error 3 @Extended 4 Return 0 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error 3 @Extended 5 Return 0 = Cannot set Left Border Color when Left Border width not set.
;                  @Error 3 @Extended 6 Return 0 = Cannot set Right Border Color when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOCalc_PageStyleHeaderBorderWidth, _LOCalc_PageStyleHeaderBorderStyle, _LOCalc_PageStyleHeaderBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleHeaderBorderColor(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOCalc_PageStyleHeaderBorder($oPageStyle, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_PageStyleHeaderBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleHeaderBorderPadding
; Description ...: Set or retrieve the Header Border Padding settings.
; Syntax ........: _LOCalc_PageStyleHeaderBorderPadding(ByRef $oPageStyle[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The Top Distance between the Border and Page Header contents in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Distance between the Border and Page Header contents in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Distance between the Border and Page Header contents in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The Right Distance between the Border and Page Header contents in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iAll not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iBottom not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $Left not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iRight not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Headers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAll border distance
;                  |                               2 = Error setting $iTop border distance
;                  |                               4 = Error setting $iBottom border distance
;                  |                               8 = Error setting $iLeft border distance
;                  |                               16 = Error setting $iRight border distance
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert, _LOCalc_PageStyleHeaderBorderWidth, _LOCalc_PageStyleHeaderBorderStyle, _LOCalc_PageStyleHeaderBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleHeaderBorderPadding(ByRef $oPageStyle, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oPageStyle.HeaderBorderDistance(), $oPageStyle.HeaderTopBorderDistance(), _
				$oPageStyle.HeaderBottomBorderDistance(), $oPageStyle.HeaderLeftBorderDistance(), $oPageStyle.HeaderRightBorderDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LO_IntIsBetween($iAll, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.HeaderBorderDistance = $iAll
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderBorderDistance(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.HeaderTopBorderDistance = $iTop
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderTopBorderDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.HeaderBottomBorderDistance = $iBottom
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderBottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.HeaderLeftBorderDistance = $iLeft
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderLeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.HeaderRightBorderDistance = $iRight
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderRightBorderDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleHeaderBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleHeaderBorderStyle
; Description ...: Set and retrieve the Page Style Header Border Line style.
; Syntax ........: _LOCalc_PageStyleHeaderBorderStyle(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. The Top Border Line Style of the Page Style. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. The Bottom Border Line Style of the Page Style. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. The Left Border Line Style of the Page Style. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. The Right Border Line Style of the Page Style. See Constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See constants, $LOC_BORDERSTYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Headers are not enabled for this Page Style.
;                  @Error 3 @Extended 3 Return 0 = Cannot set Top Border Style Top when Border width not set.
;                  @Error 3 @Extended 4 Return 0 = Cannot set Bottom Border Style Bottom when Border width not set.
;                  @Error 3 @Extended 5 Return 0 = Cannot set Left Border Style when Left Border width not set.
;                  @Error 3 @Extended 6 Return 0 = Cannot set Right Border Style when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LOCalc_PageStyleHeaderBorderWidth, _LOCalc_PageStyleHeaderBorderColor, _LOCalc_PageStyleHeaderBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleHeaderBorderStyle(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOC_BORDERSTYLE_SOLID, $LOC_BORDERSTYLE_DASH_DOT_DOT, "", $LOC_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOCalc_PageStyleHeaderBorder($oPageStyle, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_PageStyleHeaderBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleHeaderBorderWidth
; Description ...: Set and retrieve the Page Style Header Border Line Width.
; Syntax ........: _LOCalc_PageStyleHeaderBorderWidth(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value. Default is Null. The Top Border Line width of the Page Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Border Line Width of the Page Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Border Line width of the Page Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. The Right Border Line Width of the Page Style in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOC_BORDERWIDTH_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, or less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Headers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set Width to 0.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert, _LOCalc_PageStyleHeaderBorderStyle, _LOCalc_PageStyleHeaderBorderColor, _LOCalc_PageStyleHeaderBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleHeaderBorderWidth(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOCalc_PageStyleHeaderBorder($oPageStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_PageStyleHeaderBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleHeaderCreateTextCursor
; Description ...: Create a Text Cursor in a Header area.
; Syntax ........: _LOCalc_PageStyleHeaderCreateTextCursor(ByRef $oHeader[, $bAtEnd = False[, $bLeftArea = False[, $bCenterArea = False[, $bRightArea = False]]]])
; Parameters ....: $oHeader             - [in/out] an object. A Header Object from a previous _LOCalc_PageStyleHeaderObj function.
;                  $bAtEnd              - [optional] a boolean value. Default is False. If True, the Text Cursor is created at the end of the Text, else it will be created at the beginning.
;                  $bLeftArea           - [optional] a boolean value. Default is False. If True, the Text Cursor will be created in the Left Area of the Header.
;                  $bCenterArea         - [optional] a boolean value. Default is False. If True, the Text Cursor will be created in the Center Area of the Header.
;                  $bRightArea          - [optional] a boolean value. Default is False. If True, the Text Cursor will be created in the Right Area of the Header.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oHeader not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oHeader not a Header/Footer Object.
;                  @Error 1 @Extended 3 Return 0 = $bAtEnd not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bLeftArea not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bCenterArea not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bRightArea not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = Either more than one of the following are called with True, or all are False, $bLeftArea, $bCenterArea, $bRightArea.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Text Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully created a Text Cursor in the requested Header area, returning the Text Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You can only create a Text Cursor in one area at a time per function call. i.e. only one of the Area parameters can be called with True.
; Related .......: _LOCalc_PageStyleHeaderObj, _LOCalc_TextCursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleHeaderCreateTextCursor(ByRef $oHeader, $bAtEnd = False, $bLeftArea = False, $bCenterArea = False, $bRightArea = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oHeader) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oHeader.supportsService("com.sun.star.sheet.HeaderFooterContent")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAtEnd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bLeftArea) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bCenterArea) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bRightArea) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($bLeftArea + $bCenterArea + $bRightArea <> 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0) ; More than one True, or all False.

	If $bLeftArea Then
		$oTextCursor = $oHeader.LeftText.createTextCursor()

	ElseIf $bCenterArea Then
		$oTextCursor = $oHeader.CenterText.createTextCursor()

	ElseIf $bRightArea Then
		$oTextCursor = $oHeader.RightText.createTextCursor()
	EndIf

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $bAtEnd Then
		$oTextCursor.gotoEnd(False)

	Else
		$oTextCursor.gotoStart(False)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOCalc_PageStyleHeaderCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleHeaderObj
; Description ...: Set or Retrieve the Object for the Page Style Header Object. See Remarks.
; Syntax ........: _LOCalc_PageStyleHeaderObj(ByRef $oPageStyle[, $oFirstPage = Null[, $oRightPage = Null[, $oLeftPage = Null]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $oFirstPage          - [optional] an object. Default is Null. Set or Retrieve the First Page Header Object. Call with Default Keyword to retrieve the Object, else Call with the modified Object to set the new content.
;                  $oRightPage          - [optional] an object. Default is Null. Set or Retrieve the Right Page Header Object. Call with Default Keyword to retrieve the Object, else Call with the modified Object to set the new content.
;                  $oLeftPage           - [optional] an object. Default is Null. Set or Retrieve the Left Page Header Object. Call with Default Keyword to retrieve the Object, else Call with the modified Object to set the new content.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFirstPage not a keyword (Null or Default) and not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oRightPage not a keyword (Null or Default) and not an Object.
;                  @Error 1 @Extended 4 Return 0 = $oLeftPage not a keyword (Null or Default) and not an Object.
;                  @Error 1 @Extended 5 Return 0 = None of the Parameters ($oFirstPage, $oRightPage, $oLeftPage) are called with other than Null.
;                  @Error 1 @Extended 6 Return 0 = Object called in $oFirstPage not a Header/Footer Object.
;                  @Error 1 @Extended 7 Return 0 = Object called in $oRightPage not a Header/Footer Object.
;                  @Error 1 @Extended 8 Return 0 = Object called in $oLeftPage not a Header/Footer Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve First Page Header Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Right Page Header Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Left Page Header Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Object = Success. One of the three parameters ($oFirstPage, $oRightPage, $oLeftPage) was called with Default keyword, returning the specified Header Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office Calc Headers are set up in a way that you must retrieve the Object for the one you wish to modify the content of, modify it, and then re-insert the modified object.
;                  To modify the Header/Footer content, retrieve the desired Page (First, Left/Right) by calling the respective parameter with Default keyword, then create a text cursor in the desired area by calling _LOCalc_PageStyleHeaderCreateTextCursor with the Header/Footer Object, insert, delete, etc. the content, then call this function again with the Object you retrieved from it earlier, in the appropriate parameter.
;                  The Object returned is interchangeable among the others. i.e. if you want identical content on all pages, you can set all three using the same Object. You can also use this method to copy header/footer content from one page Style to another, even in other documents.
;                  Only one Object can be retrieved at once. But you could set and retrieve another object at one time, such as set First page and retrieve Left Page. If more than one parameter is called with Default, the first parameter called with Default is retrieved.
;                  If Same Content on Left and Right is True, enter Content using RightPage Object.
; Related .......: _LOCalc_PageStyleHeaderCreateTextCursor, _LOCalc_PageStyleFooterObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleHeaderObj(ByRef $oPageStyle, $oFirstPage = Null, $oRightPage = Null, $oLeftPage = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oHeader

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsKeyword($oFirstPage) And Not IsObj($oFirstPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsKeyword($oRightPage) And Not IsObj($oRightPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsKeyword($oLeftPage) And Not IsObj($oLeftPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If __LO_VarsAreNull($oFirstPage, $oRightPage, $oLeftPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If ($oFirstPage = Default) Then
		$oHeader = $oPageStyle.FirstPageHeaderContent()
		If Not IsObj($oHeader) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ElseIf IsObj($oFirstPage) Then
		If Not ($oFirstPage.supportsService("com.sun.star.sheet.HeaderFooterContent")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.FirstPageHeaderContent = $oFirstPage
	EndIf

	If ($oRightPage = Default) Then
		If Not IsObj($oHeader) Then ; Only retrieve the Object if I haven't retrieved one already.
			$oHeader = $oPageStyle.RightPageHeaderContent()
			If Not IsObj($oHeader) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndIf

	ElseIf IsObj($oRightPage) Then
		If Not ($oRightPage.supportsService("com.sun.star.sheet.HeaderFooterContent")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.RightPageHeaderContent = $oRightPage
	EndIf

	If ($oLeftPage = Default) Then
		If Not IsObj($oHeader) Then ; Only retrieve the Object if I haven't retrieved one already.
			$oHeader = $oPageStyle.LeftPageHeaderContent()
			If Not IsObj($oHeader) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		EndIf

	ElseIf IsObj($oLeftPage) Then
		If Not ($oLeftPage.supportsService("com.sun.star.sheet.HeaderFooterContent")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPageStyle.LeftPageHeaderContent = $oLeftPage
	EndIf

	Return IsObj($oHeader) ? (SetError($__LO_STATUS_SUCCESS, 0, $oHeader)) : SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_PageStyleHeaderObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleHeaderShadow
; Description ...: Set or Retrieve the shadow settings for a Page Style Header.
; Syntax ........: _LOCalc_PageStyleHeaderShadow(ByRef $oPageStyle[, $iWidth = Null[, $iColor = Null[, $iLocation = Null]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Shadow Width of the Header, set in Hundredths of a Millimeter (HMM).
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Color of the Header shadow, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Header Shadow. See constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Headers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving ShadowFormat Object.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving ShadowFormat Object for Error Checking.
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
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleHeaderShadow(ByRef $oPageStyle, $iWidth = Null, $iColor = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tShdwFrmt
	Local $iError = 0
	Local $avShadow[3]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tShdwFrmt = $oPageStyle.HeaderShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.Location())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Or ($iWidth < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tShdwFrmt.ShadowWidth = $iWidth
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tShdwFrmt.Color = $iColor
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOC_SHADOW_NONE, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oPageStyle.HeaderShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oPageStyle.HeaderShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($tShdwFrmt.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iColor)) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iLocation)) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 4)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleHeaderShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleLayout
; Description ...: Modify or retrieve the Layout settings for a Page Style.
; Syntax ........: _LOCalc_PageStyleLayout(ByRef $oPageStyle[, $iLayout = Null[, $iNumFormat = Null[, $bTableAlignHori = Null[, $bTableAlignVert = Null[, $sPaperTray = Null]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iLayout             - [optional] an integer value (0-4). Default is Null. Specify the current Page layout style, either Left(Even) pages, Right(Odd) pages, or both Left(Even) and Right(Odd) pages or mirrored. See Constants, $LOC_PAGE_LAYOUT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The page numbering format to use for this Page Style. See Constants, $LOC_NUM_STYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bTableAlignHori     - [optional] a boolean value. Default is Null. If True, Centers the cells Horizontally on the printed page.
;                  $bTableAlignVert     - [optional] a boolean value. Default is Null. If True, Centers the cells Vertically on the printed page.
;                  $sPaperTray          - [optional] a string value. Default is Null. The paper source for your printer. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iLayout not an Integer, less than 0 or greater than 4. See Constants, $LOC_PAGE_LAYOUT_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iNumFormat not an Integer, less than 0 or greater than 71. See Constants, $LOC_NUM_STYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bTableAlignHori not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bTableAlignVert not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $sPaperTray not a string.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating Document Settings Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLayout
;                  |                               2 = Error setting $iNumFormat
;                  |                               4 = Error setting $bTableAlignHori
;                  |                               8 = Error setting $bTableAlignVert
;                  |                               16 = Error setting $sPaperTray
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I have no way to retrieve possible values for the Paper Tray parameter, at least that I can find. You may still use it if you know the appropriate value.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleLayout(ByRef $oPageStyle, $iLayout = Null, $iNumFormat = Null, $bTableAlignHori = Null, $bTableAlignVert = Null, $sPaperTray = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avLayout[5]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iLayout, $iNumFormat, $bTableAlignHori, $bTableAlignVert, $sPaperTray) Then
		__LO_ArrayFill($avLayout, $oPageStyle.PageStyleLayout(), $oPageStyle.NumberingType(), $oPageStyle.CenterHorizontally(), $oPageStyle.CenterVertically(), _
				$oPageStyle.PrinterPaperTray())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avLayout)
	EndIf

	If ($iLayout <> Null) Then
		If Not __LO_IntIsBetween($iLayout, $LOC_PAGE_LAYOUT_ALL, $LOC_PAGE_LAYOUT_MIRRORED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.PageStyleLayout = $iLayout
		$iError = ($oPageStyle.PageStyleLayout() = $iLayout) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LO_IntIsBetween($iNumFormat, $LOC_NUM_STYLE_CHARS_UPPER_LETTER, $LOC_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.NumberingType = $iNumFormat
		$iError = ($oPageStyle.NumberingType() = $iNumFormat) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bTableAlignHori <> Null) Then
		If Not IsBool($bTableAlignHori) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.CenterHorizontally = $bTableAlignHori
		$iError = ($oPageStyle.CenterHorizontally() = $bTableAlignHori) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bTableAlignVert <> Null) Then
		If Not IsBool($bTableAlignVert) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.CenterVertically = $bTableAlignVert
		$iError = ($oPageStyle.CenterVertically() = $bTableAlignVert) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sPaperTray <> Null) Then
		If Not IsString($sPaperTray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.PrinterPaperTray = $sPaperTray
		$iError = ($oPageStyle.PrinterPaperTray() = $sPaperTray) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleLayout

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleMargins
; Description ...: Modify or retrieve the margin settings for a Page Style.
; Syntax ........: _LOCalc_PageStyleMargins(ByRef $oPageStyle[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iLeft               - [optional] an integer value. Default is Null. The amount of space to leave between the left edge of the page and the document text. If you are using the Mirrored page layout, enter the amount of space to leave between the inner text margin and the inner edge of the page. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The amount of space to leave between the right edge of the page and the document text. If you are using the Mirrored page layout, enter the amount of space to leave between the outer text margin and the outer edge of the page. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The amount of space to leave between the upper edge of the page and the document text. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The amount of space to leave between the lower edge of the page and the document text. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iLeft not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iRight not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleMargins(ByRef $oPageStyle, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiMargins[4]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iLeft, $iRight, $iTop, $iBottom) Then
		__LO_ArrayFill($aiMargins, $oPageStyle.LeftMargin(), $oPageStyle.RightMargin(), $oPageStyle.TopMargin(), $oPageStyle.BottomMargin())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiMargins)
	EndIf

	If ($iLeft <> Null) Then
		If Not IsInt($iLeft) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.LeftMargin = $iLeft
		$iError = (__LO_IntIsBetween($oPageStyle.LeftMargin(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRight <> Null) Then
		If Not IsInt($iRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.RightMargin = $iRight
		$iError = (__LO_IntIsBetween($oPageStyle.RightMargin(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTop <> Null) Then
		If Not IsInt($iTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.TopMargin = $iTop
		$iError = (__LO_IntIsBetween($oPageStyle.TopMargin(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iBottom <> Null) Then
		If Not IsInt($iBottom) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.BottomMargin = $iBottom
		$iError = (__LO_IntIsBetween($oPageStyle.BottomMargin(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleOrganizer
; Description ...: Set or retrieve the Organizer settings of a Page Style.
; Syntax ........: _LOCalc_PageStyleOrganizer(ByRef $oDoc, $oPageStyle[, $sNewPageStyleName = Null[, $bHidden = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $sNewPageStyleName   - [optional] a string value. Default is Null. The new name to set the Page Style called in $oPageStyle to.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, the style is hidden in L.O. UI. Libre Office 4.0 and Up.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 4 Return 0 = $sNewPageStyleName not a String.
;                  @Error 1 @Extended 5 Return 0 = Page Style name called in $sNewPageStyleName already exists in document.
;                  @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sNewParStyleName
;                  |                               2 = Error setting $bHidden
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 1 or 2 Element Array with values in order of function parameters. If the Libre Office version is below 4.0, the Array will contain 1 element because $bHidden is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LOCalc_PageStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleOrganizer(ByRef $oDoc, ByRef $oPageStyle, $sNewPageStyleName = Null, $bHidden = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOrganizer[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($sNewPageStyleName, $bHidden) Then
		If __LO_VersionCheck(4.0) Then
			__LO_ArrayFill($avOrganizer, $oPageStyle.Name(), $oPageStyle.Hidden())

		Else
			__LO_ArrayFill($avOrganizer, $oPageStyle.Name())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOrganizer)
	EndIf

	If ($sNewPageStyleName <> Null) Then
		If Not IsString($sNewPageStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If _LOCalc_PageStyleExists($oDoc, $sNewPageStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.Name = $sNewPageStyleName
		$iError = ($oPageStyle.Name() = $sNewPageStyleName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not __LO_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oPageStyle.Hidden = $bHidden
		$iError = ($oPageStyle.Hidden() = $bHidden) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleOrganizer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStylePaperFormat
; Description ...: Modify or retrieve the paper format settings for a Page Style.
; Syntax ........: _LOCalc_PageStylePaperFormat(ByRef $oPageStyle[, $iWidth = Null[, $iHeight = Null[, $bLandscape = Null]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Width of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOC_PAPER_WIDTH_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iHeight             - [optional] an integer value. Default is Null. The Height of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOC_PAPER_HEIGHT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bLandscape          - [optional] a boolean value. Default is Null. If True, displays the page in Landscape layout.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $bLandscape not a Boolean value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iHeight
;                  |                               4 = Error setting $bLandscape
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStylePaperFormat(ByRef $oPageStyle, $iWidth = Null, $iHeight = Null, $bLandscape = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFormat[3]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iWidth, $iHeight, $bLandscape) Then
		__LO_ArrayFill($avFormat, $oPageStyle.Width(), $oPageStyle.Height(), $oPageStyle.IsLandscape())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFormat)
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.Width = $iWidth
		$iError = (__LO_IntIsBetween($oPageStyle.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iHeight <> Null) Then
		If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.Height = $iHeight
		$iError = (__LO_IntIsBetween($oPageStyle.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bLandscape <> Null) Then
		If Not IsBool($bLandscape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		If ($oPageStyle.IsLandscape() = $bLandscape) Then
			; If $bLandscape called setting is the same as the current setting, do nothing.

		Else
			; Retrieve current settings.
			$iHeight = $oPageStyle.Height()
			$iWidth = $oPageStyle.Width()

			; Switch Width with height, height with width.
			$oPageStyle.Height = $iWidth
			$oPageStyle.Width() = $iHeight
		EndIf

		$oPageStyle.IsLandscape = $bLandscape
		$iError = ($oPageStyle.IsLandscape() = $bLandscape) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStylePaperFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStylesGetNames
; Description ...: Retrieve an array of all Page Style names available for a document.
; Syntax ........: _LOCalc_PageStylesGetNames(ByRef $oDoc[, $bUserOnly = False[, $bAppliedOnly = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True only User-Created Page Styles are returned.
;                  $bAppliedOnly        - [optional] a boolean value. Default is False. If True only Applied Page Styles are returned.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bUserOnly not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bAppliedOnly not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Page Styles Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. An Array containing all Page Styles matching the input parameters. @Extended contains the count of results returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If Only a Document object is input, all available Page styles will be returned.
;                  Else if $bUserOnly is called with True, only User-Created Page Styles are returned.
;                  Else if $bAppliedOnly is called with True, only Applied Page Styles are returned.
;                  If Both are True then only User-Created Page styles that are applied are returned.
; Related .......: _LOCalc_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStylesGetNames(ByRef $oDoc, $bUserOnly = False, $bAppliedOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local $sExecute = ""
	Local $aStyles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	Local $oStyles = $oDoc.StyleFamilies.getByName("PageStyles")
	If Not IsObj($oStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $aStyles[$oStyles.getCount()]

	If Not $bUserOnly And Not $bAppliedOnly Then
		For $i = 0 To $oStyles.getCount() - 1
			$aStyles[$i] = $oStyles.getByIndex($i).DisplayName
			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, $i, $aStyles)
	EndIf

	$sExecute = ($bUserOnly) ? ("($oStyles.getByIndex($i).isUserDefined())") : ($sExecute)
	$sExecute = ($bUserOnly And $bAppliedOnly) ? ($sExecute & " And ") : ($sExecute)
	$sExecute = ($bAppliedOnly) ? ($sExecute & "($oStyles.getByIndex($i).isInUse())") : ($sExecute)

	For $i = 0 To $oStyles.getCount() - 1
		If Execute($sExecute) Then
			$aStyles[$iCount] = $oStyles.getByIndex($i).DisplayName()
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next
	ReDim $aStyles[$iCount]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aStyles)
EndFunc   ;==>_LOCalc_PageStylesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleShadow
; Description ...: Set or Retrieve the shadow settings for a Page Style.
; Syntax ........: _LOCalc_PageStyleShadow(ByRef $oPageStyle[, $iWidth = Null[, $iColor = Null[, $iLocation = Null]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Shadow Width of the Page, set in Hundredths of a Millimeter (HMM).
;                  $iColor              - [optional] an integer value. Default is Null (0-16777215). The shadow Color of the Page, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Page Shadow. See constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ShadowFormat Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving ShadowFormat Object for Error checking.
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
; Related .......: _LOCalc_PageStyleCreate, _LOCalc_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleShadow(ByRef $oPageStyle, $iWidth = Null, $iColor = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tShdwFrmt
	Local $iError = 0
	Local $avShadow[3]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tShdwFrmt = $oPageStyle.ShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.Location())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Or ($iWidth < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tShdwFrmt.ShadowWidth = $iWidth
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tShdwFrmt.Color = $iColor
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOC_SHADOW_NONE, $LOC_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oPageStyle.ShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oPageStyle.ShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($tShdwFrmt.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iColor)) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iLocation)) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 4)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleSheetPageOrder
; Description ...: Set or Retrieve Sheet Printing Page order settings.
; Syntax ........: _LOCalc_PageStyleSheetPageOrder(ByRef $oPageStyle[, $bTop2Bottom = Null[, $bFirstPageNum = Null[, $iFirstPage = Null]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $bTop2Bottom         - [optional] a boolean value. Default is Null. If True, the Sheet is printed from Top to Bottom and then Right. If False, the Sheet is printed Left to Right, and then down.
;                  $bFirstPageNum       - [optional] a boolean value. Default is Null. If True Page numbering will be restarted.
;                  $iFirstPage          - [optional] an integer value (0-9999). Default is Null. The Page number you want the numbering to restart at.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $bTop2Bottom not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bFirstPageNum not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iFirstPage not an Integer, less than 0 or greater than 9999.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bTop2Bottom
;                  |                               2 = Error setting $bFirstPageNum
;                  |                               4 = Error setting $iFirstPage
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleSheetScale, _LOCalc_PageStyleSheetPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleSheetPageOrder(ByRef $oPageStyle, $bTop2Bottom = Null, $bFirstPageNum = Null, $iFirstPage = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPageOrder[3]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bTop2Bottom, $bFirstPageNum, $iFirstPage) Then
		__LO_ArrayFill($avPageOrder, $oPageStyle.PrintDownFirst(), ($oPageStyle.FirstPageNumber() = 0) ? (False) : (True), $oPageStyle.FirstPageNumber()) ; When First Page Number is unchecked in L.O., FirstPageNumber is set to 0.

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPageOrder)
	EndIf

	If ($bTop2Bottom <> Null) Then
		If Not IsBool($bTop2Bottom) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.PrintDownFirst = $bTop2Bottom
		$iError = ($oPageStyle.PrintDownFirst() = $bTop2Bottom) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bFirstPageNum <> Null) Then
		If Not IsBool($bFirstPageNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.FirstPageNumber = ($bFirstPageNum) ? (1) : (0)
		$iError = ($oPageStyle.FirstPageNumber() = ($bFirstPageNum) ? (1) : (0)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iFirstPage <> Null) Then
		If Not __LO_IntIsBetween($iFirstPage, 0, 9999) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.FirstPageNumber = $iFirstPage
		$iError = ($oPageStyle.FirstPageNumber() = $iFirstPage) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleSheetPageOrder

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleSheetPrint
; Description ...: Set or Retrieve Sheet Printing settings.
; Syntax ........: _LOCalc_PageStyleSheetPrint(ByRef $oPageStyle[, $bHeaders = Null[, $bGrid = Null[, $bComments = Null[, $bObjectsOrImages = Null[, $bCharts = Null[, $bDrawing = Null[, $bFormulas = Null[, $bZeroValues = Null]]]]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $bHeaders            - [optional] a boolean value. Default is Null. If True, Column and Row Headers will be included when printed.
;                  $bGrid               - [optional] a boolean value. Default is Null. If True, Cell Border lines are printed as a grid.
;                  $bComments           - [optional] a boolean value. Default is Null. If True, Comments are printed.
;                  $bObjectsOrImages    - [optional] a boolean value. Default is Null. If True, Objects or Images are printed.
;                  $bCharts             - [optional] a boolean value. Default is Null. If True, Charts are printed.
;                  $bDrawing            - [optional] a boolean value. Default is Null. If True, Drawings are printed.
;                  $bFormulas           - [optional] a boolean value. Default is Null. If True, Formulas are printed instead of the results.
;                  $bZeroValues         - [optional] a boolean value. Default is Null. IF True, Cells containing Zero Values are printed.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $bHeaders not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bGrid not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bComments not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bObjectsOrImages not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bCharts not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bDrawing not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bFormulas not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bZeroValues not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bHeaders
;                  |                               2 = Error setting $bGrid
;                  |                               4 = Error setting $bComments
;                  |                               8 = Error setting $bObjectsOrImages
;                  |                               16 = Error setting $bCharts
;                  |                               32 = Error setting $bDrawing
;                  |                               64 = Error setting $bFormulas
;                  |                               128 = Error setting $bZeroValues
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_PageStyleSheetPageOrder, _LOCalc_PageStyleSheetScale
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleSheetPrint(ByRef $oPageStyle, $bHeaders = Null, $bGrid = Null, $bComments = Null, $bObjectsOrImages = Null, $bCharts = Null, $bDrawing = Null, $bFormulas = Null, $bZeroValues = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abSheetPrint[8]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bHeaders, $bGrid, $bComments, $bObjectsOrImages, $bCharts, $bDrawing, $bFormulas, $bZeroValues) Then
		__LO_ArrayFill($abSheetPrint, $oPageStyle.PrintHeaders(), $oPageStyle.PrintGrid(), $oPageStyle.PrintAnnotations(), $oPageStyle.PrintObjects(), _
				$oPageStyle.PrintCharts(), $oPageStyle.PrintDrawing(), $oPageStyle.PrintFormulas(), $oPageStyle.PrintZeroValues())

		Return SetError($__LO_STATUS_SUCCESS, 1, $abSheetPrint)
	EndIf

	If ($bHeaders <> Null) Then
		If Not IsBool($bHeaders) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.PrintHeaders = $bHeaders
		$iError = ($oPageStyle.PrintHeaders() = $bHeaders) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bGrid <> Null) Then
		If Not IsBool($bGrid) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.PrintGrid = $bGrid
		$iError = ($oPageStyle.PrintGrid() = $bGrid) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bComments <> Null) Then
		If Not IsBool($bComments) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.PrintAnnotations = $bComments
		$iError = ($oPageStyle.PrintAnnotations() = $bComments) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bObjectsOrImages <> Null) Then
		If Not IsBool($bObjectsOrImages) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.PrintObjects = $bObjectsOrImages
		$iError = ($oPageStyle.PrintObjects() = $bObjectsOrImages) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bCharts <> Null) Then
		If Not IsBool($bCharts) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.PrintCharts = $bCharts
		$iError = ($oPageStyle.PrintCharts() = $bCharts) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bDrawing <> Null) Then
		If Not IsBool($bDrawing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPageStyle.PrintDrawing = $bDrawing
		$iError = ($oPageStyle.PrintDrawing() = $bDrawing) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bFormulas <> Null) Then
		If Not IsBool($bFormulas) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oPageStyle.PrintFormulas = $bFormulas
		$iError = ($oPageStyle.PrintFormulas() = $bFormulas) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bZeroValues <> Null) Then
		If Not IsBool($bZeroValues) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oPageStyle.PrintZeroValues = $bZeroValues
		$iError = ($oPageStyle.PrintZeroValues() = $bZeroValues) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleSheetPrint

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PageStyleSheetScale
; Description ...: Set or Retrieve the Sheet Printing Scale settings.
; Syntax ........: _LOCalc_PageStyleSheetScale(ByRef $oPageStyle[, $iMode = Null[, $iVariable1 = Null[, $iVariable2 = Null]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $iMode               - [optional] an integer value (1-3). Default is Null. The Scaling mode when the spreadsheet is printed. See Constants $LOC_SCALE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iVariable1          - [optional] an integer value. Default is Null. The First Scale Value. See Remarks
;                  $iVariable2          - [optional] an integer value. Default is Null. The Second Scale Value. See Remarks
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iMode not an Integer, less than 1 or greater than 3. See Constants $LOC_SCALE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = Current mode set to $LOC_SCALE_REDUCE_ENLARGE, but $iVariable1 is not an Integer, less than 10 or greater than 400%.
;                  @Error 1 @Extended 5 Return 0 = Current mode set to $LOC_SCALE_FIT_WIDTH_HEIGHT, but $iVariable1 is not an Integer, less than 1 or greater than 1000.
;                  @Error 1 @Extended 6 Return 0 = Current mode set to $LOC_SCALE_FIT_WIDTH_HEIGHT, but $iVariable2 is not an Integer, less than 1 or greater than 1000.
;                  @Error 1 @Extended 7 Return 0 = Current mode set to $LOC_SCALE_FIT_PAGES, but $iVariable1 is not an Integer, less than 1 or greater than 1000.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to determine Scale Mode.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iMode
;                  |                               2 = Error setting $iVariable1 for mode $LOC_SCALE_REDUCE_ENLARGE
;                  |                               4 = Error setting $iVariable1 for mode $LOC_SCALE_FIT_WIDTH_HEIGHT
;                  |                               8 = Error setting $iVariable2 for mode $LOC_SCALE_FIT_WIDTH_HEIGHT
;                  |                               16 = Error setting $iVariable1 for mode $LOC_SCALE_FIT_PAGES
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 or 3 Element Array with values in order of function parameters. If the current mode is equal to $LOC_SCALE_FIT_WIDTH_HEIGHT, there will be three elements, Elelemnt 1(0) will be the current mode, element 2(1) will be the Width value, and Element 3(2) will be the height value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $iVariable1 and $iVariable2 setting values depend on what the current or new mode is. For Modes Reduce/Enlarge and Fit Pages, $iVariable1 ALONE is used for setting Scale Percentage or Number of Pages Respectively.
;                  If Mode is set to Fit Width and Height, $iVariable1 is for setting Width, and $iVariable2 is for setting height.
;                  You can set the Scale Values still without setting Mode each time.
;                  When Scale Mode is set to $LOC_SCALE_REDUCE_ENLARGE, the Minimum scaling value for $iVariable1 is 10%, and the Maximum is 400%.
;                  When Scale Mode is set to $LOC_SCALE_FIT_WIDTH_HEIGHT, the Minimum scaling value for both Vairable1 and Variable2 is 1, and the Maximum is 1000.
;                  When Scale Mode is set to $LOC_SCALE_FIT_PAGES, the Minimum scaling value for $iVariable1 is 1, and the Maximum is 1000.
; Related .......: _LOCalc_PageStyleSheetPageOrder, _LOCalc_PageStyleSheetPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_PageStyleSheetScale(ByRef $oPageStyle, $iMode = Null, $iVariable1 = Null, $iVariable2 = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abSheetScale[2]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iMode, $iVariable1, $iVariable2) Then
		If ($oPageStyle.ScaleToPagesX() > 0) Or ($oPageStyle.ScaleToPagesY() > 0) Then ; Determine which Scale mode is active
			__LO_ArrayFill($abSheetScale, $LOC_SCALE_FIT_WIDTH_HEIGHT, $oPageStyle.ScaleToPagesX(), $oPageStyle.ScaleToPagesY)

		ElseIf ($oPageStyle.ScaleToPages() > 0) Then
			__LO_ArrayFill($abSheetScale, $LOC_SCALE_FIT_PAGES, $oPageStyle.ScaleToPages())

		ElseIf ($oPageStyle.PageScale() > 0) Then ; Page Scale has to be last because each time I Set one of the other settings, Scale returns to 100%, if I set it back to 0 I lose my other settings. So Scale seems to be alway above 0.
			__LO_ArrayFill($abSheetScale, $LOC_SCALE_REDUCE_ENLARGE, $oPageStyle.PageScale())

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Failed to determine Scale Mode
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $abSheetScale)
	EndIf

	If ($iMode <> Null) Then
		If Not __LO_IntIsBetween($iMode, $LOC_SCALE_REDUCE_ENLARGE, $LOC_SCALE_FIT_PAGES) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		Switch $iMode
			Case $LOC_SCALE_REDUCE_ENLARGE
				If ($oPageStyle.PageScale() = 0) Then ; Only change settings if the mode isn't already set to this mode.
					$oPageStyle.ScaleToPages = 0 ; There is no Mode setting, Either one setting or another is set to higher than 0, so set all others to 0.
					$oPageStyle.ScaleToPagesX = 0
					$oPageStyle.ScaleToPagesY = 0
					$oPageStyle.PageScale = 10
				EndIf
				$iError = ($oPageStyle.PageScale() = 1) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_SCALE_FIT_WIDTH_HEIGHT
				If ($oPageStyle.ScaleToPagesX() = 0) Or ($oPageStyle.ScaleToPagesY() = 0) Then
					$oPageStyle.PageScale = 0
					$oPageStyle.ScaleToPages = 0
					$oPageStyle.ScaleToPagesX = 1
					$oPageStyle.ScaleToPagesY = 1
				EndIf
				$iError = (($oPageStyle.ScaleToPagesX() = 1) And ($oPageStyle.ScaleToPagesY = 1)) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_SCALE_FIT_PAGES
				If ($oPageStyle.ScaleToPages() = 0) Then
					$oPageStyle.PageScale = 0
					$oPageStyle.ScaleToPagesX = 0
					$oPageStyle.ScaleToPagesY = 0
					$oPageStyle.ScaleToPages = 1
				EndIf
				$iError = ($oPageStyle.ScaleToPages = 1) ? ($iError) : (BitOR($iError, 1))
		EndSwitch
	EndIf

	If ($iVariable1 <> Null) Or ($iVariable2 <> Null) Then
		If ($oPageStyle.ScaleToPagesX() > 0) Or ($oPageStyle.ScaleToPagesY() > 0) Then ; Determine which Scale mode is active
			If ($iVariable1 <> Null) Then
				If Not __LO_IntIsBetween($iVariable1, 1, 1000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

				$oPageStyle.ScaleToPagesX = $iVariable1
				$iError = ($oPageStyle.ScaleToPagesX = $iVariable1) ? ($iError) : (BitOR($iError, 4))
			EndIf

			If ($iVariable2 <> Null) Then
				If Not __LO_IntIsBetween($iVariable2, 1, 1000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

				$oPageStyle.ScaleToPagesY = $iVariable2
				$iError = ($oPageStyle.ScaleToPagesY = $iVariable2) ? ($iError) : (BitOR($iError, 8))
			EndIf

		ElseIf ($oPageStyle.ScaleToPages() > 0) Then
			If Not __LO_IntIsBetween($iVariable1, 1, 1000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oPageStyle.ScaleToPages = $iVariable1
			$iError = ($oPageStyle.ScaleToPages = $iVariable1) ? ($iError) : (BitOR($iError, 16))

		ElseIf ($oPageStyle.PageScale() > 0) Then ; Page Scale has to be last because each time I Set one of the other settings, Scale returns to 100%, if I set it back to 0 I lose my other settings.
			If Not __LO_IntIsBetween($iVariable1, 10, 400) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oPageStyle.PageScale = $iVariable1
			$iError = ($oPageStyle.PageScale = $iVariable1) ? ($iError) : (BitOR($iError, 2))
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_PageStyleSheetScale
