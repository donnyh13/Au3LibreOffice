#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Impress
#include "LibreOfficeImpress_Constants.au3"
#include "LibreOfficeImpress_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Various functions for internal data processing, data retrieval, retrieving and applying settings for LibreOffice Impress.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LOImpress_CharEffect
; __LOImpress_CharFont
; __LOImpress_CharFontColor
; __LOImpress_CharOverLine
; __LOImpress_CharPosition
; __LOImpress_CharScaling
; __LOImpress_CharSpacing
; __LOImpress_CharStrikeOut
; __LOImpress_CharUnderLine
; __LOImpress_ColorRemoveAlpha
; __LOImpress_CreatePoint
; __LOImpress_CursorParHasTabStop
; __LOImpress_DimensionSettings
; __LOImpress_DrawShape_CreateArrow
; __LOImpress_DrawShape_CreateBasic
; __LOImpress_DrawShape_CreateCallout
; __LOImpress_DrawShape_CreateFlowchart
; __LOImpress_DrawShape_CreateLine
; __LOImpress_DrawShape_CreateStars
; __LOImpress_DrawShape_CreateSymbol
; __LOImpress_DrawShape_GetCustomType
; __LOImpress_DrawShapePointGetSettings
; __LOImpress_DrawShapePointModify
; __LOImpress_FilterNameGet
; __LOImpress_GetShapeName
; __LOImpress_GradientIsModified
; __LOImpress_GradientNameInsert
; __LOImpress_GradientPresets
; __LOImpress_InternalComErrorHandler
; __LOImpress_NumRuleCreateMap
; __LOImpress_ParAlignment
; __LOImpress_ParIndent
; __LOImpress_ParSpacing
; __LOImpress_ParTabStopCreate
; __LOImpress_ParTabStopDelete
; __LOImpress_ParTabStopMod
; __LOImpress_ParTabStopsGetList
; __LOImpress_ShapeAreaColor
; __LOImpress_ShapeAreaGradientMulticolor
; __LOImpress_ShapeAreaShadow
; __LOImpress_ShapeAreaShadowModify
; __LOImpress_ShapeAreaTransparency
; __LOImpress_ShapeAreaTransparencyGradientMulti
; __LOImpress_ShapeGetType
; __LOImpress_ShapeLineArrowheadNameInsert
; __LOImpress_ShapeLineArrowStyleName
; __LOImpress_ShapeLineDashNameInsert
; __LOImpress_ShapeLineStyleName
; __LOImpress_ShapePresStyleNumCreateScript
; __LOImpress_ShapePresStyleNumDeleteScript
; __LOImpress_ShapePresStyleNumInitiateDocument
; __LOImpress_ShapePresStyleNumModify
; __LOImpress_ShapeStyleAreaGradient
; __LOImpress_ShapeStyleAreaTransparencyGradient
; __LOImpress_ShapeStyleCompare
; __LOImpress_ShapeStyleLineArrowStyles
; __LOImpress_ShapeStyleLineProperties
; __LOImpress_ShapeTextAttrAnimation
; __LOImpress_ShapeTextAttrFit
; __LOImpress_ShapeTextAttrSettings
; __LOImpress_StyleCharFontColor
; __LOImpress_Transition
; __LOImpress_TransparencyGradientConvert
; __LOImpress_TransparencyGradientNameInsert
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CharEffect
; Description ...: Set or Retrieve the Font Effect settings.
; Syntax ........: __LOImpress_CharEffect(ByRef $oObj[, $iCase = Null[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iCase               - [optional] an integer value (0-4). Default is Null. The Character Case Style. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRelief             - [optional] an integer value (0-2). Default is Null. The Character Relief style. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOutline            - [optional] a boolean value. Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] a boolean value. Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iCase not an Integer, less than 0 or greater than 4. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $bOutline not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bShadow not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iCase
;                  |                               2 = Error setting $iRelief
;                  |                               4 = Error setting $bOutline
;                  |                               8 = Error setting $bShadow
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CharEffect(ByRef $oObj, $iCase = Null, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avEffect[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iCase, $iRelief, $bOutline, $bShadow) Then
		__LO_ArrayFill($avEffect, $oObj.CharCaseMap(), $oObj.CharRelief(), $oObj.CharContoured(), $oObj.CharShadowed())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avEffect)
	EndIf

	If ($iCase <> Null) Then
		If Not __LO_IntIsBetween($iCase, $LOI_CHAR_CASEMAP_NONE, $LOI_CHAR_CASEMAP_SM_CAPS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.CharCaseMap = $iCase
		$iError = ($oObj.CharCaseMap() = $iCase) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRelief <> Null) Then
		If Not __LO_IntIsBetween($iRelief, $LOI_CHAR_RELIEF_NONE, $LOI_CHAR_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.CharRelief = $iRelief
		$iError = ($oObj.CharRelief() = $iRelief) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bOutline <> Null) Then
		If Not IsBool($bOutline) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharContoured = $bOutline
		$iError = ($oObj.CharContoured() = $bOutline) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bShadow <> Null) Then
		If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharShadowed = $bShadow
		$iError = ($oObj.CharShadowed() = $bShadow) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_CharEffect

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CharFont
; Description ...: Set and Retrieve the Font Settings.
; Syntax ........: __LOImpress_CharFont(ByRef $oObj[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $sFontName           - [optional] a string value. Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] a general number value. Default is Null. The new Font size.
;                  $iPosture            - [optional] an integer value (0-5). Default is Null. The Font Italic setting. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] an integer value (0, 50-200). Default is Null. The Font Bold settings see Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 3 Return 0 = Font called in $sFontName not available.
;                  @Error 1 @Extended 4 Return 0 = $nFontSize not a number.
;                  @Error 1 @Extended 5 Return 0 = $iPosture not an Integer, less than 0 or greater than 5. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iWeight not an Integer, less than 50 but not equal to 0, or greater than 200. See Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $nFontSize
;                  |                               4 = Error setting $iPosture
;                  |                               8 = Error setting $iWeight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;                  LibreOffice accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CharFont(ByRef $oObj, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFont[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sFontName, $nFontSize, $iPosture, $iWeight) Then
		__LO_ArrayFill($avFont, $oObj.CharFontName(), $oObj.CharHeight(), $oObj.CharPosture(), $oObj.CharWeight())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFont)
	EndIf

	If ($sFontName <> Null) Then
		If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		If Not _LOImpress_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.CharFontName = $sFontName
		$iError = ($oObj.CharFontName() = $sFontName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($nFontSize <> Null) Then
		If Not IsNumber($nFontSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharHeight = $nFontSize
		$iError = ($oObj.CharHeight() = $nFontSize) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iPosture <> Null) Then
		If Not __LO_IntIsBetween($iPosture, $LOI_CHAR_POSTURE_NONE, $LOI_CHAR_POSTURE_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharPosture = $iPosture
		$iError = ($oObj.CharPosture() = $iPosture) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iWeight <> Null) Then
		If Not __LO_IntIsBetween($iWeight, $LOI_CHAR_WEIGHT_THIN, $LOI_CHAR_WEIGHT_BLACK, "", $LOI_CHAR_WEIGHT_DONT_KNOW) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharWeight = $iWeight
		$iError = ($oObj.CharWeight() = $iWeight) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_CharFont

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CharFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting values.
; Syntax ........: __LOImpress_CharFontColor(ByRef $oObj[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor or Shape object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert or _LOImpress_ShapesGetList function.
;                  $iFontColor          - [optional] an integer value (-1-16777215). Default is Null. The font Color value, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. Transparency percentage. 0 is visible, 100 is invisible. Available for LibreOffice 7.0 and up.
;                  $iHighlight          - [optional] an integer value (-1-16777215). Default is Null. The highlight Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 3 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100%.
;                  @Error 1 @Extended 4 Return 0 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $FontColor
;                  |                               2 = Error setting $iTransparency.
;                  |                               4 = Error setting $iHighlight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current LibreOffice version lower than 7.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 or 3 Element Array with values in order of function parameters. If The current LibreOffice version is below 7.0 the returned array will contain 2 elements, because $iTransparency is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CharFontColor(ByRef $oObj, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency
	Local $avColor[2]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iFontColor, $iTransparency, $iHighlight) Then
		If __LO_VersionCheck(7.0) Then
			__LO_ArrayFill($avColor, __LOImpress_ColorRemoveAlpha($oObj.CharColor()), $oObj.CharTransparence(), $oObj.CharBackColor())

		Else
			__LO_ArrayFill($avColor, __LOImpress_ColorRemoveAlpha($oObj.CharColor()), $oObj.CharBackColor())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iFontColor <> Null) Then
		If Not __LO_IntIsBetween($iFontColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If __LO_VersionCheck(7.0) Then
			$iOldTransparency = $oObj.CharTransparence()
			If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		EndIf

		$oObj.CharColor = $iFontColor
		$iError = ($oObj.CharColor() = $iFontColor) ? ($iError) : (BitOR($iError, 1))

		If IsInt($iOldTransparency) Then $oObj.CharTransparence = $iOldTransparency
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If Not __LO_VersionCheck(7.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oObj.CharTransparence = $iTransparency
		$iError = ($oObj.CharTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iHighlight <> Null) Then
		If Not __LO_IntIsBetween($iHighlight, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		; CharHighlight; same as CharBackColor---Libre seems to use back color for highlighting however, so using that for setting.
;~ 		If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 2, 0)
;~ 		$oObj.CharHighlight = $iHighlight ;-- keeping old method in case.
;~ 		$iError = ($oObj.CharHighlight() = $iHighlight) ? ($iError) : (BitOR($iError, 4)
		$oObj.CharBackColor = $iHighlight
		$iError = ($oObj.CharBackColor() = $iHighlight) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_CharFontColor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CharOverLine
; Description ...: Set and retrieve the OverLine settings.
; Syntax ........: __LOImpress_CharOverLine(ByRef $oObj[, $iOverLineStyle = Null[, $iOLColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iOverLineStyle      - [optional] an integer value (0-18). Default is Null. The style of the Overline line, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iOLColor            - [optional] an integer value (-1-16777215). Default is Null. The Overline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If True, white spaces are not Overlined.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iOverLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  @Error 1 @Extended 3 Return 0 = $iOLColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iOverLineStyle
;                  |                               2 = Error setting $iOLColor
;                  |                               4 = Error setting $bWordOnly
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Overline line style uses the same constants as underline style.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CharOverLine(ByRef $oObj, $iOverLineStyle = Null, $iOLColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOverLine[3]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iOverLineStyle, $iOLColor, $bWordOnly) Then
		__LO_ArrayFill($avOverLine, $oObj.CharOverline(), $oObj.CharOverlineColor(), $oObj.CharWordMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOverLine)
	EndIf

	If ($iOverLineStyle <> Null) Then
		If Not __LO_IntIsBetween($iOverLineStyle, $LOI_CHAR_UNDERLINE_NONE, $LOI_CHAR_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.CharOverline = $iOverLineStyle
		$iError = ($oObj.CharOverline() = $iOverLineStyle) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iOLColor <> Null) Then
		If Not __LO_IntIsBetween($iOLColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If ($iOLColor = $LO_COLOR_OFF) Then
			If ($oObj.CharOverlineHasColor() = True) Then $oObj.CharOverlineHasColor = False
			$oObj.CharOverlineColor = $iOLColor

		Else
			If ($oObj.CharOverlineHasColor() = False) Then $oObj.CharOverlineHasColor = True
			$oObj.CharOverlineColor = $iOLColor
		EndIf

		$oObj.CharOverlineColor = $iOLColor
		$iError = ($oObj.CharOverlineColor() = $iOLColor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bWordOnly <> Null) Then
		If Not IsBool($bWordOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharWordMode = $bWordOnly
		$iError = ($oObj.CharWordMode() = $bWordOnly) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_CharOverLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size.
; Syntax ........: __LOImpress_CharPosition(ByRef $oObj[, $iSuperScript = Null[, $iSubScript = Null[, $iRelativeSize = Null]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor or Shape object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert or _LOImpress_ShapesGetList function.
;                  $iSuperScript        - [optional] an integer value (-1-100). Default is Null. The Superscript percentage value. Call with -1 for Automatic SuperScript. See Remarks.
;                  $iSubScript          - [optional] an integer value (-1-100). Default is Null. Subscript percentage value. Call with -1 for Automatic SubScript. See Remarks.
;                  $iRelativeSize       - [optional] an integer value (1-100). Default is Null. The size percentage relative to current font size.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj does not support Character properties.
;                  @Error 1 @Extended 3 Return 0 = $iSuperScript not an Integer, less than 0 or greater than 100, but not 14000.
;                  @Error 1 @Extended 4 Return 0 = $iSubScript not an Integer, less than -100 or greater than 100, but not 14000 or -14000.
;                  @Error 1 @Extended 5 Return 0 = $iRelativeSize not an Integer, less than 1 or greater than 100.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iSuperScript
;                  |                               2 = Error setting $iSubScript
;                  |                               4 = Error setting $iRelativeSize.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
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
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CharPosition(ByRef $oObj, $iSuperScript = Null, $iSubScript = Null, $iRelativeSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPosition[3]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iSuperScript, $iSubScript, $iRelativeSize) Then
		__LO_ArrayFill($avPosition, ($oObj.CharEscapement() <= 0) ? (0) : ((__LO_IntIsBetween($oObj.CharEscapement(), 1, 100)) ? ($oObj.CharEscapement()) : (-1)), _
				($oObj.CharEscapement() >= 0) ? (0) : ((__LO_IntIsBetween($oObj.CharEscapement(), -1, -100)) ? (($oObj.CharEscapement() * -1)) : (-1)), _
				$oObj.CharEscapementHeight())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosition)
	EndIf

	If ($iSuperScript <> Null) Then
		If Not __LO_IntIsBetween($iSuperScript, -1, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If ($iSuperScript = -1) Then
			$oObj.CharEscapement = 14000
			$iError = ($oObj.CharEscapement() = 14000) ? ($iError) : (BitOR($iError, 1))

		Else
			$oObj.CharEscapement = $iSuperScript
			$iError = ($oObj.CharEscapement() = $iSuperScript) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($iSubScript <> Null) Then
		If Not __LO_IntIsBetween($iSubScript, -1, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If ($iSubScript = -1) Then
			$oObj.CharEscapement = -14000
			$iError = ($oObj.CharEscapement() = -14000) ? ($iError) : (BitOR($iError, 2))

		Else
			$iSubScript = ($iSubScript * -1) ; Change to negative value, as SubScript is set in negative integers.
			$oObj.CharEscapement = $iSubScript
			$iError = ($oObj.CharEscapement() = $iSubScript) ? ($iError) : (BitOR($iError, 2))
		EndIf
	EndIf

	If ($iRelativeSize <> Null) Then
		If Not __LO_IntIsBetween($iRelativeSize, 1, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharEscapementHeight = $iRelativeSize
		$iError = ($oObj.CharEscapementHeight() = $iRelativeSize) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_CharPosition

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CharScaling
; Description ...: Set or retrieve the character Scale settings.
; Syntax ........: __LOImpress_CharScaling(ByRef $oObj[, $iScaleWidth = Null])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor or Shape object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert or _LOImpress_ShapesGetList function.
;                  $iScaleWidth         - [optional] an integer value (1-100). Default is Null. The percentage to horizontally stretch or compress the text. 100 is normal sizing.
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iScaleWidth not an Integer or less than 1% or greater than 100%.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Scale width.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iScaleWidth
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current Scale Width value as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Fit to line seems to be unavailable in the API, and does not seem to work in LibreOffice anyway.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CharScaling(ByRef $oObj, $iScaleWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iCurrScale

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iScaleWidth) Then
		$iCurrScale = $oObj.CharScaleWidth()
		If Not IsInt($iCurrScale) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurrScale)
	EndIf

	If Not __LO_IntIsBetween($iScaleWidth, 1, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)     ; can't be less than 1%

	$oObj.CharScaleWidth = $iScaleWidth
	$iError = ($oObj.CharScaleWidth() = $iScaleWidth) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_CharScaling

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning).
; Syntax ........: __LOImpress_CharSpacing(ByRef $oObj[, $bAutoKerning = Null[, $nKerning = Null]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor or Shape object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert or _LOImpress_ShapesGetList function.
;                  $bAutoKerning        - [optional] a boolean value. Default is Null. If True, applies a spacing in between certain pairs of characters.
;                  $nKerning            - [optional] a general number value (-928.8-928.8). Default is Null. The kerning value of the characters. See Remarks. Values are in Printer's Points as set in the LibreOffice UI.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bAutoKerning not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $nKerning not a number, less than -928.8 or greater than 928.8 Points.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoKerning
;                  |                               2 = Error setting $nKerning.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User Display, however the internal setting is measured in Hundredths of a Millimeter (HMM). They will be automatically converted from Points to Hundredths of a Millimeter and back for retrieval of settings.
;                  The acceptable values are from -2 Pt to 928.8 Pt. The values can be directly converted easily, however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative Hundredths of a Millimeter internally from 928.9 up to 1000 Pt (Max setting).
;                  For example, 928.8Pt is the last correct value, which equals 32766 Hundredths of a Millimeter (HMM), after this LibreOffice reports the following: 928.9 Pt = -32766 HMM; 929 Pt = -32763 HMM; 929.1 = -32759; 1000 pt = -30258. Attempting to set Libre's kerning value to anything over 32768 HMM causes a COM exception, and attempting to set the kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the max settable kerning is -2.0 Pt to 928.8 Pt.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CharSpacing(ByRef $oObj, $bAutoKerning = Null, $nKerning = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avKerning[2]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bAutoKerning, $nKerning) Then
		$nKerning = _LO_UnitConvert($oObj.CharKerning(), $LO_CONVERT_UNIT_HMM_PT)
		__LO_ArrayFill($avKerning, $oObj.CharAutoKerning(), (($nKerning > 928.8) ? (1000) : (($nKerning < -928.8) ? (-1000) : ($nKerning))))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avKerning)
	EndIf

	If ($bAutoKerning <> Null) Then
		If Not IsBool($bAutoKerning) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.CharAutoKerning = $bAutoKerning
		$iError = ($oObj.CharAutoKerning() = $bAutoKerning) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($nKerning <> Null) Then
		If Not __LO_NumIsBetween($nKerning, -928.8, 928.8) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$nKerning = _LO_UnitConvert($nKerning, $LO_CONVERT_UNIT_PT_HMM)
		$oObj.CharKerning = $nKerning
		$iError = ($oObj.CharKerning() = $nKerning) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_CharSpacing

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CharStrikeOut
; Description ...: Set or Retrieve the Strikeout settings.
; Syntax ........: __LOImpress_CharStrikeOut(ByRef $oObj[, $iStrikeLineStyle = Null[, $bWordOnly = Null]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iStrikeLineStyle    - [optional] an integer value (0-6). Default is Null. The Strikeout Line Style, see constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If True, strike out is applied to words only, skipping whitespaces.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iStrikeLineStyle not an Integer, less than 0 or greater than 6. See constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $bWordOnly not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iStrikeLineStyle
;                  |                               2 = Error setting $bWordOnly
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CharStrikeOut(ByRef $oObj, $iStrikeLineStyle = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avStrikeOut[2]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iStrikeLineStyle, $bWordOnly) Then
		__LO_ArrayFill($avStrikeOut, $oObj.CharStrikeout(), $oObj.CharWordMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avStrikeOut)
	EndIf

	If ($iStrikeLineStyle <> Null) Then
		If Not __LO_IntIsBetween($iStrikeLineStyle, $LOI_CHAR_STRIKEOUT_NONE, $LOI_CHAR_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.CharStrikeout = $iStrikeLineStyle
		$iError = ($oObj.CharStrikeout() = $iStrikeLineStyle) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bWordOnly <> Null) Then
		If Not IsBool($bWordOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.CharWordMode = $bWordOnly
		$iError = ($oObj.CharWordMode() = $bWordOnly) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_CharStrikeOut

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CharUnderLine
; Description ...: Set and retrieve the Underline settings.
; Syntax ........: __LOImpress_CharUnderLine(ByRef $oObj[, $iUnderLineStyle = Null[, $iULColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iUnderLineStyle     - [optional] an integer value (0-18). Default is Null. The Underline line style, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iULColor            - [optional] an integer value (-1-16777215). Default is Null. The underline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If True, white spaces are not underlined.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj an Object.
;                  @Error 1 @Extended 2 Return 0 = $iUnderLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iULColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iUnderLineStyle
;                  |                               2 = Error setting $iULColor
;                  |                               4 = Error setting $bWordOnly
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CharUnderLine(ByRef $oObj, $iUnderLineStyle = Null, $iULColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avUnderLine[3]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iUnderLineStyle, $iULColor, $bWordOnly) Then
		__LO_ArrayFill($avUnderLine, $oObj.CharUnderline(), $oObj.CharUnderlineColor(), $oObj.CharWordMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avUnderLine)
	EndIf

	If ($iUnderLineStyle <> Null) Then
		If Not __LO_IntIsBetween($iUnderLineStyle, $LOI_CHAR_UNDERLINE_NONE, $LOI_CHAR_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.CharUnderline = $iUnderLineStyle
		$iError = ($oObj.CharUnderline() = $iUnderLineStyle) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iULColor <> Null) Then
		If Not __LO_IntIsBetween($iULColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If ($iULColor = $LO_COLOR_OFF) Then
			If ($oObj.CharUnderlineHasColor() = True) Then $oObj.CharUnderlineHasColor = False
			$oObj.CharUnderlineColor = $iULColor

		Else
			If ($oObj.CharUnderlineHasColor() = False) Then $oObj.CharUnderlineHasColor = True
			$oObj.CharUnderlineColor = $iULColor
		EndIf

		$iError = ($oObj.CharUnderlineColor() = $iULColor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bWordOnly <> Null) Then
		If Not IsBool($bWordOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharWordMode = $bWordOnly
		$iError = ($oObj.CharWordMode() = $bWordOnly) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_CharUnderLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ColorRemoveAlpha
; Description ...: Remove the Alpha value from a RGB Color Integer.
; Syntax ........: __LOImpress_ColorRemoveAlpha($iColor)
; Parameters ....: $iColor              - an integer value. A RGB Color Integer to remove Alpha from.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return $iColor = $iColor not an Integer. Returning $iColor to be sure not to lose the value.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Color already has no Alpha value, returning same color.
;                  @Error 0 @Extended 1 Return Integer = Success. Removed Alpha value from RGB Color Integer, returning new Color value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In functions which return the current color value, generally background colors, if Transparency (alpha) is set, the background color value is not the literal color set, but also includes the transparency value added to it. This functions removes that value for simpler color values.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ColorRemoveAlpha($iColor)
	Local $iRed, $iGreen, $iBlue, $iLong

	If Not IsInt($iColor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, $iColor)

	If __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_SUCCESS, 0, $iColor) ; If Color value is not greater than White(16777215) or less than -1, then there is no alpha to remove.

	; Obtain individual color values.
	$iRed = BitAND(BitShift($iColor, 16), 0xff)
	$iGreen = BitAND(BitShift($iColor, 8), 0xff)
	$iBlue = BitAND($iColor, 0xff)
	$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

	Return SetError($__LO_STATUS_SUCCESS, 1, $iLong)
EndFunc   ;==>__LOImpress_ColorRemoveAlpha

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CreatePoint
; Description ...: Creates a Position structure.
; Syntax ........: __LOImpress_CreatePoint($iX, $iY)
; Parameters ....: $iX                  - an integer value. The X position, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y position, in Hundredths of a Millimeter (HMM).
; Return values .: Success: Structure
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 2 Return 0 = $iY not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a Position Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Structure = Success. Returning created Position Structure using $iX and $iY values.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Modified from A. Pitonyak, Listing 493. in OOME 3.0
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CreatePoint($iX, $iY)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tPoint

	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tPoint = __LO_CreateStruct("com.sun.star.awt.Point")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tPoint.X = $iX
	$tPoint.Y = $iY

	Return SetError($__LO_STATUS_SUCCESS, 0, $tPoint)
EndFunc   ;==>__LOImpress_CreatePoint

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CursorParHasTabStop
; Description ...: Check whether a Paragraph has a requested TabStop.
; Syntax ........: __LOImpress_CursorParHasTabStop(ByRef $oTextCursor, $iTabStop)
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iTabStop            - an integer value. The Tab Stop to look for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTabStop not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve ParaTabStops Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = True if Paragraph has the requested TabStop. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CursorParHasTabStop(ByRef $oTextCursor, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atTabStops

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$atTabStops = $oTextCursor.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($atTabStops) - 1
		If ($atTabStops[$i].Position() = $iTabStop) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>__LOImpress_CursorParHasTabStop

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DimensionSettings
; Description ...: Set or Retrieve Dimension line settings.
; Syntax ........: __LOImpress_DimensionSettings(ByRef $oObj[, $iDistance = Null[, $iGuideOverhang = Null[, $iGuideDistance = Null[, $iLGuide = Null[, $iRGuide = Null[, $bBelow = Null[, $iDecimal = Null[, $iVertPos = Null[, $iHoriPos = Null[, $bParallel = Null[, $iUnitType = Null]]]]]]]]]]])
; Parameters ....: $oObj                - [in/out] an object. A Dimension Shape or Shape Style object returned by a previous _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iDistance           - [optional] an integer value (-10,008-10,008). Default is Null. The distance between the dimension line and the baseline, in Hundredths of a Millimeter (HMM).
;                  $iGuideOverhang      - [optional] an integer value (-10,008-10,008). Default is Null. The length of the left and right guides starting at the baseline. Positive values extend the guides above the baseline and negative values extend the guides below the baseline, in Hundredths of a Millimeter (HMM).
;                  $iGuideDistance      - [optional] an integer value (-10,008-10,008). Default is Null. The length of the right and left guides starting at the dimension line. Positive values extend the guides above the dimension line and negative values extend the guides below the dimension line, in Hundredths of a Millimeter (HMM).
;                  $iLGuide             - [optional] an integer value (-10,008-10,008). Default is Null. The length of the left guide starting at the dimension line. Positive values extend the guide below the dimension line and negative values extend the guide above the dimension line, in Hundredths of a Millimeter (HMM).
;                  $iRGuide             - [optional] an integer value (-10,008-10,008). Default is Null. The length of the right guide starting at the dimension line. Positive values extend the guide below the dimension line and negative values extend the guide above the dimension line, in Hundredths of a Millimeter (HMM).
;                  $bBelow              - [optional] a boolean value. Default is Null. If True, the properties set in the Line area are Reversed.
;                  $iDecimal            - [optional] an integer value (0-99). Default is Null. The number of decimal places.
;                  $iVertPos            - [optional] an integer value (0-4). Default is Null. The position of the dimension line in reference to the text vertically. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iHoriPos            - [optional] an integer value (0-3). Default is Null. The position of the dimension text horizontally. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bParallel           - [optional] a boolean value. Default is Null. If True, Displays the text parallel to or at 90 degrees to the dimension line.
;                  $iUnitType           - [optional] an integer value (-1-15). Default is Null. The type of measurement units, if any, to display. See Constants, $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iDistance not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 3 Return 0 = $iGuideOverhang not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 4 Return 0 = $iGuideDistance not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 5 Return 0 = $iLGuide not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 6 Return 0 = $iRGuide not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 7 Return 0 = $bBelow not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iDecimal not an Integer, less than 0 or greater than 99.
;                  @Error 1 @Extended 9 Return 0 = $iVertPos not an Integer, less than 0 or greater than 4. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $iHoriPos not an Integer, less than 0 or greater than 3. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $bParallel not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iUnitType not an Integer, less than -1 or greater than 15. See Constants, $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iDistance
;                  |                               2 = Error setting $iGuideOverhang
;                  |                               4 = Error setting $iGuideDistance
;                  |                               8 = Error setting $iLGuide
;                  |                               16 = Error setting $iRGuide
;                  |                               32 = Error setting $bBelow
;                  |                               64 = Error setting $iDecimal
;                  |                               128 = Error setting $iVertPos
;                  |                               256 = Error setting $iHoriPos
;                  |                               512 = Error setting $bParallel
;                  |                               1024 = Error setting $iUnitType
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DimensionSettings(ByRef $oObj, $iDistance = Null, $iGuideOverhang = Null, $iGuideDistance = Null, $iLGuide = Null, $iRGuide = Null, $bBelow = Null, $iDecimal = Null, $iVertPos = Null, $iHoriPos = Null, $bParallel = Null, $iUnitType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDimension[11]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iDistance, $iGuideOverhang, $iGuideDistance, $iLGuide, $iRGuide, $bBelow, $iDecimal, $iVertPos, $iHoriPos, $bParallel, $iUnitType) Then
		__LO_ArrayFill($avDimension, $oObj.MeasureLineDistance(), $oObj.MeasureHelpLineOverhang(), $oObj.MeasureHelpLineDistance(), $oObj.MeasureHelpLine1Length(), _
				$oObj.MeasureHelpLine2Length(), $oObj.MeasureBelowReferenceEdge(), $oObj.MeasureDecimalPlaces(), $oObj.MeasureTextVerticalPosition(), _
				$oObj.MeasureTextHorizontalPosition(), _
				($oObj.MeasureTextRotate90()) ? (False) : (True), _ ; When MeasureTextRotate90 is True, $bParallel is False and vice versa.
				($oObj.MeasureShowUnit()) ? ($oObj.MeasureUnit()) : ($LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_OFF)) ; If MeasureShowUnit is True, return the Unit type, else indicate units are off.

		Return SetError($__LO_STATUS_SUCCESS, 1, $avDimension)
	EndIf

	If ($iDistance <> Null) Then
		If Not __LO_IntIsBetween($iDistance, -10008, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.MeasureLineDistance = $iDistance
		$iError = (__LO_IntIsBetween($oObj.MeasureLineDistance(), $iDistance - 1, $iDistance + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iGuideOverhang <> Null) Then
		If Not __LO_IntIsBetween($iGuideOverhang, -10008, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.MeasureHelpLineOverhang = $iGuideOverhang
		$iError = (__LO_IntIsBetween($oObj.MeasureHelpLineOverhang(), $iGuideOverhang - 1, $iGuideOverhang + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iGuideDistance <> Null) Then
		If Not __LO_IntIsBetween($iGuideDistance, -10008, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.MeasureHelpLineDistance = $iGuideDistance
		$iError = (__LO_IntIsBetween($oObj.MeasureHelpLineDistance(), $iGuideDistance - 1, $iGuideDistance + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLGuide <> Null) Then
		If Not __LO_IntIsBetween($iLGuide, -10008, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.MeasureHelpLine1Length = $iLGuide
		$iError = (__LO_IntIsBetween($oObj.MeasureHelpLine1Length(), $iLGuide - 1, $iLGuide + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRGuide <> Null) Then
		If Not __LO_IntIsBetween($iRGuide, -10008, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.MeasureHelpLine2Length = $iRGuide
		$iError = (__LO_IntIsBetween($oObj.MeasureHelpLine2Length(), $iRGuide - 1, $iRGuide + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bBelow <> Null) Then
		If Not IsBool($bBelow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.MeasureBelowReferenceEdge = $bBelow
		$iError = ($oObj.MeasureBelowReferenceEdge() = $bBelow) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iDecimal <> Null) Then
		If Not __LO_IntIsBetween($iDecimal, 0, 99) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.MeasureDecimalPlaces = $iDecimal
		$iError = (__LO_IntIsBetween($oObj.MeasureDecimalPlaces(), $iDecimal - 1, $iDecimal + 1)) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iVertPos <> Null) Then
		If Not __LO_IntIsBetween($iVertPos, $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_AUTO, $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_MIDDLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oObj.MeasureTextVerticalPosition = $iVertPos
		$iError = ($oObj.MeasureTextVerticalPosition() = $iVertPos) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iHoriPos <> Null) Then
		If Not __LO_IntIsBetween($iHoriPos, $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_AUTO, $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oObj.MeasureTextHorizontalPosition = $iHoriPos
		$iError = ($oObj.MeasureTextHorizontalPosition() = $iHoriPos) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bParallel <> Null) Then
		If Not IsBool($bParallel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oObj.MeasureTextRotate90 = ($bParallel) ? (False) : (True) ; When MeasureTextRotate90 is True, $bParallel is False and vice versa.
		$iError = ($oObj.MeasureTextRotate90() = ($bParallel) ? (False) : (True)) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iUnitType <> Null) Then
		If Not __LO_IntIsBetween($iUnitType, $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_OFF, $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_LINE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		If ($iUnitType = $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_OFF) Then
			$oObj.MeasureShowUnit = False
			$iError = ($oObj.MeasureShowUnit() = False) ? ($iError) : (BitOR($iError, 1024))

		Else
			$oObj.MeasureShowUnit = True
			$oObj.MeasureUnit = $iUnitType
			$iError = ($oObj.MeasureUnit() = $iUnitType) ? ($iError) : (BitOR($iError, 1024))
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_DimensionSettings

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateArrow
; Description ...: Create an Arrow type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateArrow(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (HMM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (HMM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iShapeType          - an integer value (0-25). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_ARROWS_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer. See $LOI_DRAWSHAPE_TYPE_ARROWS_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" or "com.sun.star.drawing.EllipseShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  @Error 2 @Extended 3 Return 0 = Failed to create "MirroredX" property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT,
;                  $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateArrow(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tProp2, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY
			$tProp.Value = "quad-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY
			$tProp.Value = "quad-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN
			$tProp.Value = "down-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT
			$tProp.Value = "left-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT
			$tProp.Value = "left-right-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT
			$tProp.Value = "right-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP
			$tProp.Value = "up-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN
			$tProp.Value = "up-down-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
			$tProp.Value = "mso-spt100"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CIRCULAR
			$tProp.Value = "circular-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT
			$tProp.Value = "corner-right-arrow" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_DOWN
			$tProp.Value = "down-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT
			$tProp.Value = "left-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT
			$tProp.Value = "left-right-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT
			$tProp.Value = "notched-right-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT
			$tProp.Value = "right-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT
			$tProp.Value = "split-arrow" ; "non-primitive"??

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED
			$tProp.Value = "s-sharped-arrow" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT
			$tProp.Value = "split-arrow" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT
			$tProp.Value = "striped-right-arrow" ; "mso-spt100"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP
			$tProp.Value = "up-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_DOWN
			$tProp.Value = "up-down-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT
			$tProp.Value = "up-right-arrow-callout" ; "mso-spt89"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN
			$tProp.Value = "up-right-down-arrow" ; "mso-spt100"

			$tProp2 = __LO_SetPropertyValue("MirroredX", True) ; Shape is an up and left arrow without this Property.
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

			ReDim $atCusShapeGeo[2]
			$atCusShapeGeo[1] = $tProp2

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_CHEVRON
			$tProp.Value = "chevron"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_PENTAGON
			$tProp.Value = "pentagon-right"
	EndSwitch

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateArrow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateBasic
; Description ...: Create a Basic type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateBasic(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (HMM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (HMM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iShapeType          - an integer value (26-49). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_BASIC_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer. See $LOI_DRAWSHAPE_TYPE_BASIC_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" or "com.sun.star.drawing.EllipseShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, $LOI_DRAWSHAPE_TYPE_BASIC_FRAME
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateBasic(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]
	Local $iCircleKind_CUT = 2 ; a circle with a cut connected by a line.
	Local $iCircleKind_ARC = 3 ; a circle with an open cut.

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iShapeType = $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT) Or ($iShapeType = $LOI_DRAWSHAPE_TYPE_BASIC_ARC) Then ; These two shapes need special procedures.
		$oShape = $oDoc.createInstance("com.sun.star.drawing.EllipseShape")
		If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Switch $iShapeType
			Case $LOI_DRAWSHAPE_TYPE_BASIC_ARC
				$oShape.Name = __LOImpress_GetShapeName($oSlide, "Elliptical arc ")
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Case $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT
				$oShape.Name = __LOImpress_GetShapeName($oSlide, "Ellipse Segment ")
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndSwitch

		$oSlide.add($oShape)

	Else
		$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
		If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tProp = __LO_SetPropertyValue("Type", "")
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oSlide.add($oShape)
	EndIf

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_BASIC_ARC
			$oShape.FillColor = $LO_COLOR_OFF

			$oShape.CircleKind = $iCircleKind_ARC
			$oShape.CircleStartAngle = 0
			$oShape.CircleEndAngle = 25000

		Case $LOI_DRAWSHAPE_TYPE_BASIC_ARC_BLOCK
			$tProp.Value = "block-arc"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE
			$tProp.Value = "circle-pie" ; "mso-spt100"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT
			$oShape.CircleKind = $iCircleKind_CUT
			$oShape.CircleStartAngle = 0
			$oShape.CircleEndAngle = 25000

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CROSS
			$tProp.Value = "cross"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CUBE
			$tProp.Value = "cube"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CYLINDER
			$tProp.Value = "can"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_DIAMOND
			$tProp.Value = "diamond"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE
			$tProp.Value = "ellipse"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_FOLDED_CORNER
			$tProp.Value = "paper"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_FRAME
			$tProp.Value = "frame" ; Not working

		Case $LOI_DRAWSHAPE_TYPE_BASIC_HEXAGON
			$tProp.Value = "hexagon"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_OCTAGON
			$tProp.Value = "octagon"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_PARALLELOGRAM
			$tProp.Value = "parallelogram"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE
			$tProp.Value = "rectangle"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED
			$tProp.Value = "round-rectangle"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_REGULAR_PENTAGON
			$tProp.Value = "pentagon"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_RING
			$tProp.Value = "ring"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_TRAPEZOID
			$tProp.Value = "trapezoid"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES
			$tProp.Value = "isosceles-triangle"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT
			$tProp.Value = "right-triangle"
	EndSwitch

	If ($iShapeType <> $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT) And ($iShapeType <> $LOI_DRAWSHAPE_TYPE_BASIC_ARC) Then
		$atCusShapeGeo[0] = $tProp
		$oShape.CustomShapeGeometry = $atCusShapeGeo

		; Settings for TextBox use.
		$oShape.TextMinimumFrameWidth = $iWidth
		$oShape.TextMinimumFrameHeight = $iHeight
		$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
		$oShape.TextAutoGrowHeight = False
		$oShape.TextAutoGrowWidth = False
	EndIf

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateBasic

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateCallout
; Description ...: Create a Callout type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateCallout(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (HMM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (HMM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iShapeType          - an integer value (50-56). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_CALLOUT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer. See $LOI_DRAWSHAPE_TYPE_CALLOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateCallout(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_CLOUD
			$tProp.Value = "cloud-callout"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_1
			$tProp.Value = "line-callout-1"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_2
			$tProp.Value = "line-callout-2"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_3
			$tProp.Value = "line-callout-3"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR
			$tProp.Value = "rectangular-callout"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED
			$tProp.Value = "round-rectangular-callout"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_ROUND
			$tProp.Value = "round-callout"
	EndSwitch

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateCallout

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateFlowchart
; Description ...: Create a FlowChart type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateFlowchart(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (HMM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (HMM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iShapeType          - an integer value (57-84). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_FLOWCHART_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer. See $LOI_DRAWSHAPE_TYPE_FLOWCHART_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateFlowchart(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_CARD
			$tProp.Value = "flowchart-card"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_COLLATE
			$tProp.Value = "flowchart-collate"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR
			$tProp.Value = "flowchart-connector"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE
			$tProp.Value = "flowchart-off-page-connector"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DATA
			$tProp.Value = "flowchart-data"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DECISION
			$tProp.Value = "flowchart-decision"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DELAY
			$tProp.Value = "flowchart-delay"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE
			$tProp.Value = "flowchart-direct-access-storage"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DISPLAY
			$tProp.Value = "flowchart-display"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DOCUMENT
			$tProp.Value = "flowchart-document"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_EXTRACT
			$tProp.Value = "flowchart-extract"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE
			$tProp.Value = "flowchart-internal-storage"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MAGNETIC_DISC
			$tProp.Value = "flowchart-magnetic-disk"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_INPUT
			$tProp.Value = "flowchart-manual-input"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_OPERATION
			$tProp.Value = "flowchart-manual-operation"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MERGE
			$tProp.Value = "flowchart-merge"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MULTIDOCUMENT
			$tProp.Value = "flowchart-multidocument"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_OR
			$tProp.Value = "flowchart-or"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PREPARATION
			$tProp.Value = "flowchart-preparation"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS
			$tProp.Value = "flowchart-process"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE
			$tProp.Value = "flowchart-alternate-process"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED
			$tProp.Value = "flowchart-predefined-process"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PUNCHED_TAPE
			$tProp.Value = "flowchart-punched-tape"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS
			$tProp.Value = "flowchart-sequential-access"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_SORT
			$tProp.Value = "flowchart-sort"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_STORED_DATA
			$tProp.Value = "flowchart-stored-data"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION
			$tProp.Value = "flowchart-summing-junction"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR
			$tProp.Value = "flowchart-terminator"
	EndSwitch

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateFlowchart

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateLine
; Description ...: Create a Line type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateLine(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (HMM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (HMM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iShapeType          - an integer value (85-92). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_LINE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer. See $LOI_DRAWSHAPE_TYPE_LINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create the requested Line type Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a Position structure.
;                  @Error 2 @Extended 3 Return 0 = Failed to create "com.sun.star.drawing.PolyPolygonBezierCoords" Structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateLine(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tSize, $tPos, $tPolyCoords, $tStart, $tEnd
	Local $atPoint[0], $aiFlags[0]
	Local $avArray[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iShapeType <> $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION) Then
		$tPolyCoords = __LO_CreateStruct("com.sun.star.drawing.PolyPolygonBezierCoords")
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)
	EndIf

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS To $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_STARTS_ARROW
			$oShape = $oDoc.createInstance("com.sun.star.drawing.LineShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Line ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[2]
			ReDim $aiFlags[2]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			Switch $iShapeType
				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_CIRCLE
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Circle"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_SQUARE
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Square"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_CIRCLE_ARROW
					$oShape.LineStartName = "Circle"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ENDS_ARROW
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_SQUARE_ARROW
					$oShape.LineStartName = "Square"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_STARTS_ARROW
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
			EndSwitch

		Case $LOI_DRAWSHAPE_TYPE_CONNECTOR To $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW
			$oShape = $oDoc.createInstance("com.sun.star.drawing.ConnectorShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Connector ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			$tStart = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($tStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$tEnd = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($tEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			Switch $iShapeType
				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_ARROWS
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_CURVE

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ARROWS
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_CURVE
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ENDS_ARROW
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_CURVE
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_ENDS_ARROW
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ARROWS
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ENDS_ARROW
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STRAIGHT

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ARROWS
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STRAIGHT
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STRAIGHT
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300
			EndSwitch

		Case $LOI_DRAWSHAPE_TYPE_LINE_CURVE
			$oShape = $oDoc.createInstance("com.sun.star.drawing.OpenBezierShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[4]
			ReDim $aiFlags[4]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight / 2))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint(Int($iX + $iWidth), $iY)
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED
			$oShape = $oDoc.createInstance("com.sun.star.drawing.ClosedBezierShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[4]
			ReDim $aiFlags[4]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight / 2))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint(Int($iX + $iWidth), $iY)
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = 7512015 ; Light blue

		Case $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION
			$oShape = $oDoc.createInstance("com.sun.star.drawing.MeasureShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oSlide.add($oShape)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Dimension Line ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tStart = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($tStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$tEnd = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($tEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		Case $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE
			$oShape = $oDoc.createInstance("com.sun.star.drawing.OpenFreeHandShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[3]
			ReDim $aiFlags[3]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight / 2))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

		Case $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED
			$oShape = $oDoc.createInstance("com.sun.star.drawing.ClosedFreeHandShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[4]
			ReDim $aiFlags[4]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth) + Int(($iX + $iWidth / 8)), Int(($iY + $iHeight / 2)))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = 7512015 ; Light blue

		Case $LOI_DRAWSHAPE_TYPE_LINE_LINE, $LOI_DRAWSHAPE_TYPE_LINE_LINE_45
			$oShape = $oDoc.createInstance("com.sun.star.drawing.LineShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Line ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[2]
			ReDim $aiFlags[2]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

		Case $LOI_DRAWSHAPE_TYPE_LINE_POLYGON, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45
			$oShape = $oDoc.createInstance("com.sun.star.drawing.PolyLineShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Polygon 4 corners ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[5]
			ReDim $aiFlags[5]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth), $iY)
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint($iX, Int($iY + $iHeight))
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[4] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[4]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[4] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED
			$oShape = $oDoc.createInstance("com.sun.star.drawing.PolyPolygonShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Polygon 4 corners ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[5]
			ReDim $aiFlags[5]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth), $iY)
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint($iX, Int($iY + $iHeight))
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[4] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[4]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[4] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = 7512015 ; Light blue
	EndSwitch

	If __LO_IntIsBetween($iShapeType, $LOI_DRAWSHAPE_TYPE_CONNECTOR, $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW, "", $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION) Then
		$oShape.StartPosition = $tStart
		$oShape.EndPosition = $tEnd

	Else
		$avArray[0] = $atPoint
		$tPolyCoords.Coordinates = $avArray

		$avArray[0] = $aiFlags
		$tPolyCoords.Flags = $avArray

		$oShape.PolyPolygonBezier = $tPolyCoords
	EndIf

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateStars
; Description ...: Create a Star or Banner type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateStars(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (HMM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (HMM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iShapeType          - an integer value (93-104). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_STARS_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer. See $LOI_DRAWSHAPE_TYPE_STARS_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOI_DRAWSHAPE_TYPE_STARS_6_POINT, $LOI_DRAWSHAPE_TYPE_STARS_12_POINT, $LOI_DRAWSHAPE_TYPE_STARS_SIGNET, $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateStars(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_STARS_4_POINT
			$tProp.Value = "star4"

		Case $LOI_DRAWSHAPE_TYPE_STARS_5_POINT
			$tProp.Value = "star5"

		Case $LOI_DRAWSHAPE_TYPE_STARS_6_POINT
			$tProp.Value = "star6" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE
			$tProp.Value = "concave-star6" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_STARS_8_POINT
			$tProp.Value = "star8"

		Case $LOI_DRAWSHAPE_TYPE_STARS_12_POINT
			$tProp.Value = "star12" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_STARS_24_POINT
			$tProp.Value = "star24"

		Case $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE
			$tProp.Value = "mso-spt21" ; "doorplate"

		Case $LOI_DRAWSHAPE_TYPE_STARS_EXPLOSION
			$tProp.Value = "bang"

		Case $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_HORIZONTAL
			$tProp.Value = "horizontal-scroll"

		Case $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_VERTICAL
			$tProp.Value = "vertical-scroll"

		Case $LOI_DRAWSHAPE_TYPE_STARS_SIGNET
			$tProp.Value = "signet" ; "non-primitive"
	EndSwitch

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateStars

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateSymbol
; Description ...: Create a Symbol type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateSymbol(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (HMM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (HMM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iShapeType          - an integer value (105-122). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_SYMBOL_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer. See $LOI_DRAWSHAPE_TYPE_SYMBOL_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD, $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER, $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
;                  The following shape is visually different from the manually inserted one in L.O. 7.3.4.2:
;                  $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateSymbol(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
			$tProp.Value = "col-502ad400"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON
			$tProp.Value = "col-60da8460"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_SQUARE
			$tProp.Value = "quad-bevel"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_DOUBLE
			$tProp.Value = "brace-pair"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_LEFT
			$tProp.Value = "left-brace"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_RIGHT
			$tProp.Value = "right-brace"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_DOUBLE
			$tProp.Value = "bracket-pair"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_LEFT
			$tProp.Value = "left-bracket"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_RIGHT
			$tProp.Value = "right-bracket"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "cloud"
			$tProp.Value = "cloud"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "flower"
			$tProp.Value = "flower"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_HEART
			$tProp.Value = "heart"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "lightning"
			$tProp.Value = "lightning"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_MOON
			$tProp.Value = "moon"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_SMILEY
			$tProp.Value = "smiley"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_SUN
			$tProp.Value = "sun"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_PROHIBITED
			$tProp.Value = "forbidden"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE
			$tProp.Value = "puzzle"
	EndSwitch

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateSymbol

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_GetCustomType
; Description ...: Return the Shape Type Constant corresponding to the Custom Shape Type string.
; Syntax ........: __LOImpress_DrawShape_GetCustomType($sCusShapeType)
; Parameters ....: $sCusShapeType       - a string value. The Returned Custom Shape Type Value from CustomShapeGeometry Array of properties.
; Return values .: Success: Integer or -1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sCusShapeType not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Custom Shape Type was successfully identified. Returning the Constant value of the Shape, see Constants $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 0 @Extended 0 Return -1 = Success. Custom Shape is of an unimplemented type that has an ambiguous name, and cannot be identified. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Some shapes are not implemented, or not fully implemented into LibreOffice for automation, consequently they do not have appropriate type names as of yet. Many have simply ambiguous names, such as "non-primitive".
;                  #1 Because of this the following shape types cannot be identified, and this function will return -1:
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT, known as "mso-spt100".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, known as "non-primitive", should be "corner-right-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, known as "non-primitive", should be "right-left-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, known as "non-primitive", should be "s-sharped-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, known as "non-primitive", should be "split-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT, known as "mso-spt100", should be "striped-right-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT, known as "mso-spt89", should be "up-right-arrow-callout".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, known as "mso-spt100", should be "up-right-down-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, known as "mso-spt100", should be "circle-pie".
;                  - $LOI_DRAWSHAPE_TYPE_STARS_6_POINT, known as "non-primitive", should be "star6".
;                  - $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE, known as "non-primitive", should be "concave-star6".
;                  - $LOI_DRAWSHAPE_TYPE_STARS_12_POINT, known as "non-primitive", should be "star12".
;                  - $LOI_DRAWSHAPE_TYPE_STARS_SIGNET, known as "non-primitive", should be "signet".
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD, known as "non-primitive", should be "cloud"?
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER, known as "non-primitive", should be "flower"?
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING, known as "non-primitive", should be "lightning".
;                  #2 The following Shapes implement the same type names, and are consequently indistinguishable:
;                  - $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE, $LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE is returned for either one.)
;                  - $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE is returned for either one.)
;                  - $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED is returned for either one.)
;                  #3 The following Shapes have strange names that may change in the future, but currently are able to be identified:
;                  - $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE, known as, "mso-spt21", should be "doorplate"
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, known as, "col-502ad400", should be ??
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, known as, "col-60da8460", should be ??
;                  #4 The following Shapes are customizable one to another, and are consequently indistinguishable:
;                  - $LOI_DRAWSHAPE_TYPE_FONTWORK_* (The Value of $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL is returned for any of these.)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_GetCustomType($sCusShapeType)
	If Not IsString($sCusShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $sCusShapeType
		Case "quad-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY)

		Case "quad-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY)

		Case "down-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN)

		Case "left-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT)

		Case "left-right-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT)

		Case "right-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT)

		Case "up-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP)

		Case "up-down-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN)

;~ 	Case "mso-spt100" ; Can't include this one as other shapes return mso-spt100 also
;~ Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT)

		Case "circular-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CIRCULAR)

		Case "corner-right-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT)

		Case "down-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_DOWN)

		Case "left-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT)

		Case "left-right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT)

		Case "notched-right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT)

		Case "right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT)

		Case "right-left-arrow" ; "non-primitive"??

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT)

		Case "s-sharped-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED)

		Case "split-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT)

		Case "striped-right-arrow" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT)

		Case "up-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP)

		Case "up-down-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_DOWN)

		Case "up-right-arrow-callout", "mso-spt89" ; "mso-spt89"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT)

		Case "up-right-down-arrow" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN)

		Case "chevron"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_CHEVRON)

		Case "pentagon-right"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_PENTAGON)

		Case "block-arc"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_ARC_BLOCK)

		Case "circle-pie" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE)

		Case "cross"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CROSS)

		Case "cube"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CUBE)

		Case "can"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CYLINDER)

		Case "diamond"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_DIAMOND)

		Case "ellipse"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE)
;~ $LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE

		Case "paper"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_FOLDED_CORNER)

		Case "frame" ; Not working

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_FRAME)

		Case "hexagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_HEXAGON)

		Case "octagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_OCTAGON)

		Case "parallelogram"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_PARALLELOGRAM)

		Case "rectangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE)
;~ $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE

		Case "round-rectangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED)
;~ $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED

		Case "pentagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_REGULAR_PENTAGON)

		Case "ring"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_RING)

		Case "trapezoid"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_TRAPEZOID)

		Case "isosceles-triangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES)

		Case "right-triangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT)

		Case "cloud-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_CLOUD)

		Case "line-callout-1"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_1)

		Case "line-callout-2"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_2)

		Case "line-callout-3"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_3)

		Case "rectangular-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR)

		Case "round-rectangular-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED)

		Case "round-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_ROUND)

		Case "flowchart-card"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_CARD)

		Case "flowchart-collate"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_COLLATE)

		Case "flowchart-connector"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR)

		Case "flowchart-off-page-connector"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE)

		Case "flowchart-data"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DATA)

		Case "flowchart-decision"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DECISION)

		Case "flowchart-delay"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DELAY)

		Case "flowchart-direct-access-storage"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE)

		Case "flowchart-display"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DISPLAY)

		Case "flowchart-document"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DOCUMENT)

		Case "flowchart-extract"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_EXTRACT)

		Case "flowchart-internal-storage"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE)

		Case "flowchart-magnetic-disk"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MAGNETIC_DISC)

		Case "flowchart-manual-input"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_INPUT)

		Case "flowchart-manual-operation"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_OPERATION)

		Case "flowchart-merge"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MERGE)

		Case "flowchart-multidocument"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MULTIDOCUMENT)

		Case "flowchart-or"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_OR)

		Case "flowchart-preparation"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PREPARATION)

		Case "flowchart-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS)

		Case "flowchart-alternate-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE)

		Case "flowchart-predefined-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED)

		Case "flowchart-punched-tape"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PUNCHED_TAPE)

		Case "flowchart-sequential-access"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS)

		Case "flowchart-sort"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_SORT)

		Case "flowchart-stored-data"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_STORED_DATA)

		Case "flowchart-summing-junction"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION)

		Case "flowchart-terminator"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR)

		Case "fontwork-arch-down-pour", "fontwork-arch-left-pour", "fontwork-arch-right-pour", "fontwork-arch-up-pour", "fontwork-arch-down-curve", _
				"fontwork-arch-left-curve", "fontwork-arch-right-curve", "fontwork-arch-up-curve", "fontwork-chevron-down", "fontwork-chevron-up", _
				"fontwork-circle-curve", "fontwork-circle-pour", "fontwork-curve-down", "fontwork-curve-up", "fontwork-fade-down", "fontwork-fade-left", _
				"fontwork-fade-right", "fontwork-fade-up", "fontwork-fade-up-and-left", "fontwork-fade-up-and-right", "fontwork-inflate", "fontwork-slant-down", _
				"fontwork-slant-up", "fontwork-stop", "fontwork-triangle-up", "fontwork-triangle-down", "fontwork-open-circle-curve", "fontwork-open-circle-pour", _
				"fontwork-plain-text", "fontwork-wave"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL) ; Can't differentiate reliably

		Case "star4"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_4_POINT)

		Case "star5"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_5_POINT)

		Case "star6" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_6_POINT)

		Case "concave-star6" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE)

		Case "star8"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_8_POINT)

		Case "star12" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_12_POINT)

		Case "star24"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_24_POINT)

		Case "mso-spt21", "doorplate" ; "doorplate"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE)

		Case "bang"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_EXPLOSION)

		Case "horizontal-scroll"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_HORIZONTAL)

		Case "vertical-scroll"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_VERTICAL)

		Case "signet" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_SIGNET)

		Case "col-502ad400"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND)

		Case "col-60da8460"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON)

		Case "quad-bevel"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_SQUARE)

		Case "brace-pair"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_DOUBLE)

		Case "left-brace"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_LEFT)

		Case "right-brace"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_RIGHT)

		Case "bracket-pair"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_DOUBLE)

		Case "left-bracket"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_LEFT)

		Case "right-bracket"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_RIGHT)

		Case "cloud"
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "cloud"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD)

		Case "flower"
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "flower"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER)

		Case "heart"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_HEART)

		Case "lightning"
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "lightning"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING)

		Case "moon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_MOON)

		Case "smiley"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_SMILEY)

		Case "sun"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_SUN)

		Case "forbidden"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_PROHIBITED)

		Case "puzzle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE)

		Case Else

			Return SetError($__LO_STATUS_SUCCESS, 0, -1)
	EndSwitch
EndFunc   ;==>__LOImpress_DrawShape_GetCustomType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShapePointGetSettings
; Description ...: Retrieve the current settings for a particular point in a shape.
; Syntax ........: __LOImpress_DrawShapePointGetSettings(ByRef $avArray, ByRef $aiFlags, ByRef $atPoints, $iArrayElement)
; Parameters ....: $avArray             - [in/out] an array of variants. An array to fill with settings. Array will be directly modified.
;                  $aiFlags             - [in/out] an array of integers. An Array of Point Type Flags returned from the Shape.
;                  $atPoints            - [in/out] an array of dll structs. An Array of Points returned from the Shape.
;                  $iArrayElement       - an integer value. The Array element that contains the point to retrieve the settings for.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $avArray is not an Array.
;                  @Error 1 @Extended 2 Return 0 = $aiFlags is not an Array.
;                  @Error 1 @Extended 3 Return 0 = $atPoints is not an Array.
;                  @Error 1 @Extended 4 Return 0 = $iArrayElement not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the current X coordinate.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the current Y coordinate.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the current Point Type Flag.
;                  @Error 3 @Extended 4 Return 0 = Failed to determine if the Point is a Curve or not.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Current Settings were successfully retrieved, $avArray has been filled with the current settings.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShapePointGetSettings(ByRef $avArray, ByRef $aiFlags, ByRef $atPoints, $iArrayElement)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iX, $iY, $iPointType
	Local $bIsCurve

	If Not IsArray($avArray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsArray($atPoints) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iArrayElement) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iX = $atPoints[$iArrayElement].X()
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$avArray[0] = $iX

	$iY = $atPoints[$iArrayElement].Y()
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$avArray[1] = $iY

	$iPointType = $aiFlags[$iArrayElement]
	If Not IsInt($iPointType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$avArray[2] = $iPointType

	If ($iPointType = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then
		If ($iArrayElement <> (UBound($atPoints) - 1)) Then ; Requested point is not at the end of the array of points.

			If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Point after requested point is a Control Point.
				; If a Point and the following Control point have the same coordinates, the point is not a curve.
				$bIsCurve = (($atPoints[$iArrayElement].X() = $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() = $atPoints[$iArrayElement + 1].Y())) ? (False) : (True)

			Else ; Next point after requested point is not a control type point.
				$bIsCurve = False
			EndIf

		Else ; Point is the last point, cant be a curve.
			$bIsCurve = False
		EndIf

	Else ; Point is a Smooth, or Symmetrical Point type, point is a curve regardless.
		$bIsCurve = True
	EndIf

	If Not IsBool($bIsCurve) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$avArray[3] = $bIsCurve

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_DrawShapePointGetSettings

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShapePointModify
; Description ...: Internal function for modifying A Shape's Points.
; Syntax ........: __LOImpress_DrawShapePointModify(ByRef $aiFlags, ByRef $atPoints, ByRef $iArrayElement[, $iX = Null[, $iY = Null[, $iPointType = Null[, $bIsCurve = Null]]]])
; Parameters ....: $aiFlags             - [in/out] an array of integers. An Array of Point Type Flags returned from the Shape. Array will be directly modified.
;                  $atPoints            - [in/out] an array of dll structs. An Array of Points returned from the Shape. Array will be directly modified.
;                  $iArrayElement       - [in/out] an integer value. The Array element that contains the point to modify. This may be directly modified, depending on the settings.
;                  $iX                  - [optional] an integer value. Default is Null. The X coordinate value, set in Hundredths of a Millimeter (HMM).
;                  $iY                  - [optional] an integer value. Default is Null. The Y coordinate value, set in Hundredths of a Millimeter (HMM).
;                  $iPointType          - [optional] an integer value (0,1,3). Default is Null. The Type of Point to change the called point to. See Remarks. See constants $LOI_DRAWSHAPE_POINT_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $bIsCurve            - [optional] a boolean value. Default is Null. If True, the Normal Point is a Curve. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $aiFlags not an Array.
;                  @Error 1 @Extended 2 Return 0 = $atPoints not an Array.
;                  @Error 1 @Extended 3 Return 0 = $iArrayElement not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a new Position Point Structure for the First Control Point.
;                  @Error 2 @Extended 2 Return 0 = Failed to Create a new Position Point Structure for the Second Control Point.
;                  @Error 2 @Extended 3 Return 0 = Failed to Create a new Position Point Structure for the Third Control Point.
;                  @Error 2 @Extended 4 Return 0 = Failed to Create a new Position Point Structure for the Fourth Control Point.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify the next position point in the shape.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify the previous position point in the shape.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;                  Only $LOI_DRAWSHAPE_TYPE_LINE_* type shapes have Points that can be added to, removed, or modified.
;                  This is a homemade function as LibreOffice doesn't offer an easy way for modifying points in a shape. Consequently this will not produce similar results as when working with LibreOffice manually, and may wreck your shape's shape. Use with caution.
;                  For an unknown reason, I am unable to insert "SMOOTH" Points, and consequently, any smooth Points are reverted back to "Normal" points, but still having their Smooth control points upon insertion that were already present in the shape. If you modify a point to "SMOOTH" type, it will be, for now, replaced with "Symmetrical".
;                  The first and last points in a shape can only be a "Normal" Point Type. The last point cannot be Curved, but the first can be.
;                  Calling and Smooth or Symmetrical point types with $bIsCurve = True, will be ignored, as they are already a curve.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShapePointModify(ByRef $aiFlags, ByRef $atPoints, ByRef $iArrayElement, $iX = Null, $iY = Null, $iPointType = Null, $bIsCurve = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iNextArrayElement, $iPreviousArrayElement, $iSymmetricalPointXValue, $iSymmetricalPointYValue, $iOffset, $iForOffset, $iReDimCount
	Local $tControlPoint1, $tControlPoint2, $tControlPoint3, $tControlPoint4
	Local $avArray[0], $avArray2[0]

	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($atPoints) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iArrayElement) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Error if point called is not between 0 or number of points.

	If ($iArrayElement <> UBound($atPoints) - 1) Then ; If The requested point to be modified is not at the end of the Array of points, find the next regular point.

		For $i = ($iArrayElement + 1) To UBound($aiFlags) - 1 ; Locate the next non-Control Point in the Array for later use.
			If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$iNextArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		If Not IsInt($iNextArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else
		$iNextArrayElement = -1
	EndIf

	If ($iArrayElement > 0) Then ; If Point requested is not the first point, find the previous Point's position.

		For $i = ($iArrayElement - 1) To 0 Step -1 ; Locate the previous non-Control Point in the Array for later use.
			If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$iPreviousArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		If Not IsInt($iPreviousArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Else
		$iPreviousArrayElement = -1
	EndIf

	If ($iX <> Null) Then
		If ($iArrayElement < UBound($atPoints) - 1) And ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Next point is a control point, check if this point is a curve.

			If ($atPoints[$iArrayElement].X() = $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() = $atPoints[$iArrayElement + 1].Y()) Then ; Update the coordinates, because the point is not a curve.
				$atPoints[$iArrayElement + 1].X = $iX
			EndIf
		EndIf

		$atPoints[$iArrayElement].X = $iX
	EndIf

	If ($iY <> Null) Then
		If ($iArrayElement < UBound($atPoints) - 1) And ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Next point is a control point, check if this point is a curve.

			If ($atPoints[$iArrayElement].X() = $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() = $atPoints[$iArrayElement + 1].Y()) Then ; Update the coordinates, because the point is not a curve.
				$atPoints[$iArrayElement + 1].Y = $iY
			EndIf
		EndIf

		$atPoints[$iArrayElement].Y = $iY
	EndIf

	If ($iPointType <> Null) Then
		If ($iPointType <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; New point type is a curve.

			If ($aiFlags[$iArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; Converting point from Normal to a curve.

				; Pick the lowest X and Y value difference between previous point and current point and Next point and current Point.
				$iSymmetricalPointXValue = ((($atPoints[$iArrayElement].X() - $atPoints[$iPreviousArrayElement].X()) * .5) < (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)) ? Int((($atPoints[$iArrayElement].X() - $atPoints[$iPreviousArrayElement].X()) * .5)) : Int((($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5))
				$iSymmetricalPointYValue = ((($atPoints[$iArrayElement].Y() - $atPoints[$iPreviousArrayElement].Y()) * .5) < (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)) ? Int((($atPoints[$iArrayElement].Y() - $atPoints[$iPreviousArrayElement].Y()) * .5)) : Int((($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5))

				If ($aiFlags[$iArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; previous point is a control Point, might just need to modify it.

					If (($iArrayElement - 2 > $iPreviousArrayElement) And $aiFlags[$iArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; there are two control points before this point, I can just modify the first point before.
						$tControlPoint1 = $atPoints[$iArrayElement - 2]

						$tControlPoint2 = $atPoints[$iArrayElement - 1]
						$tControlPoint2.X = ($atPoints[$iArrayElement].X() - $iSymmetricalPointXValue)
						$tControlPoint2.Y = ($atPoints[$iArrayElement].Y() - $iSymmetricalPointYValue)

					Else ; There is only one control point, I need to create a new one.
						$tControlPoint1 = $atPoints[$iArrayElement - 1]

						$tControlPoint2 = __LOImpress_CreatePoint($atPoints[$iArrayElement].X() - $iSymmetricalPointXValue, $atPoints[$iArrayElement].Y() - $iSymmetricalPointYValue)
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
					EndIf

				Else ; Previous point is a normal point, need to create new control points.
					$tControlPoint1 = __LOImpress_CreatePoint($atPoints[$iPreviousArrayElement].X(), $atPoints[$iPreviousArrayElement].Y())
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tControlPoint2 = __LOImpress_CreatePoint($atPoints[$iArrayElement].X() - $iSymmetricalPointXValue, $atPoints[$iArrayElement].Y() - $iSymmetricalPointYValue)
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
				EndIf

				If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Next point is a control Point, might just need to modify it.

					If (($iArrayElement + 2 < $iNextArrayElement) And $aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; there are two control points after this point, I can just modify the first point after.
						$tControlPoint4 = $atPoints[$iArrayElement + 2]

						$tControlPoint3 = $atPoints[$iArrayElement + 1]
						$tControlPoint3.X = ($atPoints[$iArrayElement].X() + $iSymmetricalPointXValue)
						$tControlPoint3.Y = ($atPoints[$iArrayElement].Y() + $iSymmetricalPointYValue)

					Else ; There is only one control point, I need to create a new one and modify the other.
						$tControlPoint3 = __LOImpress_CreatePoint($atPoints[$iArrayElement].X() + $iSymmetricalPointXValue, $atPoints[$iArrayElement].Y() + $iSymmetricalPointYValue)
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

						$tControlPoint4 = $atPoints[$iArrayElement + 1] ; Modify the Control Point.
						$tControlPoint4.X = ($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5))
						$tControlPoint4.Y = ($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5))
					EndIf

				Else ; Next point is a normal point, need to create new control points.
					$tControlPoint3 = __LOImpress_CreatePoint(($atPoints[$iArrayElement].X() + $iSymmetricalPointXValue), ($atPoints[$iArrayElement].Y() + $iSymmetricalPointYValue))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

					$tControlPoint4 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 4
				; Check if there already was 4 control point present around this point I am modifying.
				$iReDimCount -= ($aiFlags[$iArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0)
				$iReDimCount -= (($iArrayElement - 2 > $iPreviousArrayElement) And ($aiFlags[$iArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)
				$iReDimCount -= ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0)
				$iReDimCount -= (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) + $iReDimCount]
				ReDim $avArray2[UBound($aiFlags) + $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the points to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iPreviousArrayElement) Then ; Insert the new or modified control points.

						$avArray[$i + 1] = $tControlPoint1
						$avArray2[$i + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$avArray[$i + 2] = $tControlPoint2
						$avArray2[$i + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$avArray[$i + 3] = $atPoints[$iArrayElement]
						$avArray2[$i + 3] = $iPointType
						$avArray[$i + 4] = $tControlPoint3
						$avArray2[$i + 4] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$avArray[$i + 5] = $tControlPoint4
						$avArray2[$i + 5] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						$iOffset = 1 ; Add one to offset to skip the point I am modifying.
						$iOffset += ($aiFlags[$iArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am modifying has a control point before it, I need to skip them in the PointsArray.
						$iOffset += (($iArrayElement - 2 > $iPreviousArrayElement) And ($aiFlags[$iArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0) ; If the point I am modifying has two control points before it, I need to skip them in the PointsArray.
						$iOffset += ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am modifying has a control point after it, I need to skip them in the PointsArray.
						$iOffset += (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0) ; If the point I am modifying has two control points after it, I need to skip them in the PointsArray.

						$iForOffset += 5 ; Add to $i to skip the elements I manually added.
					EndIf

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
				Next

				; Update the ArrayElement value to its new position.
				$iArrayElement += ($aiFlags[$iArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (0) : (1) ; If the point I am modifying has a control point before it, don't add one to array element, because I didn't have to create and insert a new control point.
				$iArrayElement += (($iArrayElement - 2 > $iPreviousArrayElement) And ($aiFlags[$iArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (0) : (1) ; If the point I am modifying has two control points before it, don't add one to array element, because I didn't have to create and insert a new control point.

				$atPoints = $avArray
				$aiFlags = $avArray2

			Else ; Point is already a curve.
				; Do nothing?
			EndIf

		Else ; New Point is a Normal Point.
			If ($aiFlags[$iArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; Point being modified is not a normal type of point.

				If ($aiFlags[$iPreviousArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; If previous point is a normal point, see if I need to delete control points or not.

					If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Point after previous point is a control point, see if previous point is a curved point.

						If ($atPoints[$iPreviousArrayElement].X() <> $atPoints[$iPreviousArrayElement + 1].X()) And ($atPoints[$iPreviousArrayElement].Y() <> $atPoints[$iPreviousArrayElement + 1].Y()) Then
							; Previous Point is a Curved normal point, copy the control points present.

							$tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]

							If ($iPreviousArrayElement + 2 < $iArrayElement) And ($atPoints[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2] ; If two control points are present, copy them.
						EndIf
					EndIf

				Else ; Previous point is not a normal point.
					; Copy Control Points present.

					If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]

					If ($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2] ; If two control points are present, copy them.
				EndIf

				If ($aiFlags[$iNextArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then
					; Next point is a curve of some form, copy the control points.

					If ($aiFlags[$iNextArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint4 = $atPoints[$iNextArrayElement - 1]

					If ($iNextArrayElement - 2 > $iArrayElement) And ($aiFlags[$iNextArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint3 = $atPoints[$iNextArrayElement - 2] ; If two control points are present, copy them.
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 4
				; Check how many control points I am keeping.
				$iReDimCount -= (IsObj($tControlPoint1)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint2)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint3)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint4)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) - $iReDimCount]
				ReDim $avArray2[UBound($aiFlags) - $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i + $iOffset] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i + $iOffset] ; Add the rest of the points to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iPreviousArrayElement) Then ; Insert the old control points or remove them.

						If IsObj($tControlPoint1) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint1
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						If IsObj($tControlPoint2) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint2
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If (($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						$iForOffset += 1
						$iOffset += 1
						$avArray[$i + $iForOffset] = $atPoints[$iArrayElement]
						$avArray2[$i + $iForOffset] = $iPointType

						If IsObj($tControlPoint3) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint3
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iNextArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1
						EndIf

						If IsObj($tControlPoint3) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint4
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If (($iNextArrayElement - 2 > $iArrayElement) And ($aiFlags[$iNextArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1
						EndIf
					EndIf

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
				Next

				; Update the ArrayElement value to its new position.
				$iArrayElement -= (IsObj($tControlPoint1)) ? (0) : (1) ; If ControlPoint 1 is a object, it means I copied it, meaning I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.
				$iArrayElement -= (IsObj($tControlPoint2)) ? (0) : (1) ; If ControlPoint 2 is a object, it means I copied it, meaning I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.

				$atPoints = $avArray
				$aiFlags = $avArray2

			Else ; Point being modified is a normal point already.
				; Do nothing?
			EndIf
		EndIf
	EndIf

	If ($bIsCurve <> Null) Then
		If ($aiFlags[$iArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; If Point to modify is a normal point, then proceed, else point is a curve already.

			If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Point after point to modify is a control point, just modify it.
				$tControlPoint3 = $atPoints[$iArrayElement + 1]

				If ($bIsCurve = True) Then
					$tControlPoint3.X = ($atPoints[$iArrayElement].X() + (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5))
					$tControlPoint3.Y = ($atPoints[$iArrayElement].Y() + (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5))

					If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then
						$tControlPoint4 = $atPoints[$iArrayElement + 2] ; Copy the second control point.

					Else ; Create a new control point.
						$tControlPoint4 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
					EndIf

				ElseIf ($bIsCurve = False) And ($aiFlags[$iNextArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; Next point is a curve, so just modify the control point.
					$tControlPoint3.X = $atPoints[$iArrayElement].X() ; When the control point after a point has the same coordinates, it means it is not a curve.
					$tControlPoint3.Y = $atPoints[$iArrayElement].Y()
					; Copy the second control point if it exists.
					If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $tControlPoint4 = $atPoints[$iArrayElement + 2]

				Else ; IsCurve = False, and next point is normal. delete control points.
					$tControlPoint3 = Null
				EndIf

			Else ; Need to create new control points if IsCurve = True.
				If ($bIsCurve = True) Then
					$tControlPoint3 = __LOImpress_CreatePoint(Int($atPoints[$iArrayElement].X() + (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iArrayElement].Y() + (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

					$tControlPoint4 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
				EndIf
			EndIf

			$iOffset = 0
			$iForOffset = 0
			$iReDimCount = 0
			; Check how many control points I am keeping vs creating.
			$iReDimCount += (IsObj($tControlPoint3)) ? (1) : (0)
			$iReDimCount += (IsObj($tControlPoint4)) ? (1) : (0)
			$iReDimCount -= ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If a control point already existed, minus one from ReDim as it is either not new, or I am deleting it.
			$iReDimCount -= (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)

			ReDim $avArray[UBound($atPoints) + $iReDimCount]
			ReDim $avArray2[UBound($aiFlags) + $iReDimCount]

			$iReDimCount = 0

			For $i = 0 To UBound($atPoints) - 1
				If ($iOffset = 0) Then
					$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
					$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the points to the array.

				Else
					$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
					$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
				EndIf

				If ($i = $iArrayElement) Then ; Insert the new or modified control points.

					If IsObj($tControlPoint3) Then
						$iForOffset += 1
						If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.

						$avArray[$i + $iForOffset] = $tControlPoint3
						$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

					Else
						If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
					EndIf

					If IsObj($tControlPoint4) Then
						$iForOffset += 1
						If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1

						$avArray[$i + $iForOffset] = $tControlPoint4
						$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

					Else
						If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
					EndIf
				EndIf

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$atPoints = $avArray
			$aiFlags = $avArray2

		Else ; Point is a Curve, see if bIsCurve = False.
			If ($bIsCurve = False) Then ; If bIsCurve = True, I can just skip it, as there is nothing to do when the point is a curve already.

				If ($aiFlags[$iNextArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; Next point is a curve, need to keep the control points.
					If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint3 = $atPoints[$iArrayElement + 1]
					If ($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint4 = $atPoints[$iArrayElement + 2]
				EndIf

				If ($iPreviousArrayElement <> -1) And ($aiFlags[$iPreviousArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; There is a previous point, and it is a curve, I need to keep the control points.
					If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]
					If ($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2]

				ElseIf ($iPreviousArrayElement <> -1) And ($aiFlags[$iPreviousArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; There is a previous point, and it is a normal point.
					; See if it is curved.

					If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And _
							(($atPoints[$iPreviousArrayElement].X() <> $atPoints[$iPreviousArrayElement + 1].X()) And _
							($atPoints[$iPreviousArrayElement].Y() <> $atPoints[$iPreviousArrayElement + 1].Y())) Then ; Previous Point is a curve, need to keep the control points.
						$tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]

						If ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2]
					EndIf
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 4
				; Check how many control points I am keeping vs deleting.
				$iReDimCount -= (IsObj($tControlPoint1)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint2)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint3)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint4)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) - $iReDimCount]
				ReDim $avArray2[UBound($aiFlags) - $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the points to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iPreviousArrayElement) Then
						If IsObj($tControlPoint1) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint1
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						If IsObj($tControlPoint2) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint2
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If (($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

					ElseIf ($i = $iArrayElement) Then ; Insert or skip Control Points as necessary.
						$avArray[$i] = $atPoints[$iArrayElement]
						$avArray2[$i] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

						If IsObj($tControlPoint3) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + 1] = $tControlPoint3
							$avArray2[$i + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						If IsObj($tControlPoint4) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + 2] = $tControlPoint4
							$avArray2[$i + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If (($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf
					EndIf

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
				Next

				; Update the ArrayElement value to its new position.
				If ($iPreviousArrayElement <> -1) Then $iArrayElement -= ((IsObj($tControlPoint2) And ($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL))) ? (0) : (1) ; If ControlPoint 2 is a object, it means I copied it, meaining I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.
				If ($iPreviousArrayElement <> -1) Then $iArrayElement -= ((IsObj($tControlPoint1) And ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL))) ? (0) : (1) ; If ControlPoint 1 is a object, it means I copied it, meaning I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.

				$atPoints = $avArray
				$aiFlags = $avArray2
			EndIf
		EndIf
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_DrawShapePointModify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_FilterNameGet
; Description ...: Retrieves the correct L.O. Filter name for use in SaveAs and Export.
; Syntax ........: __LOImpress_FilterNameGet(ByRef $sDocSavePath[, $bExportFilters = False])
; Parameters ....: $sDocSavePath        - [in/out] a string value. Full path with extension.
;                  $bExportFilters      - [optional] a boolean value. Default is False. If True, includes the Filter Names that can be used to Export only, in the search.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sDocSavePath is not a string.
;                  @Error 1 @Extended 2 Return 0 = $bExportFilters not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sDocSavePath is not a correct path or URL.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Success. Returning required filter name from "SaveAs" Filter Names.
;                  @Error 0 @Extended 2 Return String = Success. Returning required filter name from "Export" Filter Names.
;                  @Error 0 @Extended 3 Return String = Filter Name not found for given file extension, defaulting to .odp file format and updating save path accordingly.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Searches a predefined list of extensions stored in an array. Not all FilterNames are listed.
;                  For finding your own Filter Names, see convertfilters.html found in L.O. Install Folder: LibreOffice\help\en-US\text\shared\guide
;                  Or See: "OOME_3_0", "OpenOffice.org Macros Explained OOME Third Edition" by Andrew D. Pitonyak, which has a handy Macro for listing all Filter Names, found on page 284 of the above book in the ODT format.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_FilterNameGet(ByRef $sDocSavePath, $bExportFilters = False)
	Local $iLength, $iSlashLocation, $iDotLocation
	Local Const $STR_NOCASESENSE = 0, $STR_STRIPALL = 8
	Local $sFileExtension, $sFilterName
	Local $msSaveAsFilters[], $msExportFilters[]

	If Not IsString($sDocSavePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExportFilters) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iLength = StringLen($sDocSavePath)

	$msSaveAsFilters[".fodp"] = "OpenDocument Presentation Flat XML"
	$msSaveAsFilters[".pot"] = "PowerPoint 3"
	$msSaveAsFilters[".potx"] = "Impress MS PowerPoint 2007 XML Template" ; Note these have a XML version too.
	$msSaveAsFilters[".pps"] = "MS PowerPoint 97 AutoPlay"
	$msSaveAsFilters[".ppt"] = "PowerPoint 3"
	$msSaveAsFilters[".ppsx"] = "Impress MS PowerPoint 2007 XML AutoPlay" ; Note these have a XML version too.
	$msSaveAsFilters[".pptm"] = "Impress MS PowerPoint 2007 XML VBA"
	$msSaveAsFilters[".pptx"] = "Impress MS PowerPoint 2007 XML" ; Note these have a XML version too.
	$msSaveAsFilters[".odg"] = "impress8_draw"
	$msSaveAsFilters[".odp"] = "impress8"
	$msSaveAsFilters[".otp"] = "impress8_template"
	$msSaveAsFilters[".uop"] = "UOF presentation"

	If $bExportFilters Then
		$msExportFilters[".apng"] = "impress_png_Export"
		$msExportFilters[".bmp"] = "impress_bmp_Export"
		$msExportFilters[".emf"] = "impress_emf_Export"
		$msExportFilters[".eps"] = "impress_eps_Export"
		$msExportFilters[".gif"] = "impress_gif_Export"
		$msExportFilters[".htm"] = "impress_html_Export"
		$msExportFilters[".html"] = "impress_html_Export"
		$msExportFilters[".jfif"] = "impress_jpg_Export"
		$msExportFilters[".jif"] = "impress_jpg_Export"
		$msExportFilters[".jpg"] = "impress_jpg_Export"
		$msExportFilters[".jpeg"] = "impress_jpg_Export"
		$msExportFilters[".svg"] = "impress_svg_Export"
		$msExportFilters[".pdf"] = "impress_pdf_Export"
		$msExportFilters[".png"] = "impress_png_Export"
		$msExportFilters[".tif"] = "impress_tif_Export"
		$msExportFilters[".tiff"] = "impress_tif_Export"
		$msExportFilters[".webp"] = "writer_webp_Export"
		$msExportFilters[".wmf"] = "impress_wmf_Export"
		$msExportFilters[".xhtml"] = "XHTML Impress File"
	EndIf

	If StringInStr($sDocSavePath, "file:///") Then ;  If L.O. URl Then
		$iSlashLocation = StringInStr($sDocSavePath, "/", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, ($iLength - $iDotLocation) + 1)

	ElseIf StringInStr($sDocSavePath, "\") Then ;  Else if PC Path Then
		$iSlashLocation = StringInStr($sDocSavePath, "\", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, $iLength - $iDotLocation + 1)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	If $sFileExtension = $sDocSavePath Then ;  If no file extension identified, append .ods extension and return.
		$sDocSavePath = $sDocSavePath & ".odp"

		Return SetError($__LO_STATUS_SUCCESS, 3, "impress8")

	Else
		$sFileExtension = StringLower(StringStripWS($sFileExtension, $STR_STRIPALL))
	EndIf

	$sFilterName = $msSaveAsFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilterName)

	If $bExportFilters Then $sFilterName = $msExportFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilterName)

	$sDocSavePath = StringReplace($sDocSavePath, $sFileExtension, ".odp") ; If No results, replace with ODS extension.

	Return SetError($__LO_STATUS_SUCCESS, 3, "impress8")
EndFunc   ;==>__LOImpress_FilterNameGet

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_GetShapeName
; Description ...: Create a Shape Name that hasn't been used yet in the slide.
; Syntax ........: __LOImpress_GetShapeName(ByRef $oSlide, $sShapeName)
; Parameters ....: $oSlide              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $sShapeName          - a string value. The Shape name to begin with.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sShapeName not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Slide contained no shapes, returning the Shape name with a "1" appended.
;                  @Error 0 @Extended 1 Return String = Success. Returning the unique Shape name to use.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function adds a digit after the shape name, incrementing it until that name hasn't been used yet in L.O.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_GetShapeName(ByRef $oSlide, $sShapeName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oSlide.hasElements() Then
		Do ; Cycle through until I find a unique name.
			$iCount += 1
			For $i = 0 To $oSlide.getCount() - 1
				; Impress doesn't set the Shape name on new shapes. It has names in the UI that would correspond to the order of the shapes inserted, i.e. Shape 1, Shape 2. Etc.
				If ($oSlide.getByIndex($i).Name() = $sShapeName & $iCount) Or (($oSlide.getByIndex($i).Name() = "") And (("Shape " & ($i + 1)) = $sShapeName & $iCount)) Then ExitLoop

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
			Next
		Until $i = $oSlide.getCount()

	Else

		Return SetError($__LO_STATUS_SUCCESS, 0, $sShapeName & "1") ; If Doc has no shapes, just return the name with a "1" appended.
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 1, $sShapeName & $iCount)
EndFunc   ;==>__LOImpress_GetShapeName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_GradientIsModified
; Description ...: Check whether a pre-set gradient has been modified from its default values.
; Syntax ........: __LOImpress_GradientIsModified($tGradient, $sGradientName)
; Parameters ....: $tGradient           - a dll struct value. A Gradient Structure to compare property values with.
;                  $sGradientName       - a string value. The Gradient's current name.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $tGradient not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sGradientName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current color stop array.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the Gradient has been modified, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_GradientIsModified($tGradient, $sGradientName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStopColor
	Local $atColorStop[0]

	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch $sGradientName
		Case $LOI_GRAD_NAME_PASTEL_BOUQUET
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_LINEAR) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 0) Or _
						(.YOffset() <> 0) Or _
						(.Angle() <> 300) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> 14543051) Or _
						(.EndColor() <> 16766935) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_PASTEL_DREAM
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_RECT) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 50) Or _
						(.YOffset() <> 50) Or _
						(.Angle() <> 450) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> 16766935) Or _
						(.EndColor() <> 11847644) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_BLUE_TOUCH
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_LINEAR) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 0) Or _
						(.YOffset() <> 0) Or _
						(.Angle() <> 10) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> 11847644) Or _
						(.EndColor() <> 14608111) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_BLANK_W_GRAY
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_LINEAR) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 0) Or _
						(.YOffset() <> 0) Or _
						(.Angle() <> 900) Or _
						(.Border() <> 75) Or _
						(.StartColor() <> $LO_COLOR_WHITE) Or _
						(.EndColor() <> 14540253) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_LONDON_MIST
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_LINEAR) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 0) Or _
						(.YOffset() <> 0) Or _
						(.Angle() <> 300) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> 13421772) Or _
						(.EndColor() <> 6710886) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUBMARINE
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_LINEAR) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 0) Or _
						(.YOffset() <> 0) Or _
						(.Angle() <> 0) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> 14543051) Or _
						(.EndColor() <> 11847644) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_MIDNIGHT
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_LINEAR) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 0) Or _
						(.YOffset() <> 0) Or _
						(.Angle() <> 0) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> $LO_COLOR_BLACK) Or _
						(.EndColor() <> 2777241) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_DEEP_OCEAN
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_RADIAL) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 50) Or _
						(.YOffset() <> 50) Or _
						(.Angle() <> 0) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> $LO_COLOR_BLACK) Or _
						(.EndColor() <> 7512015) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_MAHOGANY
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_SQUARE) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 50) Or _
						(.YOffset() <> 50) Or _
						(.Angle() <> 450) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> $LO_COLOR_BLACK) Or _
						(.EndColor() <> 9250846) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_GREEN_GRASS
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_LINEAR) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 0) Or _
						(.YOffset() <> 0) Or _
						(.Angle() <> 300) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> 16776960) Or _
						(.EndColor() <> 8508442) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_NEON_LIGHT
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_ELLIPTICAL) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 50) Or _
						(.YOffset() <> 50) Or _
						(.Angle() <> 0) Or _
						(.Border() <> 15) Or _
						(.StartColor() <> 1209890) Or _
						(.EndColor() <> $LO_COLOR_WHITE) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNSHINE
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_RADIAL) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 66) Or _
						(.YOffset() <> 33) Or _
						(.Angle() <> 0) Or _
						(.Border() <> 33) Or _
						(.StartColor() <> 16760576) Or _
						(.EndColor() <> 16776960) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_RAINBOW
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_RADIAL) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 50) Or _
						(.YOffset() <> 100) Or _
						(.Angle() <> 0) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> $LO_COLOR_WHITE) Or _
						(.EndColor() <> $LO_COLOR_WHITE) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 7) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0.2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor()
					If _
							($tStopColor.Red() <> 1) Or _
							($tStopColor.Green() <> 0) Or _
							($tStopColor.Blue() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset() <> 0.2) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[2].StopColor()
					If _
							($tStopColor.Red() <> 1) Or _
							($tStopColor.Green() <> 1) Or _
							($tStopColor.Blue() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[2].StopOffset() <> 0.4) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[3].StopColor()
					If _
							($tStopColor.Red() <> 0) Or _
							($tStopColor.Green() <> 1) Or _
							($tStopColor.Blue() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[3].StopOffset() <> 0.5) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[4].StopColor()
					If _
							($tStopColor.Red() <> 0) Or _
							($tStopColor.Green() <> 1) Or _
							($tStopColor.Blue() <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[4].StopOffset() <> 0.65) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[5].StopColor()
					If _
							($tStopColor.Red() <> 1) Or _
							($tStopColor.Green() <> 0) Or _
							($tStopColor.Blue() <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[5].StopOffset() <> 0.8) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[6].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[6].StopOffset <> 0.8) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNRISE
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_LINEAR) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 0) Or _
						(.YOffset() <> 0) Or _
						(.Angle() <> 0) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> 3713206) Or _
						(.EndColor() <> 14065797) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 4) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> 0.505882352941176) Or _
							($tStopColor.Green() <> 0.784313725490196) Or _
							($tStopColor.Blue() <> 0.768627450980392) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset() <> 0.5) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[2].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> 0.717647058823529) Or _
							($tStopColor.Green() <> 0.807843137254902) Or _
							($tStopColor.Blue() <> 0.698039215686275) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[2].StopOffset() <> 0.75) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[3].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[3].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNDOWN
			With $tGradient
				If _
						(.Style() <> $LOI_GRAD_TYPE_LINEAR) Or _
						(.StepCount() <> 0) Or _
						(.XOffset() <> 0) Or _
						(.YOffset() <> 0) Or _
						(.Angle() <> 0) Or _
						(.Border() <> 0) Or _
						(.StartColor() <> 985943) Or _
						(.EndColor() <> 16759664) Or _
						(.StartIntensity() <> 100) Or _
						(.EndIntensity() <> 100) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				If __LO_VersionCheck(7.6) Then
					$atColorStop = .ColorStops()
					If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					If (UBound($atColorStop) <> 5) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[0].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> (BitAND(BitShift(.StartColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green() <> (BitAND(BitShift(.StartColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue() <> (BitAND(.StartColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[0].StopOffset() <> 0) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[1].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> 0.392156862745098) Or _
							($tStopColor.Green() <> 0.305882352941177) Or _
							($tStopColor.Blue() <> 0.690196078431373) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[1].StopOffset() <> 0.3) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[2].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> 0.827450980392157) Or _
							($tStopColor.Green() <> 0.572549019607843) Or _
							($tStopColor.Blue() <> 0.83921568627451) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[2].StopOffset() <> 0.5) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[3].StopColor() ; "Start Color" Value.
					If _
							($tStopColor.Red() <> 0.996078431372549) Or _
							($tStopColor.Green() <> 0.733333333333333) Or _
							($tStopColor.Blue() <> 0.76078431372549) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[3].StopOffset() <> 0.75) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					$tStopColor = $atColorStop[4].StopColor() ; Last element is "End Color" Value.
					If _
							($tStopColor.Red <> (BitAND(BitShift(.EndColor(), 16), 0xff) / 255)) Or _
							($tStopColor.Green <> (BitAND(BitShift(.EndColor(), 8), 0xff) / 255)) Or _
							($tStopColor.Blue <> (BitAND(.EndColor(), 0xff) / 255)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

					If ($atColorStop[4].StopOffset <> 1) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
				EndIf
			EndWith
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>__LOImpress_GradientIsModified

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_GradientNameInsert
; Description ...: Create and insert a new Gradient name.
; Syntax ........: __LOImpress_GradientNameInsert(ByRef $oDoc, $tGradient[, $sGradientName = "Gradient "])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $tGradient           - a dll struct value. A Gradient Structure to copy settings from.
;                  $sGradientName       - [optional] a string value. Default is "Gradient ". The Gradient name to create.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tGradient not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sGradientName not a string.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.GradientTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" or "com.sun.star.awt.Gradient2" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error creating Gradient Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. A new Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Gradient name is blank, I need to create a new name and apply it. I think I could re-use an old one without problems, but I'm not sure, so to be safe, I will create a new one.
;                  If there are no names that have been already created, then I need to create and apply one before the gradient will be displayed.
;                  Else if a preset Gradient is called, I need to create its name before it can be used.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_GradientNameInsert(ByRef $oDoc, $tGradient, $sGradientName = "Gradient ")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tNewGradient
	Local $oGradTable
	Local $iCount = 1
	Local $sGradient = "com.sun.star.awt.Gradient2"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If Not __LO_VersionCheck(7.6) Then $sGradient = "com.sun.star.awt.Gradient"

	$oGradTable = $oDoc.createInstance("com.sun.star.drawing.GradientTable")
	If Not IsObj($oGradTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sGradientName = "Gradient ") Then
		While $oGradTable.hasByName($sGradientName & $iCount)
			$iCount += 1
			Sleep((IsInt($iCount / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		WEnd
		$sGradientName = $sGradientName & $iCount
	EndIf

	$tNewGradient = __LO_CreateStruct($sGradient)
	If Not IsObj($tNewGradient) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; Copy the settings over from the input Style Gradient to my new one. This may not be necessary? But just in case.
	With $tNewGradient
		.Style = $tGradient.Style()
		.XOffset = $tGradient.XOffset()
		.YOffset = $tGradient.YOffset()
		.Angle = $tGradient.Angle()
		.Border = $tGradient.Border()
		.StartColor = $tGradient.StartColor()
		.EndColor = $tGradient.EndColor()
		.StartIntensity = $tGradient.StartIntensity()
		.EndIntensity = $tGradient.EndIntensity()

		If __LO_VersionCheck(7.6) Then .ColorStops = $tGradient.ColorStops()

	EndWith

	If Not $oGradTable.hasByName($sGradientName) Then
		$oGradTable.insertByName($sGradientName, $tNewGradient)
		If Not ($oGradTable.hasByName($sGradientName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $sGradientName)
EndFunc   ;==>__LOImpress_GradientNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_GradientPresets
; Description ...: Set Page background Gradient to preset settings.
; Syntax ........: __LOImpress_GradientPresets(ByRef $oDoc, ByRef $oObject, ByRef $tGradient, $sGradientName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oObject             - [in/out] an object. The Object to modify the Gradient settings for.
;                  $tGradient           - [in/out] an object. The Fill Gradient Object to modify the Gradient settings for.
;                  $sGradientName       - a string value. The Gradient Preset name to apply.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.awt.ColorStop" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to create Gradient name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. The Style Gradient settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_GradientPresets(ByRef $oDoc, ByRef $oObject, ByRef $tGradient, $sGradientName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tColorStop, $tStopColor
	Local $atColorStop[2]

	If __LO_VersionCheck(7.6) Then
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$atColorStop[0] = $tColorStop

		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$atColorStop[1] = $tColorStop
	EndIf

	Switch $sGradientName
		Case $LOI_GRAD_NAME_PASTEL_BOUQUET
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 14543051
				.EndColor = 16766935
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_PASTEL_DREAM
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RECT
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 450
				.Border = 0
				.StartColor = 16766935
				.EndColor = 11847644
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_BLUE_TOUCH
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 10
				.Border = 0
				.StartColor = 11847644
				.EndColor = 14608111
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_BLANK_W_GRAY
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 900
				.Border = 75
				.StartColor = $LO_COLOR_WHITE
				.EndColor = 14540253
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_LONDON_MIST
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 13421772
				.EndColor = 6710886
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUBMARINE
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 14543051
				.EndColor = 11847644
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_MIDNIGHT
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = $LO_COLOR_BLACK
				.EndColor = 2777241
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_DEEP_OCEAN
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 0
				.StartColor = $LO_COLOR_BLACK
				.EndColor = 7512015
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_MAHOGANY
			With $tGradient
				.Style = $LOI_GRAD_TYPE_SQUARE
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 450
				.Border = 0
				.StartColor = $LO_COLOR_BLACK
				.EndColor = 9250846
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_GREEN_GRASS
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 16776960
				.EndColor = 8508442
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_NEON_LIGHT
			With $tGradient
				.Style = $LOI_GRAD_TYPE_ELLIPTICAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 15
				.StartColor = 1209890
				.EndColor = $LO_COLOR_WHITE
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNSHINE
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 66
				.YOffset = 33
				.Angle = 0
				.Border = 33
				.StartColor = 16760576
				.EndColor = 16776960
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_RAINBOW
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 100
				.Angle = 0
				.Border = 0
				.StartColor = $LO_COLOR_WHITE
				.EndColor = $LO_COLOR_WHITE
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					ReDim $atColorStop[7]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0.2
					$atColorStop[0] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.2

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 0
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.4

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 1
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0
					$tStopColor.Green = 1
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[3] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.65

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0
					$tStopColor.Green = 1
					$tStopColor.Blue = 1
					$tColorStop.StopColor = $tStopColor

					$atColorStop[4] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.8

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 0
					$tStopColor.Blue = 1
					$tColorStop.StopColor = $tStopColor

					$atColorStop[5] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.8
					$atColorStop[6] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNRISE
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 3713206
				.EndColor = 14065797
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					ReDim $atColorStop[4]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.505882352941176
					$tStopColor.Green = 0.784313725490196
					$tStopColor.Blue = 0.768627450980392
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.75

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.717647058823529
					$tStopColor.Green = 0.807843137254902
					$tStopColor.Blue = 0.698039215686275
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 1
					$atColorStop[3] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNDOWN
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 985943
				.EndColor = 16759664
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					ReDim $atColorStop[5]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.3

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.392156862745098
					$tStopColor.Green = 0.305882352941177
					$tStopColor.Blue = 0.690196078431373
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.827450980392157
					$tStopColor.Green = 0.572549019607843
					$tStopColor.Blue = 0.83921568627451
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.75

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.996078431372549
					$tStopColor.Green = 0.733333333333333
					$tStopColor.Blue = 0.76078431372549
					$tColorStop.StopColor = $tStopColor

					$atColorStop[3] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 1
					$atColorStop[4] = $tColorStop
				EndIf
			EndWith

		Case Else ; Custom Gradient Name
			__LOImpress_GradientNameInsert($oDoc, $tGradient, $sGradientName)
			If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oObject.FillGradientName = $sGradientName

			Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndSwitch

	If __LO_VersionCheck(7.6) Then
		$tColorStop = $atColorStop[0] ; "Start Color" Value.

		$tStopColor = $tColorStop.StopColor()

		$tStopColor.Red = (BitAND(BitShift($tGradient.StartColor(), 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($tGradient.StartColor(), 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($tGradient.StartColor(), 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStop[0] = $tColorStop

		$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; Last element is "End Color" Value.

		$tStopColor = $tColorStop.StopColor()

		$tStopColor.Red = (BitAND(BitShift($tGradient.EndColor(), 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($tGradient.EndColor(), 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($tGradient.EndColor(), 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStop[UBound($atColorStop) - 1] = $tColorStop

		$tGradient.ColorStops = $atColorStop
	EndIf

	__LOImpress_GradientNameInsert($oDoc, $tGradient, $sGradientName)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oObject.FillGradient = $tGradient
	$oObject.FillGradientName = $sGradientName
	$oObject.FillGradientStepCount = $tGradient.StepCount()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_GradientPresets

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LOImpress_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LOImpress_ComError_UserFunction(Default)
	Local $vUserFunction, $avUserParams[2] = ["CallArgArray", $oComError]

	If IsArray($avUserFunction) Then
		$vUserFunction = $avUserFunction[0]
		ReDim $avUserParams[UBound($avUserFunction) + 1]
		For $i = 1 To UBound($avUserFunction) - 1
			$avUserParams[$i + 1] = $avUserFunction[$i]
		Next

	Else
		$vUserFunction = $avUserFunction
	EndIf
	If IsFunc($vUserFunction) Then
		Switch $vUserFunction
			Case ConsoleWrite
				ConsoleWrite("!--COM Error-Begin--" & @CRLF & _
						"Module: LibreOffice Impress" & @CRLF & _
						"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline & @CRLF & _
						"!--COM-Error-End--" & @CRLF)

			Case MsgBox
				MsgBox(0, "COM Error", "Module: LibreOffice Impress" & @CRLF & _
						"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline)

			Case Else
				Call($vUserFunction, $avUserParams)
		EndSwitch
	EndIf
EndFunc   ;==>__LOImpress_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_NumRuleCreateMap
; Description ...: Creates a map with values for each setting location in the Array.
; Syntax ........: __LOImpress_NumRuleCreateMap(ByRef $atNumLevel)
; Parameters ....: $atNumLevel          - [in/out] an array of dll structs. An Array of Property Structures for a Numbering Rule.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $atNumLevel not an Array.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Returning a Map containing the location in the array for each setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_NumRuleCreateMap(ByRef $atNumLevel)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $mNumLevel[]

	If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	For $i = 0 To UBound($atNumLevel) - 1
		$mNumLevel[$atNumLevel[$i].Name()] = $i
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, $mNumLevel)
EndFunc   ;==>__LOImpress_NumRuleCreateMap

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ParAlignment
; Description ...: Set and Retrieve Paragraph Alignment settings.
; Syntax ........: __LOImpress_ParAlignment(ByRef $oObj[, $iHorAlign = Null[, $iLastLineAlign = Null[, $iTxtDirection = Null]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iHorAlign           - [optional] an integer value (0-3). Default is Null. The Horizontal alignment of the paragraph. See Constants, $LOI_PAR_ALIGN_HOR_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iLastLineAlign      - [optional] an integer value (0-3). Default is Null. Specify the alignment for the last line in the paragraph. See Constants, $LOI_PAR_LAST_LINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iTxtDirection       - [optional] an integer value (0-5). Default is Null. The Text Writing Direction. See Constants, $LOI_PAR_TXT_DIR_* as defined in LibreOfficeImpress_Constants.au3. [LibreOffice Default is 4]
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iHorAlign not an Integer, less than 0 or greater than 3. See Constants, $LOI_PAR_ALIGN_HOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iLastLineAlign not an Integer, less than 0 or greater than 3. See Constants, $LOI_PAR_LAST_LINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iTxtDirection not an Integer, less than 0 or greater than 5. See Constants, $LOI_PAR_TXT_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iHorAlign
;                  |                               2 = Error setting $iLastLineALign
;                  |                               4 = Error setting $iTxtDirection
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iHorAlign must be set to $LOI_PAR_ALIGN_HOR_JUSTIFIED(2) before you can set $iLastLineAlign.
;                  $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Expand single word, Snap to grid, and Vertical align (Text-To-Text), seem to be unavailable in the API, and do not seem to work in LibreOffice.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ParAlignment(ByRef $oObj, $iHorAlign = Null, $iLastLineAlign = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avAlignment[3]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iHorAlign, $iLastLineAlign, $iTxtDirection) Then
		__LO_ArrayFill($avAlignment, $oObj.ParaAdjust(), $oObj.ParaLastLineAdjust(), $oObj.WritingMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avAlignment)
	EndIf

	If ($iHorAlign <> Null) Then
		If Not __LO_IntIsBetween($iHorAlign, $LOI_PAR_ALIGN_HOR_LEFT, $LOI_PAR_ALIGN_HOR_CENTER) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.ParaAdjust = $iHorAlign
		$iError = ($oObj.ParaAdjust() = $iHorAlign) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iLastLineAlign <> Null) Then
		If Not __LO_IntIsBetween($iLastLineAlign, $LOI_PAR_LAST_LINE_JUSTIFIED, $LOI_PAR_LAST_LINE_CENTER, "", $LOI_PAR_LAST_LINE_START) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.ParaLastLineAdjust = $iLastLineAlign
		$iError = ($oObj.ParaLastLineAdjust() = $iLastLineAlign) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDirection <> Null) Then
		If Not __LO_IntIsBetween($iTxtDirection, $LOI_PAR_TXT_DIR_LR_TB, $LOI_PAR_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.WritingMode = $iTxtDirection
		$iError = ($oObj.WritingMode() = $iTxtDirection) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ParAlignment

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ParIndent
; Description ...: Set or Retrieve Paragraph Indent settings.
; Syntax ........: __LOImpress_ParIndent(ByRef $oObj[, $iBeforeTxt = Null[, $iAfterTxt = Null[, $iFirstLine = Null]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iBeforeTxt          - [optional] an integer value (0-1162202). Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM).
;                  $iAfterTxt           - [optional] an integer value (0-1162202). Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM)
;                  $iFirstLine          - [optional] an integer value (0-1162202). Default is Null. Indentation distance of the first line of a paragraph. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iBeforeText not an Integer, less than 0 or greater than 1162202.
;                  @Error 1 @Extended 3 Return 0 = $iAfterText not an Integer, less than 0 or greater than 1162202.
;                  @Error 1 @Extended 4 Return 0 = $iFirstLine not an Integer, less than 0 or greater than 1162202.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBeforeTxt
;                  |                               2 = Error setting $iAfterTxt
;                  |                               4 = Error setting $iFirstLine
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Auto indent first line does not seem to work in LibreOffice, and seems to be not available in the API.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ParIndent(ByRef $oObj, $iBeforeTxt = Null, $iAfterTxt = Null, $iFirstLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avIndent[3]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iBeforeTxt, $iAfterTxt, $iFirstLine) Then
		__LO_ArrayFill($avIndent, $oObj.ParaLeftMargin(), $oObj.ParaRightMargin(), $oObj.ParaFirstLineIndent())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avIndent)
	EndIf

	If ($iBeforeTxt <> Null) Then
		If Not __LO_IntIsBetween($iBeforeTxt, 0, 1162202) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.ParaLeftMargin = $iBeforeTxt
		$iError = (__LO_IntIsBetween(($oObj.ParaLeftMargin()), ($iBeforeTxt - 1), ($iBeforeTxt + 1))) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iAfterTxt <> Null) Then
		If Not __LO_IntIsBetween($iAfterTxt, 0, 1162202) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.ParaRightMargin = $iAfterTxt
		$iError = (__LO_IntIsBetween(($oObj.ParaRightMargin()), ($iAfterTxt - 1), ($iAfterTxt + 1))) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iFirstLine <> Null) Then
		If Not __LO_IntIsBetween($iFirstLine, 0, 1162202) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.ParaFirstLineIndent = $iFirstLine
		$iError = (__LO_IntIsBetween(($oObj.ParaFirstLineIndent()), ($iFirstLine - 1), ($iFirstLine + 1))) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ParIndent

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ParSpacing
; Description ...: Set and Retrieve Line Spacing settings.
; Syntax ........: __LOImpress_ParSpacing(ByRef $oObj[, $iAbovePar = Null[, $iBelowPar = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null]]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iAbovePar           - [optional] an integer value (0-100000). Default is Null. The Space above a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iBelowPar           - [optional] an integer value (0-100000). Default is Null. The Space Below a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iLineSpcMode        - [optional] an integer value (0-3). Default is Null. The line spacing type of the paragraph. See Constants, $LOI_PAR_LINE_SPC_MODE_* as defined in LibreOfficeImpress_Constants.au3, also notice min and max values for each.
;                  $iLineSpcHeight      - [optional] an integer value. Default is Null. This value specifies the height in regard to Mode. See Remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iAbovePar not an Integer, less than 0 or greater than 100000.
;                  @Error 1 @Extended 3 Return 0 = $iBelowPar not an Integer, less than 0 or greater than 100000.
;                  @Error 1 @Extended 4 Return 0 = $iLineSpcMode not an Integer, less than 0 or greater than 3. See Constants, $LOI_PAR_LINE_SPC_MODE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iLineSpcHeight not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iLineSpcMode set to 0(Proportional) and $iLineSpcHeight less than 6(%) or greater than 65535(%).
;                  @Error 1 @Extended 7 Return 0 = $iLineSpcMode set to 1 or 2(Minimum, or Leading) and $iLineSpcHeight less than 0 or greater than 100000.
;                  @Error 1 @Extended 8 Return 0 = $iLineSpcMode set to 3(Fixed) and $iLineSpcHeight less than 51 or greater than 100000.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaLineSpacing Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAbovePar
;                  |                               2 = Error setting $iBelowPar
;                  |                               4 = Error setting $iLineSpcMode
;                  |                               8 = Error setting $iLineSpcHeight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
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
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ParSpacing(ByRef $oObj, $iAbovePar = Null, $iBelowPar = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tLine
	Local $iError = 0
	Local $avSpacing[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iAbovePar, $iBelowPar, $iLineSpcMode, $iLineSpcHeight) Then
		__LO_ArrayFill($avSpacing, $oObj.ParaTopMargin(), $oObj.ParaBottomMargin(), $oObj.ParaLineSpacing.Mode(), $oObj.ParaLineSpacing.Height())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iAbovePar <> Null) Then
		If Not __LO_IntIsBetween($iAbovePar, 0, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.ParaTopMargin = $iAbovePar
		$iError = (__LO_IntIsBetween(($oObj.ParaTopMargin()), ($iAbovePar - 1), ($iAbovePar + 1))) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iBelowPar <> Null) Then
		If Not __LO_IntIsBetween($iBelowPar, 0, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.ParaBottomMargin = $iBelowPar
		$iError = (__LO_IntIsBetween(($oObj.ParaBottomMargin()), ($iBelowPar - 1), ($iBelowPar + 1))) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iLineSpcMode <> Null) Then
		If Not __LO_IntIsBetween($iLineSpcMode, $LOI_PAR_LINE_SPC_MODE_PROP, $LOI_PAR_LINE_SPC_MODE_FIX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tLine = $oObj.ParaLineSpacing()
		If Not IsObj($tLine) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$tLine.Mode = $iLineSpcMode
		$oObj.ParaLineSpacing = $tLine
		$iError = ($oObj.ParaLineSpacing.Mode() = $iLineSpcMode) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLineSpcHeight <> Null) Then
		If Not IsInt($iLineSpcHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tLine = $oObj.ParaLineSpacing()
		If Not IsObj($tLine) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Switch $tLine.Mode()
			Case $LOI_PAR_LINE_SPC_MODE_PROP ; Proportional
				If Not __LO_IntIsBetween($iLineSpcHeight, 6, 65535) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; Min setting on Proportional is 6%

			Case $LOI_PAR_LINE_SPC_MODE_MIN, $LOI_PAR_LINE_SPC_MODE_LEADING ; Minimum and Leading Modes
				If Not __LO_IntIsBetween($iLineSpcHeight, 0, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			Case $LOI_PAR_LINE_SPC_MODE_FIX ; Fixed Line Spacing Mode
				If Not __LO_IntIsBetween($iLineSpcHeight, 51, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0) ; Min spacing is 51 when Fixed Mode
		EndSwitch
		$tLine.Height = $iLineSpcHeight
		$oObj.ParaLineSpacing = $tLine
		$iError = (__LO_IntIsBetween(($oObj.ParaLineSpacing.Height()), ($iLineSpcHeight - 1), ($iLineSpcHeight + 1))) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ParSpacing

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ParTabStopCreate
; Description ...: Create a new TabStop for a Paragraph.
; Syntax ........: __LOImpress_ParTabStopCreate(ByRef $oObj, $iPosition[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iPosition           - an integer value. The TabStop position to set the new TabStop to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] an integer value (0-4). The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] an integer value. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] an integer value. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iPosition not an Integer.
;                  @Error 1 @Extended 3 Return 0 = Tab Stop position called in $iPosition already exists in this Paragraph.
;                  @Error 1 @Extended 4 Return 0 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iDecChar not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iFillChar not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.style.TabStop" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaTabStops Array Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving list of TabStop Positions.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify the new Tabstop once inserted.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return Integer = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iAlignment
;                  |                               4 = Error setting $iDecChar
;                  |                               8 = Error setting $iFillChar
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Settings were successfully set. New TabStop position is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iPosition once set can vary +/- a Hundredth of a Millimeter (HMM). To ensure you can identify the tabstop to modify it again, This function returns the new TabStop position.
;                  Since $iPosition can fluctuate +/- a Hundredth of a Millimeter (HMM) when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;                  $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32. The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
;                  Call any optional parameter with Null keyword to skip it.
;                  $iNewTabStop position is still returned as even though some settings weren't successfully set, the new TabStop was still created.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ParTabStopCreate(ByRef $oObj, $iPosition, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aiTabList
	Local $bFound = False
	Local $iNewPosition = -1
	Local $atTabStops, $atNewTabStops
	Local $tFoundTabStop, $tTabStruct
	Local $iError = 0

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iPosition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If __LOImpress_CursorParHasTabStop($oObj, $iPosition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tTabStruct = __LO_CreateStruct("com.sun.star.style.TabStop")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tTabStruct.Position = $iPosition
	$tTabStruct.Alignment = 0
	$tTabStruct.DecimalChar = 0
	$tTabStruct.FillChar = 32 ; If set to 0 Libre sets fill character to Null instead of setting to None. 32 = None.(Space character)

	If ($iAlignment <> Null) Then
		If Not __LO_IntIsBetween($iAlignment, $LOI_PAR_TAB_ALIGN_LEFT, $LOI_PAR_TAB_ALIGN_DEFAULT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tTabStruct.Alignment = $iAlignment
	EndIf

	If ($iDecChar <> Null) Then
		If Not IsInt($iDecChar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tTabStruct.DecimalChar = $iDecChar
	EndIf

	If ($iFillChar <> Null) Then
		If Not IsInt($iFillChar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tTabStruct.FillChar = ($iFillChar = 0) ? (32) : ($iFillChar)
	EndIf

	__LO_AddTo1DArray($atTabStops, $tTabStruct)

	$aiTabList = __LOImpress_ParTabStopsGetList($oObj)     ; Get an array of existing tabstops to compare with
	If Not IsArray($aiTabList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	__LO_AddTo1DArray($aiTabList, 0)     ; Add a dummy to make Array sizes equal.

	$oObj.ParaTabStops = $atTabStops     ; Insert the new TabStop

	$atNewTabStops = $oObj.ParaTabStops()     ; Now retrieve a new list to find the final Tab Stop position.
	For $i = 0 To UBound($atNewTabStops) - 1
		If ($atNewTabStops[$i].Position()) <> $aiTabList[$i] Then
			$iNewPosition = $atNewTabStops[$i].Position()
			$tFoundTabStop = $atNewTabStops[$i]
			$bFound = True
			ExitLoop
		EndIf
	Next

	If Not $bFound Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)     ; Didn't find the new TabStop

	$iError = (__LO_IntIsBetween(($tFoundTabStop.Position()), ($iPosition - 1), ($iPosition + 1))) ? ($iError) : (BitOR($iError, 1))
	$iError = (__LO_VarsAreNull($iAlignment)) ? ($iError) : (($tFoundTabStop.Alignment = $iAlignment) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iDecChar)) ? ($iError) : (($tFoundTabStop.DecimalChar = $iDecChar) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iFillChar)) ? ($iError) : (($tFoundTabStop.FillChar = $iFillChar) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $iNewPosition) : SetError($__LO_STATUS_SUCCESS, 0, $iNewPosition)
EndFunc   ;==>__LOImpress_ParTabStopCreate

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ParTabStopDelete
; Description ...: Delete a TabStop from a Paragraph
; Syntax ........: __LOImpress_ParTabStopDelete(ByRef $oObj, $iTabStop)
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTabStop not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iTabStop not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Returning True if TabStop was successfully deleted, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per paragraph.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ParTabStopDelete(ByRef $oObj, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atOldTabStops[0]
	Local $bDeleted = False
	Local $iCount = 0

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOImpress_CursorParHasTabStop($oObj, $iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$atOldTabStops = $oObj.ParaTabStops()
	If Not IsArray($atOldTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($atOldTabStops) - 1
		If ($atOldTabStops[$i].Position() = $iTabStop) Then
			$bDeleted = True

		Else
			$atOldTabStops[$iCount] = $atOldTabStops[$i]
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	ReDim $atOldTabStops[$iCount]

	$oObj.ParaTabStops = $atOldTabStops

	Return SetError($__LO_STATUS_SUCCESS, 0, $bDeleted)
EndFunc   ;==>__LOImpress_ParTabStopDelete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop.
; Syntax ........: __LOImpress_ParTabStopMod(ByRef $oObj, $iTabStop[, $iPosition = Null[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]]])
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - [optional] an integer value. Default is Null. The New position to set the input position to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] an integer value (0-4). Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] an integer value. Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] an integer value. Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTabStop not an Integer.
;                  @Error 1 @Extended 3 Return 0 = TabStop called in $iTabStop not found.
;                  @Error 1 @Extended 4 Return 0 = $iPosition not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iDecChar not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iFillChar not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Requested TabStop Object.
;                  @Error 3 @Extended 3 Return 0 = Paragraph already contains a TabStop at the length/Position specified in $iPosition.
;                  @Error 3 @Extended 4 Return 0 = Error retrieving list of TabStop Positions.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iAlignment
;                  |                               4 = Error setting $iDecChar
;                  |                               8 = Error setting $iFillChar
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error 0 @Extended ? Return 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
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
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ParTabStopMod(ByRef $oObj, $iTabStop, $iPosition = Null, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atTabStops, $atNewTabStops
	Local $iError = 0, $iNewPosition = 0
	Local $tTabStruct
	Local $bNewPosition = False
	Local $aiTabList
	Local $aiTabSettings[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOImpress_CursorParHasTabStop($oObj, $iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($atTabStops) - 1
		If ($atTabStops[$i].Position() = $iTabStop) Then $tTabStruct = $atTabStops[$i]
		If IsObj($tTabStruct) Then ExitLoop
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next
	If Not IsObj($tTabStruct) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iPosition, $iAlignment, $iDecChar, $iFillChar) Then
		__LO_ArrayFill($aiTabSettings, $tTabStruct.Position(), $tTabStruct.Alignment(), $tTabStruct.DecimalChar(), $tTabStruct.FillChar())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTabSettings)
	EndIf

	If ($iPosition <> Null) Then
		If Not IsInt($iPosition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If __LOImpress_CursorParHasTabStop($oObj, $iPosition) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$tTabStruct.Position = $iPosition
		$iError = ($tTabStruct.Position() = $iPosition) ? ($iError) : (BitOR($iError, 1))
		$bNewPosition = True
	EndIf

	If ($iAlignment <> Null) Then
		If Not __LO_IntIsBetween($iAlignment, $LOI_PAR_TAB_ALIGN_LEFT, $LOI_PAR_TAB_ALIGN_DEFAULT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tTabStruct.Alignment = $iAlignment
		$iError = ($tTabStruct.Alignment = $iAlignment) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iDecChar <> Null) Then
		If Not IsInt($iDecChar) And ($iDecChar <> Null) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tTabStruct.DecimalChar = $iDecChar
		$iError = ($tTabStruct.DecimalChar = $iDecChar) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iFillChar <> Null) Then
		If Not IsInt($iFillChar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tTabStruct.FillChar = $iFillChar
		$tTabStruct.FillChar = ($tTabStruct.FillChar() = 0) ? (32) : ($tTabStruct.FillChar())
		$iError = ($tTabStruct.FillChar = $iFillChar) ? ($iError) : (BitOR($iError, 8))
	EndIf

	$atTabStops[$i] = $tTabStruct

	If $bNewPosition Then
		$aiTabList = __LOImpress_ParTabStopsGetList($oObj)
		If Not IsArray($aiTabList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oObj.ParaTabStops = $atTabStops

	If $bNewPosition Then
		$atNewTabStops = $oObj.ParaTabStops()
		For $j = 0 To UBound($atNewTabStops) - 1
			If ($atNewTabStops[$j].Position()) <> $aiTabList[$j] Then
				$iNewPosition = $atNewTabStops[$j].Position()
				ExitLoop
			EndIf
			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, $iNewPosition, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ParTabStopMod

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ParTabStopsGetList
; Description ...: Retrieve an array of TabStops available in a Paragraph.
; Syntax ........: __LOImpress_ParTabStopsGetList(ByRef $oObj)
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor, Shape, Shape Style or Presentation Style object returned by a previous  _LOImpress_ShapeCreateTextCursor, _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. An Array of TabStops. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ParTabStopsGetList(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atTabStops[0]
	Local $aiTabList[0]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $aiTabList[UBound($atTabStops)]

	For $i = 0 To UBound($atTabStops) - 1
		$aiTabList[$i] = $atTabStops[$i].Position()
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $i, $aiTabList)
EndFunc   ;==>__LOImpress_ParTabStopsGetList

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeAreaColor
; Description ...: Set or Retrieve the Fill color settings for a Shape, Shape Style or Presentation Style.
; Syntax ........: __LOImpress_ShapeAreaColor(ByRef $oObj[, $iColor = Null])
; Parameters ....: $oObj                - [in/out] an object. A Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is Null. The Fill color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current color value.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current Fill color as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeAreaColor(ByRef $oObj, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency, $iCurColor

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; If $iColor is Null, and Fill Style is set to solid, then return current color value, else return LO_COLOR_OFF.
	If __LO_VarsAreNull($iColor) Then
		If ($oObj.FillStyle() = $LOI_AREA_FILL_STYLE_SOLID) Then ; If FillStyle is set to solid, then return current color value, else return $LO_COLOR_OFF (Probably a Gradient is used or otherwise).
			$iCurColor = __LOImpress_ColorRemoveAlpha($oObj.FillColor())
			If Not IsInt($iCurColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Else
			$iCurColor = $LO_COLOR_OFF
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurColor)
	EndIf

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($iColor = $LO_COLOR_OFF) Then
		$oObj.FillStyle = $LOI_AREA_FILL_STYLE_OFF

	Else
		$iOldTransparency = $oObj.FillTransparence()
		If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oObj.FillStyle = $LOI_AREA_FILL_STYLE_SOLID
		$oObj.FillColor = $iColor
		$iError = ($oObj.FillColor() = $iColor) ? ($iError) : (BitOR($iError, 1))

		$oObj.FillTransparence = $iOldTransparency
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeAreaColor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeAreaGradientMulticolor
; Description ...: Set or Retrieve a Shape, Shape Style, or Presentation Style's Multicolor Gradient settings. See remarks.
; Syntax ........: __LOImpress_ShapeAreaGradientMulticolor(ByRef $oObj[, $avColorStops = Null])
; Parameters ....: $oObj                - [in/out] an object. A Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 4 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 5 Return ? = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve FillGradient Struct.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve StopColor Struct.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current version less than 7.6.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, as a RGB Color Integer.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......: _LO_GradientMulticolorAdd, _LO_GradientMulticolorDelete, _LO_GradientMulticolorModify, _LOImpress_ShapeAreaTransparencyGradientMulti
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeAreaGradientMulticolor(ByRef $oObj, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oObj.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($avColorStops) Then
		$atColorStops = $tStyleGradient.ColorStops()
		If Not IsArray($atColorStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		ReDim $avNewColorStops[UBound($atColorStops)][2]

		For $i = 0 To UBound($atColorStops) - 1
			$avNewColorStops[$i][0] = $atColorStops[$i].StopOffset()
			$tStopColor = $atColorStops[$i].StopColor()
			If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$avNewColorStops[$i][1] = Int(BitShift(($tStopColor.Red() * 255), -16) + BitShift(($tStopColor.Green() * 255), -8) + ($tStopColor.Blue() * 255)) ; RGB to Long
			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tStopColor.Red = (BitAND(BitShift($avColorStops[$i][1], 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($avColorStops[$i][1], 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($avColorStops[$i][1], 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oObj.FillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oObj.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeAreaGradientMulticolor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeAreaShadow
; Description ...: Set or Retrieve the shadow settings for a Shape, Shape Style, or Presentation Style.
; Syntax ........: __LOImpress_ShapeAreaShadow(ByRef $oObj[, $bShadow = Null[, $iLocation = Null[, $iColor = Null[, $iDistance = Null[, $iBlur = Null[, $iTransparency = Null]]]]]])
; Parameters ....: $oObj                - [in/out] an object. A Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $bShadow             - [optional] a boolean value. Default is Null. If True, a Shadow is present for the Shape.
;                  $iLocation           - [optional] an integer value (0-8). Default is Null. The Location of the Shadow, must be one of the Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Shadow color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iDistance           - [optional] an integer value. Default is Null. The distance of the Shadow from the Shape's edges, set in Hundredths of a Millimeter (HMM).
;                  $iBlur               - [optional] an integer value (0-150). Default is Null. The amount of blur applied to the Shadow, set in Printer's Points.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The percentage of Shadow transparency. 100% means completely transparent.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bShadow not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $iLocation not an Integer, less than 0 or greater than 8. See Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iDistance not an Integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iBlur not an Integer, less than 0 or greater than 150 Printer's Points.
;                  @Error 1 @Extended 7 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Distance and Location Values.
;                  @Error 3 @Extended 2 Return 0 = Failed to modify Location property.
;                  @Error 3 @Extended 3 Return 0 = Failed to modify Distance property.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bShadow
;                  |                               2 = Error setting $iLocation
;                  |                               4 = Error setting $iColor
;                  |                               8 = Error setting $iDistance
;                  |                               16 = Error setting $iBlur
;                  |                               32 = Error setting $iTransparency
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  LibreOffice may change the shadow distance +/- a Hundredth of a Millimeter (HMM).
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeAreaShadow(ByRef $oObj, $bShadow = Null, $iLocation = Null, $iColor = Null, $iDistance = Null, $iBlur = Null, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iInternalLocation, $iInternalDistance
	Local $avShadow[6]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bShadow, $iLocation, $iColor, $iDistance, $iBlur, $iTransparency) Then
		$iInternalDistance = __LOImpress_ShapeAreaShadowModify($oObj)
		$iInternalLocation = @extended
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LO_ArrayFill($avShadow, $oObj.Shadow(), $iInternalLocation, $oObj.ShadowColor(), $iInternalDistance, _
				_LO_UnitConvert($oObj.ShadowBlur(), $LO_CONVERT_UNIT_HMM_PT), _
				$oObj.ShadowTransparence())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($bShadow <> Null) Then
		If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.Shadow = $bShadow
		$iError = ($oObj.Shadow() = $bShadow) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOI_SHAPE_SHADOW_LOCATION_TOP_LEFT, $LOI_SHAPE_SHADOW_LOCATION_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		__LOImpress_ShapeAreaShadowModify($oObj, $iLocation)
		If (@error = $__LO_STATUS_PROP_SETTING_ERROR) Then
			$iError = BitOR($iError, 2)

		ElseIf @error Then

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndIf
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.ShadowColor = $iColor
		$iError = ($oObj.ShadowColor() = $iColor) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iDistance <> Null) Then
		If Not __LO_IntIsBetween($iDistance, 0, $iDistance) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		__LOImpress_ShapeAreaShadowModify($oObj, Null, $iDistance)
		If (@error = $__LO_STATUS_PROP_SETTING_ERROR) Then
			$iError = BitOR($iError, 8)

		ElseIf @error Then

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		EndIf
	EndIf

	If ($iBlur <> Null) Then
		If Not __LO_IntIsBetween($iBlur, 0, 150) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; 0 - 5292 max Hundredths of a Millimeter (HMM).

		$oObj.ShadowBlur = _LO_UnitConvert($iBlur, $LO_CONVERT_UNIT_PT_HMM)
		$iError = ($oObj.ShadowBlur() = _LO_UnitConvert($iBlur, $LO_CONVERT_UNIT_PT_HMM)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.ShadowTransparence = $iTransparency
		$iError = ($oObj.ShadowTransparence = $iTransparency) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeAreaShadow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeAreaShadowModify
; Description ...: Internal function for setting or retrieving Shape Shadow Location and Distance settings.
; Syntax ........: __LOImpress_ShapeAreaShadowModify($oShape[, $iLocation = Null[, $iDistance = Null]])
; Parameters ....: $oShape              - an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iLocation           - [optional] an integer value (0-8). Default is Null. The Location of the Shadow, must be one of the Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDistance           - [optional] an integer value. Default is Null. The distance of the Shadow from the Shape's edges, set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLocation
;                  |                               2 = Error setting $iDistance
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully set the settings.
;                  @Error 0 @Extended ? Return Integer = Success. $iLocation and $iDistance called with Null, returning current Values. Return will be current distance, and @Extended will be the current Location.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeAreaShadowModify($oShape, $iLocation = Null, $iDistance = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn = False, $bModifyLocation = True
	Local $iError = 1

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLocation, $iDistance) Then $bReturn = True

	If ($iLocation = Null) Then ; Determine current location)
		$bModifyLocation = False
		$iError = 2
		Select
			Case (($oShape.ShadowXDistance() < 0) And ($oShape.ShadowYDistance() < 0)) ; Top Left.
				$iLocation = $LOI_SHAPE_SHADOW_LOCATION_TOP_LEFT

			Case (($oShape.ShadowXDistance() = 0) And ($oShape.ShadowYDistance() < 0)) ; Top Center
				$iLocation = $LOI_SHAPE_SHADOW_LOCATION_TOP_CENTER

			Case (($oShape.ShadowXDistance() > 0) And ($oShape.ShadowYDistance() < 0)) ; Top Right
				$iLocation = $LOI_SHAPE_SHADOW_LOCATION_TOP_RIGHT

			Case (($oShape.ShadowXDistance() < 0) And ($oShape.ShadowYDistance() = 0)) ; Middle Left
				$iLocation = $LOI_SHAPE_SHADOW_LOCATION_MIDDLE_LEFT

			Case (($oShape.ShadowXDistance() = 0) And ($oShape.ShadowYDistance() = 0)) ; Middle Center
				$iLocation = $LOI_SHAPE_SHADOW_LOCATION_MIDDLE_CENTER

			Case (($oShape.ShadowXDistance() > 0) And ($oShape.ShadowYDistance() = 0)) ; Middle Right
				$iLocation = $LOI_SHAPE_SHADOW_LOCATION_MIDDLE_RIGHT

			Case (($oShape.ShadowXDistance() < 0) And ($oShape.ShadowYDistance() > 0)) ; Bottom Left
				$iLocation = $LOI_SHAPE_SHADOW_LOCATION_BOTTOM_LEFT

			Case (($oShape.ShadowXDistance() = 0) And ($oShape.ShadowYDistance() > 0)) ; Bottom Center
				$iLocation = $LOI_SHAPE_SHADOW_LOCATION_BOTTOM_CENTER

			Case (($oShape.ShadowXDistance() > 0) And ($oShape.ShadowYDistance() > 0)) ; Bottom Right
				$iLocation = $LOI_SHAPE_SHADOW_LOCATION_BOTTOM_RIGHT
		EndSelect
	EndIf

	If ($iDistance = Null) Then
		; Retrieve the current Distance setting
		If ($oShape.ShadowXDistance() <> 0) Then
			$iDistance = $oShape.ShadowXDistance()

		ElseIf ($oShape.ShadowYDistance() <> 0) Then
			$iDistance = $oShape.ShadowYDistance()

		Else
			$iDistance = 0
		EndIf

		If $bModifyLocation And ($iDistance = 0) Then $iDistance = 100 ; Set a non 0 value so location can be set.

		; If negative, make it positive for easier processing.
		$iDistance = ($iDistance < 0) ? ($iDistance * -1) : ($iDistance)
	EndIf

	If $bReturn Then Return SetError($__LO_STATUS_SUCCESS, $iLocation, $iDistance)

	Switch $iLocation
		Case $LOI_SHAPE_SHADOW_LOCATION_TOP_LEFT
			$oShape.ShadowXDistance = ($iDistance * -1)
			$oShape.ShadowYDistance = ($iDistance * -1)

			Return (($oShape.ShadowXDistance() = ($iDistance * -1)) And ($oShape.ShadowYDistance() = ($iDistance * -1))) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOI_SHAPE_SHADOW_LOCATION_TOP_CENTER
			$oShape.ShadowXDistance = 0
			$oShape.ShadowYDistance = ($iDistance * -1)

			Return (($oShape.ShadowXDistance() = 0) And ($oShape.ShadowYDistance() = ($iDistance * -1))) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOI_SHAPE_SHADOW_LOCATION_TOP_RIGHT
			$oShape.ShadowXDistance = $iDistance
			$oShape.ShadowYDistance = ($iDistance * -1)

			Return (($oShape.ShadowXDistance() = $iDistance) And ($oShape.ShadowYDistance() = ($iDistance * -1))) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOI_SHAPE_SHADOW_LOCATION_MIDDLE_LEFT
			$oShape.ShadowXDistance = ($iDistance * -1)
			$oShape.ShadowYDistance = 0

			Return (($oShape.ShadowXDistance() = ($iDistance * -1)) And ($oShape.ShadowYDistance() = 0)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOI_SHAPE_SHADOW_LOCATION_MIDDLE_CENTER
			$oShape.ShadowXDistance = ($bModifyLocation) ? (0) : ($iDistance)
			$oShape.ShadowYDistance = ($bModifyLocation) ? (0) : ($iDistance)

			Return (($oShape.ShadowXDistance() = (($bModifyLocation) ? (0) : ($iDistance))) And ($oShape.ShadowYDistance() = (($bModifyLocation) ? (0) : ($iDistance)))) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOI_SHAPE_SHADOW_LOCATION_MIDDLE_RIGHT
			$oShape.ShadowXDistance = $iDistance
			$oShape.ShadowYDistance = 0

			Return (($oShape.ShadowXDistance() = $iDistance) And ($oShape.ShadowYDistance() = 0)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOI_SHAPE_SHADOW_LOCATION_BOTTOM_LEFT
			$oShape.ShadowXDistance = ($iDistance * -1)
			$oShape.ShadowYDistance = $iDistance

			Return (($oShape.ShadowXDistance() = ($iDistance * -1)) And ($oShape.ShadowYDistance() = $iDistance)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOI_SHAPE_SHADOW_LOCATION_BOTTOM_CENTER
			$oShape.ShadowXDistance = 0
			$oShape.ShadowYDistance = $iDistance

			Return (($oShape.ShadowXDistance() = 0) And ($oShape.ShadowYDistance() = $iDistance)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOI_SHAPE_SHADOW_LOCATION_BOTTOM_RIGHT
			$oShape.ShadowXDistance = $iDistance
			$oShape.ShadowYDistance = $iDistance

			Return (($oShape.ShadowXDistance() = $iDistance) And ($oShape.ShadowYDistance() = $iDistance)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
	EndSwitch
EndFunc   ;==>__LOImpress_ShapeAreaShadowModify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeAreaTransparency
; Description ...: Set or retrieve Transparency settings for a Shape, Shape Style or Presentation Style.
; Syntax ........: __LOImpress_ShapeAreaTransparency(ByRef $oObj[, $iTransparency = Null])
; Parameters ....: $oObj                - [in/out] an object. A Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeAreaTransparency(ByRef $oObj, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurTransp

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransparency) Then
		$iCurTransp = $oObj.FillTransparence()
		If Not IsInt($iCurTransp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurTransp)
	EndIf

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oObj.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oObj.FillTransparence = $iTransparency
	$iError = ($oObj.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeAreaTransparency

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Shape, Shape Style, or Presentation Style's Multi Transparency Gradient settings. See remarks.
; Syntax ........: __LOImpress_ShapeAreaTransparencyGradientMulti(ByRef $oObj[, $avColorStops = Null])
; Parameters ....: $oObj                - [in/out] an object. A Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 4 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 5 Return ? = ColorStop Transparency value not an Integer, less than 0 or greater than 100. Returning problem element index.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve FillTransparenceGradient Struct.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve StopColor Struct.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current version less than 7.6.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple Transparency stops in a Gradient rather than just a beginning and an ending value, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the Transparency percentage value between 0 and 100%.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......: _LO_TransparencyGradientMultiModify, _LO_TransparencyGradientMultiDelete, _LO_TransparencyGradientMultiAdd, _LOImpress_ShapeAreaGradientMulticolor
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeAreaTransparencyGradientMulti(ByRef $oObj, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oObj.FillTransparenceGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($avColorStops) Then
		$atColorStops = $tStyleGradient.ColorStops()
		If Not IsArray($atColorStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		ReDim $avNewColorStops[UBound($atColorStops)][2]

		For $i = 0 To UBound($atColorStops) - 1
			$avNewColorStops[$i][0] = $atColorStops[$i].StopOffset()
			$tStopColor = $atColorStops[$i].StopColor()
			If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$avNewColorStops[$i][1] = Int($tStopColor.Red() * 100) ; One value is the same as all.
			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tStopColor.Red = ($avColorStops[$i][1] / 100)
		$tStopColor.Green = ($avColorStops[$i][1] / 100)
		$tStopColor.Blue = ($avColorStops[$i][1] / 100)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oObj.FillTransparenceGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oObj.FillTransparenceGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeAreaTransparencyGradientMulti

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeGetType
; Description ...: Identify a Shape's type.
; Syntax ........: __LOImpress_ShapeGetType(ByRef $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape's type.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify Shape.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning the Shape's Type, corresponding to one of the Constants $LOI_SHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeGetType(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avShapeTypes[21][2] = [[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.Shape3DSceneObject"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.CustomShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.MeasureShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.EllipseShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.ClosedBezierShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.OpenBezierShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.PolyPolygonShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.PolyLineShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.LineShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.ConnectorShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.OpenFreeHandShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.ClosedFreeHandShape"], [$LOI_SHAPE_TYPE_FORM_CONTROL, "com.sun.star.drawing.ControlShape"], _
			[$LOI_SHAPE_TYPE_IMAGE, "com.sun.star.drawing.GraphicObjectShape"], [$LOI_SHAPE_TYPE_MEDIA, "com.sun.star.drawing.MediaShape"], _
			[$LOI_SHAPE_TYPE_OLE2, "com.sun.star.drawing.OLE2Shape"], [$LOI_SHAPE_TYPE_TABLE, "com.sun.star.drawing.TableShape"], _
			[$LOI_SHAPE_TYPE_TEXTBOX, "com.sun.star.drawing.TextShape"], [$LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE, "com.sun.star.presentation.SubtitleShape"], _
			[$LOI_SHAPE_TYPE_TEXTBOX_TITLE, "com.sun.star.presentation.TitleTextShape"], [$LOI_SHAPE_TYPE_TEXTBOX_OUTLINE, "com.sun.star.presentation.OutlinerShape"]]
	Local $sShapeType

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sShapeType = $oShape.ShapeType()
	If Not IsString($sShapeType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($avShapeTypes) - 1
		If ($sShapeType = $avShapeTypes[$i][1]) Then Return SetError($__LO_STATUS_SUCCESS, 0, $avShapeTypes[$i][0])
	Next

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
EndFunc   ;==>__LOImpress_ShapeGetType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeLineArrowheadNameInsert
; Description ...: Create and insert a preset Arrowhead name.
; Syntax ........: __LOImpress_ShapeLineArrowheadNameInsert(ByRef $oDoc, $iArrowStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iArrowStyle         - an integer value. The Arrowhead style to insert. See Constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iArrowStyle not an Integer, less than 0 or greater than 32. See Constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.MarkerTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.drawing.PolyPolygonBezierCoords" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Arrowhead Name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Arrowhead preset values.
;                  @Error 3 @Extended 3 Return 0 = Error inserting preset Arrowhead Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. The Arrowhead name was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If I don't pre-insert preset Arrowhead styles, the relavent functions will fail when trying to set a line start/end style to them.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeLineArrowheadNameInsert(ByRef $oDoc, $iArrowStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tPolyCoords
	Local $oArrowTable
	Local $sArrowName
	Local $sPolyCoords = "com.sun.star.drawing.PolyPolygonBezierCoords"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iArrowStyle, $LOI_SHAPE_LINE_ARROW_TYPE_NONE, $LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oArrowTable = $oDoc.createInstance("com.sun.star.drawing.MarkerTable")
	If Not IsObj($oArrowTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	; Retrieve the Arrow name.
	$sArrowName = __LOImpress_ShapeLineArrowStyleName($iArrowStyle, Null)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sArrowName <> "") And Not $oArrowTable.hasByName($sArrowName) Then ; Check if Arrow name = "", if so, skip, as it is "None" arrow setting, and has no preset values.
		$tPolyCoords = __LO_CreateStruct($sPolyCoords)
		If Not IsObj($tPolyCoords) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tPolyCoords = __LOImpress_ShapeLineArrowStyleName($iArrowStyle, Null, True, $tPolyCoords)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oArrowTable.insertByName($sArrowName, $tPolyCoords)
		If Not ($oArrowTable.hasByName($sArrowName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_ShapeLineArrowheadNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeLineArrowStyleName
; Description ...: Convert a Arrow head Constant to the corresponding name or reverse, or return preset values.
; Syntax ........: __LOImpress_ShapeLineArrowStyleName([$iArrowStyle = Null[, $sArrowStyle = Null[, $bReturnPresets = False[, $tPolyCoords = Null]]]])
; Parameters ....: $iArrowStyle         - [optional] an integer value (0-32). Default is Null. The Arrow Style Constant to convert to its corresponding name. See $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $sArrowStyle         - [optional] a string value. Default is Null. The Arrow Style Name to convert to the corresponding constant if found.
;                  $bReturnPresets      - [optional] a boolean value. Default is False. If True, the function will try to fill and return a Structure with preset values.
;                  $tPolyCoords         - [optional] a dll struct value. Default is Null. If $bReturnPresets is True, this is a PolyPolygonBezierCoords Structure to fill.
; Return values .: Success: String or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iArrowStyle not an Integer, less than 0 or greater than Arrow type constants. See $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 2 Return 0 = $sArrowStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = $bReturnPresets not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $tPolyCoords not an Object.
;                  @Error 1 @Extended 5 Return 0 = Both $iArrowStyle and $sArrowStyle called with Null.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Shape position point.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Constant called in $iArrowStyle was successfully converted to its corresponding Arrow Type Name.
;                  @Error 0 @Extended 1 Return Integer = Success. Arrow Type Name called in $sArrowStyle was successfully converted to its corresponding Constant value.
;                  @Error 0 @Extended 2 Return String = Success. Arrow Type Name called in $sArrowStyle was not matched to an existing Constant value, returning called name. Possibly a custom value.
;                  @Error 0 @Extended 3 Return String = Success. $bReturnPresets called with True, returning $tPolyCoords filled with preset values.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeLineArrowStyleName($iArrowStyle = Null, $sArrowStyle = Null, $bReturnPresets = False, $tPolyCoords = Null)
	Local $asArrowStyles[33]
	Local $avArray[1]
	Local $atCoords[0]
	Local Const $iX = 0, $iY = 1

	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_NONE] = ""
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_ARROW_SHORT] = "Arrow short"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CONCAVE_SHORT] = "Concave short"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_ARROW] = "Arrow"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_TRIANGLE] = "Triangle"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CONCAVE] = "Concave"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_ARROW_LARGE] = "Arrow large"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CIRCLE] = "Circle"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_SQUARE] = "Square"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_SQUARE_45] = "Square 45"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_DIAMOND] = "Diamond"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE] = "Half Circle"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINES] = "Dimension Lines"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW] = "Dimension Line Arrow"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_DIMENSION_LINE] = "Dimension Line"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_LINE_SHORT] = "Line short"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_LINE] = "Line"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_TRIANGLE_UNFILLED] = "Triangle unfilled"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_DIAMOND_UNFILLED] = "Diamond unfilled"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CIRCLE_UNFILLED] = "Circle unfilled"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED] = "Square 45 unfilled"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_SQUARE_UNFILLED] = "Square unfilled"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED] = "Half Circle unfilled"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_LEFT] = "Half Arrow left"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_RIGHT] = "Half Arrow right"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_REVERSED_ARROW] = "Reversed Arrow"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_DOUBLE_ARROW] = "Double Arrow"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CF_ONE] = "CF One"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CF_ONLY_ONE] = "CF Only One"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CF_MANY] = "CF Many"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CF_MANY_ONE] = "CF Many One"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_ONE] = "CF Zero One"
	$asArrowStyles[$LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY] = "CF Zero Many"

	If Not IsBool($bReturnPresets) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($iArrowStyle <> Null) And Not $bReturnPresets Then
		If Not __LO_IntIsBetween($iArrowStyle, 0, UBound($asArrowStyles) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $asArrowStyles[$iArrowStyle]) ; Return the requested Arrow Style name.

	ElseIf ($sArrowStyle <> Null) Then
		If Not IsString($sArrowStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To UBound($asArrowStyles) - 1
			If ($asArrowStyles[$i] = $sArrowStyle) Then Return SetError($__LO_STATUS_SUCCESS, 1, $i) ; Return the array element where the matching Arrow Style was found.

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 2, $sArrowStyle) ; If no matches, just return the name, as it could be a custom value.

	ElseIf ($iArrowStyle <> Null) And $bReturnPresets Then
		If Not IsObj($tPolyCoords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		Switch $iArrowStyle
			Case $LOI_SHAPE_LINE_ARROW_TYPE_ARROW_SHORT
				Local $aiFlags[4] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[4][2] = [[0, 13], [10, 0], [20, 13], [0, 13]]

				ReDim $atCoords[4]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CONCAVE_SHORT
				Local $aiFlags[11] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[11][2] = [[566, 0], [0, 754], [114, 669], [250, 601], [398, 555], [559, 538], [720, 551], [873, 597], [1013, 665], [1131, 754], _
						[566, 0]]

				ReDim $atCoords[11]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_ARROW
				Local $aiFlags[4] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[4][2] = [[10, 0], [0, 30], [20, 30], [10, 0]]

				ReDim $atCoords[4]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_TRIANGLE
				Local $aiFlags[20] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[20][2] = [[1009, 1050], [560, 42], [538, 12], [509, 0], [475, 12], [454, 38], [5, 1050], [0, 1063], [0, 1071], [5, 1092], _
						[17, 1113], [34, 1126], [55, 1130], [958, 1130], [979, 1126], [1000, 1113], [1009, 1092], [1013, 1071], [1013, 1063], [1009, 1050]]

				ReDim $atCoords[20]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CONCAVE
				Local $aiFlags[11] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[11][2] = [[1013, 1491], [1131, 1580], [564, 0], [0, 1580], [114, 1495], [250, 1427], [398, 1381], [559, 1364], [720, 1377], [873, 1423], _
						[1013, 1491]]

				ReDim $atCoords[11]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_ARROW_LARGE
				Local $aiFlags[4] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[4][2] = [[0, 40], [10, 0], [20, 40], [0, 40]]

				ReDim $atCoords[4]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CIRCLE
				Local $aiFlags[33] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[33][2] = [[462, 1118], [360, 1089], [258, 1038], [165, 966], [93, 873], [42, 771], [13, 669], [0, 564], [13, 462], [42, 356], _
						[93, 254], [165, 165], [258, 93], [360, 43], [462, 9], [568, 0], [669, 9], [775, 43], [873, 93], [966, 165], _
						[1038, 254], [1089, 356], [1118, 462], [1131, 564], [1118, 669], [1089, 771], [1038, 873], [966, 966], [873, 1038], [775, 1089], _
						[669, 1118], [568, 1131], [462, 1118]]

				ReDim $atCoords[33]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_SQUARE
				Local $aiFlags[5] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[5][2] = [[0, 0], [10, 0], [10, 10], [0, 10], [0, 0]]

				ReDim $atCoords[5]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_SQUARE_45
				Local $aiFlags[5] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[5][2] = [[0, 564], [564, 1131], [1131, 564], [564, 0], [0, 564]]

				ReDim $atCoords[5]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_DIAMOND
				Local $aiFlags[5] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[5][2] = [[1500, 0], [3000, 3000], [1500, 6000], [0, 3000], [1500, 0]]

				ReDim $atCoords[5]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE
				Local $aiFlags[16] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[16][2] = [[29, 0], [0, 653], [244, 2596], [1005, 4400], [2188, 5960], [3750, 7142], [5556, 7902], [7450, 8146], [9444, 7902], [11250, 7142], _
						[12812, 5960], [13995, 4400], [14756, 2596], [15000, 653], [14971, 0], [29, 0]]

				ReDim $atCoords[16]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINES
				Local $aiFlags[13] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[13][2] = [[0, 0], [278, 0], [556, 0], [836, 0], [836, 36], [836, 72], [836, 110], [558, 110], [280, 110], [0, 110], _
						[0, 74], [0, 38], [0, 0]]

				ReDim $atCoords[13]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW
				Local $aiFlags[6] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[6][2] = [[0, 0], [0, 110], [418, 110], [836, 110], [836, 0], [0, 0]]

				ReDim $atCoords[6]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_DIMENSION_LINE
				Local $aiFlags[57] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[57][2] = [[0, 0], [0, 110], [418, 110], [414, 111], [409, 112], [403, 114], [397, 118], [393, 122], [389, 126], [13, 639], _
						[7, 650], [5, 659], [4, 665], [5, 675], [9, 684], [13, 691], [17, 695], [22, 699], [29, 703], [36, 706], _
						[46, 707], [55, 707], [65, 704], [74, 698], [82, 689], [422, 225], [763, 689], [768, 695], [772, 699], [778, 703], _
						[784, 705], [791, 707], [798, 707], [803, 707], [811, 705], [816, 703], [823, 699], [829, 694], [833, 688], [837, 681], _
						[840, 673], [840, 664], [840, 658], [838, 651], [836, 645], [832, 639], [456, 127], [453, 122], [449, 119], [445, 116], _
						[439, 113], [433, 111], [428, 110], [426, 110], [836, 110], [836, 0], [0, 0]]

				ReDim $atCoords[57]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_LINE_SHORT
				Local $aiFlags[53] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[53][2] = [[1126, 0], [1108, 1], [1096, 3], [1083, 6], [1067, 13], [1053, 22], [1041, 32], [1030, 45], [23, 1418], [8, 1446], _
						[1, 1472], [0, 1488], [3, 1513], [12, 1538], [24, 1557], [35, 1569], [47, 1579], [66, 1590], [86, 1597], [113, 1601], _
						[137, 1599], [163, 1591], [187, 1576], [208, 1553], [1120, 308], [2032, 1553], [2045, 1569], [2057, 1579], [2073, 1588], [2088, 1594], _
						[2107, 1599], [2126, 1600], [2139, 1599], [2161, 1594], [2175, 1589], [2193, 1578], [2209, 1564], [2220, 1550], [2231, 1529], [2237, 1508], _
						[2239, 1485], [2238, 1468], [2234, 1451], [2227, 1435], [2216, 1418], [1211, 46], [1201, 34], [1192, 26], [1181, 18], [1165, 10], _
						[1149, 5], [1135, 2], [1126, 0]]

				ReDim $atCoords[53]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_LINE
				Local $aiFlags[39] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[39][2] = [[1125, 0], [1080, 9], [1042, 34], [1021, 64], [12, 2074], [1, 2108], [0, 2142], [12, 2183], [39, 2215], [62, 2230], _
						[85, 2239], [111, 2243], [148, 2238], [179, 2223], [202, 2202], [217, 2179], [1122, 375], [2021, 2179], [2031, 2196], [2043, 2210], _
						[2058, 2223], [2073, 2232], [2094, 2240], [2124, 2244], [2144, 2242], [2159, 2238], [2176, 2231], [2192, 2221], [2204, 2211], [2217, 2196], _
						[2229, 2175], [2236, 2153], [2239, 2126], [2237, 2106], [2226, 2076], [1228, 64], [1204, 31], [1168, 8], [1125, 0]]

				ReDim $atCoords[39]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_TRIANGLE_UNFILLED
				Local $aiFlags[4] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[4][2] = [[1500, 0], [3000, 3000], [0, 3000], [1500, 0]]

				ReDim $atCoords[4]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_DIAMOND_UNFILLED
				Local $aiFlags[5] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[5][2] = [[1500, 0], [3000, 3000], [1500, 6000], [0, 3000], [1500, 0]]

				ReDim $atCoords[5]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CIRCLE_UNFILLED
				Local $aiFlags[37] = [$LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[37][2] = [[1500, 3000], [1224, 3000], [989, 2937], [750, 2799], [511, 2661], [339, 2489], [201, 2250], [63, 2011], [0, 1776], [0, 1500], _
						[0, 1224], [63, 989], [201, 750], [339, 511], [511, 339], [750, 201], [989, 63], [1224, 0], [1500, 0], [1776, 0], _
						[2011, 63], [2250, 201], [2489, 339], [2661, 511], [2799, 750], [2937, 989], [3000, 1224], [3000, 1500], [3000, 1776], [2937, 2011], _
						[2799, 2250], [2661, 2489], [2489, 2661], [2250, 2799], [2011, 2937], [1776, 3000], [1500, 3000]]

				ReDim $atCoords[37]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED
				Local $aiFlags[5] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[5][2] = [[1500, 3000], [0, 1500], [1500, 0], [3000, 1500], [1500, 3000]]

				ReDim $atCoords[5]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_SQUARE_UNFILLED
				Local $aiFlags[5] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[5][2] = [[0, 0], [300, 0], [300, 300], [0, 300], [0, 0]]

				ReDim $atCoords[5]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED
				Local $aiFlags[92] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SMOOTH, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SMOOTH, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_SMOOTH, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SMOOTH, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[92][2] = [[14971, 0], [14992, 229], [15000, 423], [15000, 653], [15000, 1343], [14921, 1981], [14756, 2596], [14591, 3210], [14340, 3802], [13995, 4400], _
						[13650, 4997], [13262, 5510], [12812, 5960], [12361, 6410], [11848, 6797], [11250, 7142], [10652, 7487], [10060, 7738], [9444, 7902], [8844, 8063], _
						[8221, 8142], [7550, 8146], [7550, 8746], [7450, 8746], [7450, 8146], [6779, 8142], [6156, 8063], [5556, 7902], [4940, 7738], [4348, 7487], _
						[3750, 7142], [3152, 6797], [2639, 6410], [2188, 5960], [1738, 5510], [1350, 4997], [1005, 4400], [660, 3802], [409, 3210], [244, 2596], _
						[79, 1981], [0, 1343], [0, 653], [0, 423], [8, 229], [29, 0], [327, 26], [626, 52], [608, 263], [600, 442], _
						[600, 653], [600, 1288], [672, 1875], [824, 2440], [975, 3006], [1207, 3550], [1524, 4099], [1842, 4649], [2198, 5121], [2612, 5536], _
						[3027, 5950], [3500, 6305], [4050, 6623], [4600, 6940], [5145, 7171], [5711, 7323], [6277, 7474], [6865, 7546], [7500, 7546], [8135, 7546], _
						[8723, 7474], [9289, 7323], [9855, 7171], [10400, 6940], [10950, 6623], [11500, 6305], [11973, 5950], [12388, 5536], [12802, 5121], [13158, 4649], _
						[13476, 4099], [13793, 3550], [14025, 3006], [14176, 2440], [14328, 1875], [14400, 1288], [14400, 653], [14400, 442], [14392, 263], [14374, 52], _
						[14673, 26], [14971, 0]]

				ReDim $atCoords[92]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_LEFT
				Local $aiFlags[2] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[2][2] = [[10, 0], [10, 0]]

				ReDim $atCoords[2]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_RIGHT
				Local $aiFlags[2] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[2][2] = [[0, 0], [0, 0]]

				ReDim $atCoords[2]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_REVERSED_ARROW
				Local $aiFlags[4] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[4][2] = [[10, 30], [0, 0], [20, 0], [10, 30]]

				ReDim $atCoords[4]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_DOUBLE_ARROW
				Local $aiFlags[8] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[8][2] = [[737, 1131], [1131, 1131], [567, 0], [0, 1131], [398, 1131], [0, 1918], [1131, 1918], [737, 1131]]

				ReDim $atCoords[8]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CF_ONE
				Local $aiFlags[6] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[6][2] = [[19, 40], [20, 40], [20, 0], [18, 0], [18, 40], [19, 40]]

				ReDim $atCoords[6]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CF_ONLY_ONE
				Local $aiFlags[6] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[6][2] = [[19, 40], [20, 40], [20, 0], [18, 0], [18, 40], [19, 40]]

				ReDim $atCoords[6]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CF_MANY
				Local $aiFlags[12] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[12][2] = [[1500, 3000], [3000, 211], [3000, 0], [2886, 0], [1600, 2392], [1600, 0], [1400, 0], [1400, 2392], [114, 0], [0, 0], _
						[0, 211], [1500, 3000]]

				ReDim $atCoords[12]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CF_MANY_ONE
				Local $aiFlags[6] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[6][2] = [[1500, 3200], [3000, 3200], [3000, 3000], [0, 3000], [0, 3200], [1500, 3200]]

				ReDim $atCoords[6]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_ONE
				Local $aiFlags[38] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SMOOTH, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[38][2] = [[100, 4300], [100, 4037], [169, 3778], [301, 3551], [433, 3322], [622, 3133], [850, 3001], [1078, 2870], [1337, 2800], [1600, 2800], _
						[1863, 2800], [2122, 2870], [2350, 3001], [2578, 3133], [2767, 3322], [2899, 3550], [3031, 3779], [3100, 4037], [3100, 4300], [3100, 4564], _
						[3031, 4822], [2899, 5050], [2767, 5279], [2578, 5468], [2350, 5600], [2122, 5731], [1863, 5801], [1600, 5801], [1337, 5801], [1078, 5731], _
						[850, 5600], [622, 5468], [433, 5279], [301, 5051], [169, 4823], [100, 4564], [100, 4301], [100, 4300]]

				ReDim $atCoords[38]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next

			Case $LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY
				Local $aiFlags[38] = [$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SMOOTH, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL, _
						$LOI_DRAWSHAPE_POINT_TYPE_CONTROL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL]
				Local $aiIndvCoords[38][2] = [[0, 4500], [0, 4237], [69, 3978], [201, 3751], [333, 3522], [522, 3333], [750, 3201], [978, 3070], [1237, 3000], [1500, 3000], _
						[1763, 3000], [2022, 3070], [2250, 3201], [2478, 3333], [2667, 3522], [2799, 3750], [2931, 3979], [3000, 4237], [3000, 4500], [3000, 4764], _
						[2931, 5022], [2799, 5250], [2667, 5479], [2478, 5668], [2250, 5800], [2022, 5931], [1763, 6001], [1500, 6001], [1237, 6001], [978, 5931], _
						[750, 5800], [522, 5668], [333, 5479], [201, 5251], [69, 5023], [0, 4764], [0, 4501], [0, 4500]]

				ReDim $atCoords[38]

				For $i = 0 To UBound($atCoords) - 1
					$atCoords[$i] = __LOImpress_CreatePoint($aiIndvCoords[$i][$iX], $aiIndvCoords[$i][$iY])
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next
		EndSwitch

		$avArray[0] = $atCoords
		$tPolyCoords.Coordinates = $avArray

		$avArray[0] = $aiFlags
		$tPolyCoords.Flags = $avArray

		Return SetError($__LO_STATUS_SUCCESS, 3, $tPolyCoords)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; No values called.
	EndIf
EndFunc   ;==>__LOImpress_ShapeLineArrowStyleName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeLineDashNameInsert
; Description ...: Create and insert a preset Line Dash name.
; Syntax ........: __LOImpress_ShapeLineDashNameInsert(ByRef $oDoc, $iLineDashType)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iLineDashType       - an integer value. The Line Dash style to insert. See Constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iLineDashType not an Integer, less than 0 or greater than 31. See Constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.DashTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.drawing.LineDash" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Line Dash name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Line Dash preset values.
;                  @Error 3 @Extended 3 Return 0 = Error inserting Line Dash Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. The Line Dash name was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If I don't pre-insert preset line dash styles, the relavent functions will fail when trying to set a line style to them.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeLineDashNameInsert(ByRef $oDoc, $iLineDashType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tNewDash
	Local $sDashName
	Local $oDashTable

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iLineDashType, $LOI_SHAPE_LINE_STYLE_NONE, $LOI_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDashTable = $oDoc.createInstance("com.sun.star.drawing.DashTable")
	If Not IsObj($oDashTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	; Retrieve the Dash name.
	$sDashName = __LOImpress_ShapeLineStyleName($iLineDashType, Null)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If Not $oDashTable.hasByName($sDashName) Then
		$tNewDash = __LO_CreateStruct("com.sun.star.drawing.LineDash")
		If Not IsObj($tNewDash) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tNewDash = __LOImpress_ShapeLineStyleName($iLineDashType, Null, True, $tNewDash)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oDashTable.insertByName($sDashName, $tNewDash)
		If Not ($oDashTable.hasByName($sDashName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_ShapeLineDashNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeLineStyleName
; Description ...: Convert a Line Style Constant to the corresponding name or reverse, or return preset values.
; Syntax ........: __LOImpress_ShapeLineStyleName([$iLineStyle = Null[, $sLineStyle = Null[, $bReturnPresets = False[, $tDash = Null]]]])
; Parameters ....: $iLineStyle          - [optional] an integer value (0-31). Default is Null. The Line Style Constant to convert to its corresponding name. See $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3
;                  $sLineStyle          - [optional] a string value. Default is Null. The Line Style Name to convert to the corresponding constant if found.
;                  $bReturnPresets      - [optional] a boolean value. Default is False. If True, the function will try to fill and return a Structure with preset values.
;                  $tDash               - [optional] a dll struct value. Default is Null. If $bReturnPresets is True, this is a Dash Structure to fill.
; Return values .: Success: String, Structure or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iLineStyle not an Integer, less than 0 or greater than Line Style constants. See $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 2 Return 0 = $sLineStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = $bReturnPresets not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $tDash not an Object.
;                  @Error 1 @Extended 5 Return 0 = Both $iLineStyle and $sLineStyle called with Null.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Constant called in $iLineStyle was successfully converted to its corresponding Line Style Name.
;                  @Error 0 @Extended 1 Return Integer = Success. Line Style Name called in $sLineStyle was successfully converted to its corresponding Constant value.
;                  @Error 0 @Extended 2 Return String = Success. Line Style Name called in $sLineStyle was not matched to an existing Constant value, returning called name. Possibly a custom value.
;                  @Error 0 @Extended 3 Return Struct = Success. $bReturnPresets called with True, returning $tDash filled with preset values.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeLineStyleName($iLineStyle = Null, $sLineStyle = Null, $bReturnPresets = False, $tDash = Null)
	Local $asLineStyles[32]
	Local Const $__LOI_DASH_STYLE_RECT = 0, $__LOI_DASH_STYLE_RECT_RELATIVE = 2, $__LOI_DASH_STYLE_ROUND_RELATIVE = 3 ; $__LOI_DASH_STYLE_ROUND = 1 (Not Used), com.sun.star.drawing.DashStyle

	; $LOI_SHAPE_LINE_STYLE_NONE, $LOI_SHAPE_LINE_STYLE_CONTINUOUS, don't have a name, so to keep things symmetrical I created my own, but those two won't be used.
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_NONE] = "NONE"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_CONTINUOUS] = "CONTINUOUS"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DOT] = "Dot"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DOT_ROUNDED] = "Dot (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_LONG_DOT] = "Long Dot"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_LONG_DOT_ROUNDED] = "Long Dot (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DASH] = "Dash"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DASH_ROUNDED] = "Dash (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_LONG_DASH] = "Long Dash"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_LONG_DASH_ROUNDED] = "Long Dash (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DOUBLE_DASH] = "Double Dash"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_ROUNDED] = "Double Dash (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DASH_DOT] = "Dash Dot"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DASH_DOT_ROUNDED] = "Dash Dot (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_LONG_DASH_DOT] = "Long Dash Dot"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_LONG_DASH_DOT_ROUNDED] = "Long Dash Dot (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT] = "Double Dash Dot"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED] = "Double Dash Dot (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DASH_DOT_DOT] = "Dash Dot Dot"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DASH_DOT_DOT_ROUNDED] = "Dash Dot Dot (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT] = "Double Dash Dot Dot"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED] = "Double Dash Dot Dot (Rounded)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_ULTRAFINE_DOTTED] = "Ultrafine Dotted (var)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_FINE_DOTTED] = "Fine Dotted"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_ULTRAFINE_DASHED] = "Ultrafine Dashed"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_FINE_DASHED] = "Fine Dashed"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_DASHED] = "Dashed (var)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_SPARSE_DASH] = "Sparse Dash"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_3_DASHES_3_DOTS] = "3 Dashes 3 Dots (var)"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES] = "Ultrafine 2 Dots 3 Dashes"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_2_DOTS_1_DASH] = "2 Dots 1 Dash"
	$asLineStyles[$LOI_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS] = "Line with Fine Dots"

	If Not __LO_VersionCheck(24.2) Then $asLineStyles[$LOI_SHAPE_LINE_STYLE_SPARSE_DASH] = "Line Style 9"

	If Not IsBool($bReturnPresets) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($iLineStyle <> Null) And Not $bReturnPresets Then
		If Not __LO_IntIsBetween($iLineStyle, 0, UBound($asLineStyles) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $asLineStyles[$iLineStyle]) ; Return the requested Line Style name.

	ElseIf ($sLineStyle <> Null) Then
		If Not IsString($sLineStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To UBound($asLineStyles) - 1
			If ($asLineStyles[$i] = $sLineStyle) Then Return SetError($__LO_STATUS_SUCCESS, 1, $i) ; Return the array element where the matching Line Style was found.

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 2, $sLineStyle) ; If no matches, just return the name, as it could be a custom value.

	ElseIf ($iLineStyle <> Null) And $bReturnPresets Then
		If Not IsObj($tDash) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		Switch $iLineStyle
			Case $LOI_SHAPE_LINE_STYLE_DOT
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 100
					.DotLen = 100
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DOT_ROUNDED
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 199
					.DotLen = 1
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_LONG_DOT
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 300
					.DotLen = 100
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_LONG_DOT_ROUNDED
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 399
					.DotLen = 1
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DASH
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 100
					.DotLen = 300
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DASH_ROUNDED
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 199
					.DotLen = 201
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_LONG_DASH
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 300
					.DotLen = 400
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_LONG_DASH_ROUNDED
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 399
					.DotLen = 301
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DOUBLE_DASH
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 300
					.DotLen = 800
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_ROUNDED
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 399
					.DotLen = 701
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DASH_DOT
				With $tDash
					.Dashes = 1
					.DashLen = 100
					.Distance = 100
					.DotLen = 300
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DASH_DOT_ROUNDED
				With $tDash
					.Dashes = 1
					.DashLen = 1
					.Distance = 199
					.DotLen = 201
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_LONG_DASH_DOT
				With $tDash
					.Dashes = 1
					.DashLen = 100
					.Distance = 300
					.DotLen = 400
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_LONG_DASH_DOT_ROUNDED
				With $tDash
					.Dashes = 1
					.DashLen = 1
					.Distance = 399
					.DotLen = 301
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT
				With $tDash
					.Dashes = 1
					.DashLen = 100
					.Distance = 300
					.DotLen = 800
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED
				With $tDash
					.Dashes = 1
					.DashLen = 1
					.Distance = 399
					.DotLen = 701
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DASH_DOT_DOT
				With $tDash
					.Dashes = 2
					.DashLen = 100
					.Distance = 100
					.DotLen = 300
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DASH_DOT_DOT_ROUNDED
				With $tDash
					.Dashes = 2
					.DashLen = 1
					.Distance = 199
					.DotLen = 201
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT
				With $tDash
					.Dashes = 2
					.DashLen = 100
					.Distance = 300
					.DotLen = 800
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED
				With $tDash
					.Dashes = 2
					.DashLen = 1
					.Distance = 399
					.DotLen = 701
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_ROUND_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_ULTRAFINE_DOTTED
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 50
					.DotLen = 0
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_FINE_DOTTED
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 457
					.DotLen = 0
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_ULTRAFINE_DASHED
				With $tDash
					.Dashes = 1
					.DashLen = 51
					.Distance = 51
					.DotLen = 51
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_FINE_DASHED
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 197
					.DotLen = 197
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_DASHED
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 127
					.DotLen = 197
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_SPARSE_DASH
				With $tDash
					.Dashes = 0
					.DashLen = 0
					.Distance = 120
					.DotLen = 197
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_3_DASHES_3_DOTS
				With $tDash
					.Dashes = 3
					.DashLen = 0
					.Distance = 100
					.DotLen = 197
					.Dots = 3
					.Style = $__LOI_DASH_STYLE_RECT_RELATIVE
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES
				With $tDash
					.Dashes = 3
					.DashLen = 254
					.Distance = 127
					.DotLen = 51
					.Dots = 2
					.Style = $__LOI_DASH_STYLE_RECT
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_2_DOTS_1_DASH
				With $tDash
					.Dashes = 1
					.DashLen = 203
					.Distance = 203
					.DotLen = 0
					.Dots = 2
					.Style = $__LOI_DASH_STYLE_RECT
				EndWith

			Case $LOI_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS
				With $tDash
					.Dashes = 10
					.DashLen = 0
					.Distance = 152
					.DotLen = 2007
					.Dots = 1
					.Style = $__LOI_DASH_STYLE_RECT
				EndWith
		EndSwitch

		Return SetError($__LO_STATUS_SUCCESS, 3, $tDash)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; No values called.
	EndIf
EndFunc   ;==>__LOImpress_ShapeLineStyleName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapePresStyleNumCreateScript
; Description ...: Part of the Presentation Style Numbering Modification workaround, creates a Macro in a document.
; Syntax ........: __LOImpress_ShapePresStyleNumCreateScript(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Standard Macro Library.
;                  @Error 3 @Extended 2 Return 0 = Error creating Macro in Document.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving Script Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Function successfully created the Macro in Document. Returning Script Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapePresStyleNumCreateScript(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sNumStyleScript = "Function ReplaceByIndex(oNumRules As Object, iIndex%, vSettings As Variant)" & @CRLF & _
			"oNumRules.replaceByIndex(iIndex,vSettings)" & @CRLF & _
			"ReplaceByIndex = oNumRules" & @CRLF & _
			"End Function"
	Local $oStandardLibrary, $oScript

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; Retrieving the BasicLibrary.Standard Object fails when using a newly opened document, I found a workaround by updating the
	; following setting.
	$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

	$oStandardLibrary = $oDoc.BasicLibraries.Standard()
	If Not IsObj($oStandardLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then $oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")

	$oStandardLibrary.insertByName("AU3LibreOffice_UDF_Macros", $sNumStyleScript)
	If Not $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.AU3LibreOffice_UDF_Macros.ReplaceByIndex?language=Basic&location=document")
	If Not IsObj($oScript) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oScript)
EndFunc   ;==>__LOImpress_ShapePresStyleNumCreateScript

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapePresStyleNumDeleteScript
; Description ...: Part of the Presentation Style Numbering Modification workaround, deletes a Macro in a document.
; Syntax ........: __LOImpress_ShapePresStyleNumDeleteScript(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Standard Macro Library.
;                  @Error 3 @Extended 2 Return 0 = Error deleting Macro.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Function successfully deleted the Macro in Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapePresStyleNumDeleteScript(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oStandardLibrary

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; Retrieving the BasicLibrary.Standard Object fails when using a newly opened document, I found a workaround by updating the
	; following setting.
	$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

	$oStandardLibrary = $oDoc.BasicLibraries.Standard()
	If Not IsObj($oStandardLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then $oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")

	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_ShapePresStyleNumDeleteScript

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapePresStyleNumInitiateDocument
; Description ...: Part of the work around method for modifying Presentation Style Numbering settings.
; Syntax ........: __LOImpress_ShapePresStyleNumInitiateDocument()
; Parameters ....: None
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.Desktop" Object.
;                  @Error 2 @Extended 3 Return 0 = Error Creating document.
;                  @Error 2 @Extended 4 Return 0 = Error creating AU3LibreOffice_UDF_Macros Module in document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting Hidden
;                  |                               2 = Error setting MacroExecutionMode
;                  |                               4 = Error setting ReadOnly
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. The Numbering Style Modification Document was successfully created.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapePresStyleNumInitiateDocument()
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iMacroExecMode_ALWAYS_EXECUTE_NO_WARN = 4, $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $iError = 0
	Local $oNumStyleDoc, $oServiceManager, $oDesktop
	Local $atProperties[3]
	Local $vProperty

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$vProperty = __LO_SetPropertyValue("Hidden", True)
	If @error Then $iError = BitOR($iError, 1)
	If Not BitAND($iError, 1) Then $atProperties[0] = $vProperty

	$vProperty = __LO_SetPropertyValue("MacroExecutionMode", $iMacroExecMode_ALWAYS_EXECUTE_NO_WARN)
	If @error Then $iError = BitOR($iError, 2)
	If Not BitAND($iError, 2) Then $atProperties[1] = $vProperty

	$vProperty = __LO_SetPropertyValue("ReadOnly", True)
	If @error Then $iError = BitOR($iError, 4)
	If Not BitAND($iError, 4) Then $atProperties[2] = $vProperty

	$oNumStyleDoc = $oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", $iURLFrameCreate, $atProperties)
	If Not IsObj($oNumStyleDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	__LOImpress_ShapePresStyleNumCreateScript($oNumStyleDoc)
	If (@error > 0) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oNumStyleDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oNumStyleDoc))
EndFunc   ;==>__LOImpress_ShapePresStyleNumInitiateDocument

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapePresStyleNumModify
; Description ...: Internal function for modifying Presentation Style Numbering settings.
; Syntax ........: __LOImpress_ShapePresStyleNumModify(ByRef $oDoc, ByRef $oNumRules, $iLevel, $atNumLevel)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function, to modify NumberingRules for.
;                  $oNumRules           - [in/out] an object. The Numbering Rules object retrieved from a Numbering Style.
;                  $iLevel              - an integer value (-1-9). The Numbering Style level to modify. -1 = all levels.
;                  $atNumLevel          - an array of dll structs. An array of Numbering Rule settings retrieved from a Numbering Style.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oNumRules not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iLevel not between -1 and 9 to indicate correct level.
;                  @Error 1 @Extended 4 Return 0 = $atNumLevel not an array.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error opening new document, and inserting ReplaceByIndex Script.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving "Standard.AU3LibreOffice_UDF_Macros.ReplaceByIndex" Macro in new document.
;                  @Error 3 @Extended 3 Return 0 = Error deleting ReplaceByIndex Macro from Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully set the requested settings.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This works, but only with a work-around method, see inside this function for a description of why a work-around method is necessary.
;                  When a lot of settings are set, especially for all levels, this function can be a bit slow.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapePresStyleNumModify(ByRef $oDoc, ByRef $oNumRules, $iLevel, $atNumLevel)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumStyleDoc, $oScript
	Local $aDummyArray[0], $avParamArray[3]
	Local $bNumDocOpen = False

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNumRules) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iLevel, -1, 9) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oScript = __LOImpress_ShapePresStyleNumCreateScript($oDoc) ; Create my modification Script.

	If Not IsObj($oScript) Then ; If creating my Mod. Script fails, open a new document and create a script in there.
		$oNumStyleDoc = __LOImpress_ShapePresStyleNumInitiateDocument()
		If Not IsObj($oNumStyleDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oScript = $oNumStyleDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.AU3LibreOffice_UDF_Macros.ReplaceByIndex?language=Basic&location=document")
		If Not IsObj($oScript) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$bNumDocOpen = True
	EndIf

	; $oNumRules.replaceByIndex($iLevel, $atNumLevel); This should work but doesn't -- It would seem that the Array passed by
	; AutoIt is not recognized as an appropriate array(or Sequence) by LibreOffice, or perhaps as variable type "Any", which is
	; what LibreOffice replace by index is expecting, and consequently causes a com.sun.star.lang.IllegalArgumentException COM error.

	$avParamArray[0] = $oNumRules
	$avParamArray[1] = $iLevel
	$avParamArray[2] = $atNumLevel

	$oNumRules = $oScript.Invoke($avParamArray, $aDummyArray, $aDummyArray)

	If ($bNumDocOpen = True) Then
		$oNumStyleDoc.Close(True)

	Else
		__LOImpress_ShapePresStyleNumDeleteScript($oDoc)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_ShapePresStyleNumModify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeStyleAreaGradient
; Description ...: Modify or retrieve the settings for Shape, Shape Style or Presentation Style Background color Gradient.
; Syntax ........: __LOImpress_ShapeStyleAreaGradient(ByRef $oDoc, ByRef $oObj[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oObj                - [in/out] an object. A Shape Style or Presentation Style object returned by a previous _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See remarks. See constants, $LOI_GRAD_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient type to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iIncrement          - [optional] an integer value (0, 3-256). Default is Null. The number of steps of color change. 0 = Automatic.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] an integer value (0-16777215). Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] an integer value (0-16777215). Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sGradientName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error 1 @Extended 6 Return 0 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error 1 @Extended 9 Return 0 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 10 Return 0 = $iFromColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 11 Return 0 = $iToColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 12 Return 0 = $iFromIntense not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 13 Return 0 = $iToIntense not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving "FillGradient" Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 3 Return 0 = Error creating Gradient Name.
;                  @Error 3 @Extended 4 Return 0 = Error setting Gradient Name.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
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
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 0 Return 2 = Success. Gradient has been successfully turned off.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeStyleAreaGradient(ByRef $oDoc, ByRef $oObj, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName
	Local $atColorStop[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tStyleGradient = $oObj.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oObj.FillGradientName(), $tStyleGradient.Style(), _
				$oObj.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oObj.FillStyle() <> $LOI_AREA_FILL_STYLE_GRADIENT) Then $oObj.FillStyle = $LOI_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		__LOImpress_GradientPresets($oDoc, $oObj, $tStyleGradient, $sGradientName)
		$iError = ($oObj.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oObj.FillStyle = $LOI_AREA_FILL_STYLE_OFF
			$oObj.FillGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOI_GRAD_TYPE_LINEAR, $LOI_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oObj.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iFromColor <> Null) Then
		If Not __LO_IntIsBetween($iFromColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tStyleGradient.StartColor = $iFromColor

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tColorStop = $atColorStop[0] ; StopOffset 0 is the "From Color" Value.

			$tStopColor = $tColorStop.StopColor()

			$tStopColor.Red = (BitAND(BitShift($iFromColor, 16), 0xff) / 255)
			$tStopColor.Green = (BitAND(BitShift($iFromColor, 8), 0xff) / 255)
			$tStopColor.Blue = (BitAND($iFromColor, 0xff) / 255)

			$tColorStop.StopColor = $tStopColor

			$atColorStop[0] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($iToColor <> Null) Then
		If Not __LO_IntIsBetween($iToColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tStyleGradient.EndColor = $iToColor

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; Last StopOffset is the "To Color" Value.

			$tStopColor = $tColorStop.StopColor()

			$tStopColor.Red = (BitAND(BitShift($iToColor, 16), 0xff) / 255)
			$tStopColor.Green = (BitAND(BitShift($iToColor, 8), 0xff) / 255)
			$tStopColor.Blue = (BitAND($iToColor, 0xff) / 255)

			$tColorStop.StopColor = $tStopColor

			$atColorStop[UBound($atColorStop) - 1] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($iFromIntense <> Null) Then
		If Not __LO_IntIsBetween($iFromIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$tStyleGradient.StartIntensity = $iFromIntense
	EndIf

	If ($iToIntense <> Null) Then
		If Not __LO_IntIsBetween($iToIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$tStyleGradient.EndIntensity = $iToIntense
	EndIf

	If ($oObj.FillGradientName() = "") Or __LOImpress_GradientIsModified($tStyleGradient, $oObj.FillGradientName()) Then
		$sGradName = __LOImpress_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oObj.FillGradientName = $sGradName
		If ($oObj.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oObj.FillGradient = $tStyleGradient

	; Error checking
	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oObj.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oObj.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oObj.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oObj.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oObj.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = (__LO_VarsAreNull($iFromColor)) ? ($iError) : (($oObj.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = (__LO_VarsAreNull($iToColor)) ? ($iError) : (($oObj.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = (__LO_VarsAreNull($iFromIntense)) ? ($iError) : (($oObj.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = (__LO_VarsAreNull($iToIntense)) ? ($iError) : (($oObj.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeStyleAreaGradient

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeStyleAreaTransparencyGradient
; Description ...: Set or retrieve the Shape, Shape Style or Presentation Style transparency gradient settings.
; Syntax ........: __LOImpress_ShapeStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oObj[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oObj                - [in/out] an object. A Shape Style or Presentation Style object returned by a previous _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient that you want to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3. Call with $LOI_GRAD_TYPE_OFF to turn Transparency Gradient off.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iStart              - [optional] an integer value (0-100). Default is Null. The transparency value for the beginning point of the gradient, where 0% is fully opaque and 100% is fully transparent.
;                  $iEnd                - [optional] an integer value (0-100). Default is Null. The transparency value for the endpoint of the gradient, where 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iType not an Integer, less than -1 or greater than 5. See constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 5 Return 0 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error 1 @Extended 7 Return 0 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iStart not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iEnd not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving "FillTransparenceGradient" Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 3 Return 0 = Error creating Transparency Gradient Name.
;                  @Error 3 @Extended 4 Return 0 = Error setting Transparency Gradient Name.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iType
;                  |                               2 = Error setting $iXCenter
;                  |                               4 = Error setting $iYCenter
;                  |                               8 = Error setting $iAngle
;                  |                               16 = Error setting $iTransitionStart
;                  |                               32 = Error setting $iStart
;                  |                               64 = Error setting $iEnd
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 0 Return 2 = Success. Transparency Gradient has been successfully turned off.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oObj, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop[0]
	Local $fValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tGradient = $oObj.FillTransparenceGradient()
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		__LO_ArrayFill($aiTransparent, $tGradient.Style(), $tGradient.XOffset(), $tGradient.YOffset(), _
				Int($tGradient.Angle() / 10), $tGradient.Border(), __LOImpress_TransparencyGradientConvert(Null, $tGradient.StartColor()), _
				__LOImpress_TransparencyGradientConvert(Null, $tGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oObj.FillTransparenceGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOI_GRAD_TYPE_LINEAR, $LOI_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tGradient.Style = $iType
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tGradient.Border = $iTransitionStart
	EndIf

	If ($iStart <> Null) Then
		If Not __LO_IntIsBetween($iStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tGradient.StartColor = __LOImpress_TransparencyGradientConvert($iStart)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tColorStop = $atColorStop[0] ; StopOffset 0 is the "Start" Value.

			$tStopColor = $tColorStop.StopColor()

			$fValue = $iStart / 100 ; Value is a decimal percentage value.

			$tStopColor.Red = $fValue
			$tStopColor.Green = $fValue
			$tStopColor.Blue = $fValue

			$tColorStop.StopColor = $tStopColor

			$atColorStop[0] = $tColorStop

			$tGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($iEnd <> Null) Then
		If Not __LO_IntIsBetween($iEnd, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tGradient.EndColor = __LOImpress_TransparencyGradientConvert($iEnd)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; StopOffset 0 is the "End" Value.

			$tStopColor = $tColorStop.StopColor()

			$fValue = $iEnd / 100 ; Value is a decimal percentage value.

			$tStopColor.Red = $fValue
			$tStopColor.Green = $fValue
			$tStopColor.Blue = $fValue

			$tColorStop.StopColor = $tStopColor

			$atColorStop[UBound($atColorStop) - 1] = $tColorStop

			$tGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($oObj.FillTransparenceGradientName() = "") Then
		$sTGradName = __LOImpress_TransparencyGradientNameInsert($oDoc, $tGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oObj.FillTransparenceGradientName = $sTGradName
		If ($oObj.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oObj.FillTransparenceGradient = $tGradient

	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oObj.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oObj.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oObj.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oObj.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oObj.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iStart)) ? ($iError) : (($oObj.FillTransparenceGradient.StartColor() = __LOImpress_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iEnd)) ? ($iError) : (($oObj.FillTransparenceGradient.EndColor() = __LOImpress_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeStyleAreaTransparencyGradient

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeStyleCompare
; Description ...: Test whether a set and current Shape Style match.
; Syntax ........: __LOImpress_ShapeStyleCompare(ByRef $oDoc, $sCurShapeStyle, $sSetShapeStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sCurShapeStyle      - a string value. The currently set Shape Style's name.
;                  $sSetShapeStyle      - a string value. The Shape Style's name intended to be set.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sCurShapeStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = $sSetShapeStyle not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape Style's internal name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if both style names are the same, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This is to aid in preventing false Property setting errors if a Display name is used instead of the Internal name for a style.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeStyleCompare(ByRef $oDoc, $sCurShapeStyle, $sSetShapeStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sInternalSetShapeStyleName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sCurShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSetShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($sCurShapeStyle = $sSetShapeStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	If $oDoc.StyleFamilies.getByName("graphics").hasByName($sSetShapeStyle) Then
		$sInternalSetShapeStyleName = $oDoc.StyleFamilies.getByName("graphics").getByName($sSetShapeStyle).Name()
		If Not IsString($sInternalSetShapeStyleName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If ($sCurShapeStyle = $sInternalSetShapeStyleName) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>__LOImpress_ShapeStyleCompare

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeStyleLineArrowStyles
; Description ...: Set or Retrieve Shape Style or Presentation Style Line Start and End Arrow Style settings.
; Syntax ........: __LOImpress_ShapeStyleLineArrowStyles(ByRef $oDoc, ByRef $oObj[, $vStartStyle = Null[, $iStartWidth = Null[, $bStartCenter = Null[, $bSync = Null[, $vEndStyle = Null[, $iEndWidth = Null[, $bEndCenter = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oObj                - [in/out] an object. A Shape Style or Presentation Style object returned by a previous _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $vStartStyle         - [optional] a variant value (0-32, or String). Default is Null. The Arrow head to apply to the start of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iStartWidth         - [optional] an integer value (0-5004). Default is Null. The Width of the Starting Arrowhead, in Hundredths of a Millimeter (HMM).
;                  $bStartCenter        - [optional] a boolean value. Default is Null. If True, Places the center of the Start arrowhead on the endpoint of the line.
;                  $bSync               - [optional] a boolean value. Default is Null. If True, Synchronizes the Start Arrowhead settings with the end Arrowhead settings. See remarks.
;                  $vEndStyle           - [optional] a variant value (0-32, or String). Default is Null. The Arrow head to apply to the end of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iEndWidth           - [optional] an integer value (0-5004). Default is Null. The Width of the Ending Arrowhead, in Hundredths of a Millimeter (HMM).
;                  $bEndCenter          - [optional] a boolean value. Default is Null. If True, Places the center of the End arrowhead on the endpoint of the line.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = $vStartStyle not a String, and not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $vStartStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iStartWidth not an Integer, less than 0 or greater than 5004.
;                  @Error 1 @Extended 6 Return 0 = $bStartCenter not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bSync not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $vEndStyle not a String, and not an Integer.
;                  @Error 1 @Extended 9 Return 0 = $vSEndStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $iEndWidth not an Integer, less than 0 or greater than 5004.
;                  @Error 1 @Extended 11 Return 0 = $bEndCenter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to convert Constant to Arrowhead name.
;                  @Error 3 @Extended 2 Return 0 = Failed to insert preset Arrowhead name and style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $vStartStyle
;                  |                               2 = Error setting $iStartWidth
;                  |                               4 = Error setting $bStartCenter
;                  |                               8 = Error setting $bSync
;                  |                               16 = Error setting $vEndStyle
;                  |                               32 = Error setting $iEndWidth
;                  |                               64 = Error setting $bEndCenter
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function works for connector shapes also.
;                  When the arrowhead type "Arrow" is set in the LO UI, or upon creation of a line with arrows, the internal name of the arrowhead is set to an incrementing name of "Arrowheads x", where x is an Integer value. Since I have no way to determine if the head is a custom arrowhead or supposed to be the "Arrow" type, the return when this is present will be the name "Arrowheads x", and not $LOI_SHAPE_LINE_ARROW_TYPE_ARROW.
;                  When setting an Arrowhead to be $LOI_SHAPE_LINE_ARROW_TYPE_ARROW, the head is set correctly, but the LibreOffice UI will show "None". The return for Arrowhead type will be correct, $LOI_SHAPE_LINE_ARROW_TYPE_ARROW.
;                  LibreOffice has no setting for $bSync, so I have made a manual version of it in this function. It only accepts True, and must be called with True each time you want it to synchronize.
;                  When retrieving the current settings, $bSync will be a Boolean value of whether the Start Arrowhead settings are currently equal to the End Arrowhead setting values.
;                  Both $vStartStyle and $vEndStyle accept a String or an Integer because there is the possibility of a custom Arrowhead being available the user may want to use.
;                  When retrieving the current settings, both $vStartStyle and $vEndStyle could be either an Integer or a String. It will be a String if the current Arrowhead is a custom Arrowhead, else an Integer, corresponding to one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeStyleLineArrowStyles(ByRef $oDoc, ByRef $oObj, $vStartStyle = Null, $iStartWidth = Null, $bStartCenter = Null, $bSync = Null, $vEndStyle = Null, $iEndWidth = Null, $bEndCenter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avArrow[7]
	Local $sStartStyle, $sEndStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($vStartStyle, $iStartWidth, $bStartCenter, $bSync, $vEndStyle, $iEndWidth, $bEndCenter) Then
		__LO_ArrayFill($avArrow, __LOImpress_ShapeLineArrowStyleName(Null, $oObj.LineStartName()), $oObj.LineStartWidth(), $oObj.LineStartCenter(), _
				((($oObj.LineStartName() = $oObj.LineEndName()) And ($oObj.LineStartWidth() = $oObj.LineEndWidth()) And ($oObj.LineStartCenter() = $oObj.LineEndCenter())) ? (True) : (False)), _ ; See if Start and End are the same.
				__LOImpress_ShapeLineArrowStyleName(Null, $oObj.LineEndName()), $oObj.LineEndWidth(), $oObj.LineEndCenter())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avArrow)
	EndIf

	If ($vStartStyle <> Null) Then
		If Not IsString($vStartStyle) And Not IsInt($vStartStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If IsInt($vStartStyle) Then
			If Not __LO_IntIsBetween($vStartStyle, $LOI_SHAPE_LINE_ARROW_TYPE_NONE, $LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$sStartStyle = __LOImpress_ShapeLineArrowStyleName($vStartStyle)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			__LOImpress_ShapeLineArrowheadNameInsert($oDoc, $vStartStyle)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Else
			$sStartStyle = $vStartStyle
		EndIf

		$oObj.LineStartName = $sStartStyle
		$iError = ($oObj.LineStartName() = $sStartStyle) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iStartWidth <> Null) Then
		If Not __LO_IntIsBetween($iStartWidth, 0, 5004) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.LineStartWidth = $iStartWidth
		$iError = (__LO_IntIsBetween($oObj.LineStartWidth(), $iStartWidth - 1, $iStartWidth + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bStartCenter <> Null) Then
		If Not IsBool($bStartCenter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.LineStartCenter = $bStartCenter
		$iError = ($oObj.LineStartCenter() = $bStartCenter) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bSync <> Null) Then
		If Not IsBool($bSync) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		If ($bSync = True) Then
			$oObj.LineEndName = $oObj.LineStartName()
			$oObj.LineEndWidth = $oObj.LineStartWidth()
			$oObj.LineEndCenter = $oObj.LineStartCenter()
			$iError = (($oObj.LineStartName() = $oObj.LineEndName()) And _
					($oObj.LineStartWidth() = $oObj.LineEndWidth()) And _
					($oObj.LineStartCenter() = $oObj.LineEndCenter())) ? ($iError) : (BitOR($iError, 8))
		EndIf
	EndIf

	If ($vEndStyle <> Null) Then
		If Not IsString($vEndStyle) And Not IsInt($vEndStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		If IsInt($vEndStyle) Then
			If Not __LO_IntIsBetween($vEndStyle, $LOI_SHAPE_LINE_ARROW_TYPE_NONE, $LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$sEndStyle = __LOImpress_ShapeLineArrowStyleName($vEndStyle)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			__LOImpress_ShapeLineArrowheadNameInsert($oDoc, $vEndStyle)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Else
			$sEndStyle = $vEndStyle
		EndIf

		$oObj.LineEndName = $sEndStyle
		$iError = ($oObj.LineEndName() = $sEndStyle) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iEndWidth <> Null) Then
		If Not __LO_IntIsBetween($iEndWidth, 0, 5004) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oObj.LineEndWidth = $iEndWidth
		$iError = (__LO_IntIsBetween($oObj.LineEndWidth(), $iEndWidth - 1, $iEndWidth + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bEndCenter <> Null) Then
		If Not IsBool($bEndCenter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oObj.LineEndCenter = $bEndCenter
		$iError = ($oObj.LineEndCenter() = $bEndCenter) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeStyleLineArrowStyles

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeStyleLineProperties
; Description ...: Set or Retrieve Shape Style or Presentation Style Line settings.
; Syntax ........: ; Syntax ........: __LOImpress_ShapeStyleLineProperties(ByRef $oDoc, ByRef $oObj[, $vStyle = Null[, $iColor = Null[, $iWidth = Null[, $iTransparency = Null[, $iCornerStyle = Null[, $iCapStyle = Null]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oObj                - [in/out] an object. A Shape Style or Presentation Style object returned by a previous _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $vStyle              - [optional] a variant value (0-31, or String). Default is Null. The Line Style to use. Can be a Custom Line Style name, or one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Line color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iWidth              - [optional] an integer value (0-5004). Default is Null. The line Width, set in Hundredths of a Millimeter (HMM).
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The Line transparency percentage. 100% = fully transparent.
;                  $iCornerStyle        - [optional] an integer value (0,2-4). Default is Null. The Line Corner Style. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  $iCapStyle           - [optional] an integer value (0-2). Default is Null. The Line Cap Style. See Constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = $vStyle not a String, and not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $vStyle is an Integer, but less than 0 or greater than 31. See constants $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 6 Return 0 = $iWidth not an Integer, less than 0 or greater than 5004.
;                  @Error 1 @Extended 7 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iCornerStyle not an Integer, not equal to 0, equal to 1, not equal to 2 or greater than 4. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 9 Return 0 = $iCapStyle is an Integer, but less than 0 or greater than 2. See constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to convert Constant to Line Style name.
;                  @Error 3 @Extended 2 Return 0 = Failed to insert Line Style name.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $vStyle
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $iWidth
;                  |                               8 = Error setting $iTransparency
;                  |                               16 = Error setting $iCornerStyle
;                  |                               32 = Error setting $iCapStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $vStyle accepts a String or an Integer because there is the possibility of a custom Line Style being available that the user may want to use.
;                  When retrieving the current settings, $vStyle could be either an Integer or a String. It will be a String if the current Line Style is a custom Line Style, else an Integer, corresponding to one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeStyleLineProperties(ByRef $oDoc, ByRef $oObj, $vStyle = Null, $iColor = Null, $iWidth = Null, $iTransparency = Null, $iCornerStyle = Null, $iCapStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local Const $__LOI_SHAPE_LINE_STYLE_NONE = 0, $__LOI_SHAPE_LINE_STYLE_SOLID = 1, $__LOI_SHAPE_LINE_STYLE_DASH = 2
	Local $avLine[6]
	Local $sStyle
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($vStyle, $iColor, $iWidth, $iTransparency, $iCornerStyle, $iCapStyle) Then
		Switch $oObj.LineStyle()
			Case $__LOI_SHAPE_LINE_STYLE_NONE
				$vReturn = $LOI_SHAPE_LINE_STYLE_NONE

			Case $__LOI_SHAPE_LINE_STYLE_SOLID
				$vReturn = $LOI_SHAPE_LINE_STYLE_CONTINUOUS

			Case $__LOI_SHAPE_LINE_STYLE_DASH
				$vReturn = __LOImpress_ShapeLineStyleName(Null, $oObj.LineDashName())
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		EndSwitch

		__LO_ArrayFill($avLine, $vReturn, $oObj.LineColor(), $oObj.LineWidth(), $oObj.LineTransparence(), $oObj.LineJoint(), $oObj.LineCap())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avLine)
	EndIf

	If ($vStyle <> Null) Then
		If Not IsString($vStyle) And Not IsInt($vStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If IsInt($vStyle) Then
			If Not __LO_IntIsBetween($vStyle, $LOI_SHAPE_LINE_STYLE_NONE, $LOI_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			Switch $vStyle
				Case $LOI_SHAPE_LINE_STYLE_NONE
					$oObj.LineStyle = $__LOI_SHAPE_LINE_STYLE_NONE
					$iError = ($oObj.LineStyle() = $__LOI_SHAPE_LINE_STYLE_NONE) ? ($iError) : (BitOR($iError, 1))

				Case $LOI_SHAPE_LINE_STYLE_CONTINUOUS
					$oObj.LineStyle = $__LOI_SHAPE_LINE_STYLE_SOLID
					$iError = ($oObj.LineStyle() = $__LOI_SHAPE_LINE_STYLE_SOLID) ? ($iError) : (BitOR($iError, 1))

				Case Else
					$sStyle = __LOImpress_ShapeLineStyleName($vStyle)
					If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					__LOImpress_ShapeLineDashNameInsert($oDoc, $vStyle)
					If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

					$oObj.LineStyle = $__LOI_SHAPE_LINE_STYLE_DASH
					$oObj.LineDashName = $sStyle
					$iError = ($oObj.LineDashName() = $sStyle) ? ($iError) : (BitOR($iError, 1))
			EndSwitch

		Else
			$sStyle = $vStyle
			$oObj.LineDashName = $sStyle
			$iError = ($oObj.LineDashName() = $sStyle) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.LineColor = $iColor
		$iError = ($oObj.LineColor() = $iColor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 0, 5004) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.LineWidth = $iWidth
		$iError = (__LO_IntIsBetween($oObj.LineWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.LineTransparence = $iTransparency
		$iError = ($oObj.LineTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iCornerStyle <> Null) Then
		If Not __LO_IntIsBetween($iCornerStyle, $LOI_SHAPE_LINE_JOINT_NONE, $LOI_SHAPE_LINE_JOINT_ROUND, $LOI_SHAPE_LINE_JOINT_MIDDLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.LineJoint = $iCornerStyle
		$iError = ($oObj.LineJoint() = $iCornerStyle) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iCapStyle <> Null) Then
		If Not __LO_IntIsBetween($iCapStyle, $LOI_SHAPE_LINE_CAP_FLAT, $LOI_SHAPE_LINE_CAP_SQUARE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oObj.LineCap = $iCapStyle
		$iError = ($oObj.LineCap() = $iCapStyle) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeStyleLineProperties

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeTextAttrAnimation
; Description ...: Set or Retrieve Shape or Shape Style Text Attribute Animation settings.
; Syntax ........: __LOImpress_ShapeTextAttrAnimation(ByRef $oObj[, $iEffect = Null[, $iDirection = Null[, $bStartInside = Null[, $bVisibleOnExit = Null[, $iCycles = Null[, $iInc = Null[, $bPixels = Null[, $iDelay = Null]]]]]]]])
; Parameters ....: $oObj                - [in/out] an object. A Shape or Shape Style object returned by a previous _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iEffect             - [optional] an integer value (0-4). Default is Null. The Animation type. See Constants, $LOI_ANIMATION_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDirection          - [optional] an integer value (0-3). Default is Null. The Direction of the text's movement, if applicable. See Constants, $LOI_ANIMATION_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bStartInside        - [optional] a boolean value. Default is Null. If True, Text is visible and inside the shape when the effect is applied.
;                  $bVisibleOnExit      - [optional] a boolean value. Default is Null. If True, Text remains visible after the effect is applied.
;                  $iCycles             - [optional] an integer value (0-100). Default is Null. The number of times to repeat the animation. 0 = Continuous.
;                  $iInc                - [optional] an integer value (1-100px/25-32766). Default is Null. the increment value for scrolling the text, in Hundredths of a Millimeter (HMM), or pixels.
;                  $bPixels             - [optional] a boolean value. Default is Null. If True, $iInc is set in pixels, else in Hundredths of a Millimeter (HMM).
;                  $iDelay              - [optional] an integer value (0-30000). Default is Null. The amount time (ms) to wait before repeating the effect.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iEffect not an Integer, less than 0 or greater than 4. See Constants, $LOI_ANIMATION_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iDirection not an Integer, less than 0 or greater than 3. See Constants, $LOI_ANIMATION_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $bStartInside not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bVisibleOnExit not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iCycles not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iInc not an Integer, less than 1 or greater than 100 pixels, less than 25 or greater than 32766 Hundredths of a Millimeter (HMM).
;                  @Error 1 @Extended 8 Return 0 = $bPixels not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iDelay not an Integer, less than 0 or greater than 30000.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current TextAnimationAmount value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iEffect
;                  |                               2 = Error setting $iDirection
;                  |                               4 = Error setting $bStartInside
;                  |                               8 = Error setting $bVisibleOnExit
;                  |                               16 = Error setting $iCycles
;                  |                               32 = Error setting $iInc
;                  |                               64 = Error setting $bPixels
;                  |                               128 = Error setting $iDelay
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeTextAttrAnimation(ByRef $oObj, $iEffect = Null, $iDirection = Null, $bStartInside = Null, $bVisibleOnExit = Null, $iCycles = Null, $iInc = Null, $bPixels = Null, $iDelay = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iValue
	Local $avTextAttr[8]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iEffect, $iDirection, $bStartInside, $bVisibleOnExit, $iCycles, $iInc, $bPixels, $iDelay) Then
		__LO_ArrayFill($avTextAttr, $oObj.TextAnimationKind(), $oObj.TextAnimationDirection(), $oObj.TextAnimationStartInside(), _
				$oObj.TextAnimationStopInside(), $oObj.TextAnimationCount(), _
				($oObj.TextAnimationAmount() < 0) ? ($oObj.TextAnimationAmount() * -1) : ($oObj.TextAnimationAmount()), _ ; If TextAnimationAmount is negative, Pixels are used, if positive Hundredths of a Millimeter (HMM).
				($oObj.TextAnimationAmount() < 0) ? (True) : (False), _ ; $bPixels
				$oObj.TextAnimationDelay())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTextAttr)
	EndIf

	If ($iEffect <> Null) Then
		If Not __LO_IntIsBetween($iEffect, $LOI_ANIMATION_TYPE_NONE, $LOI_ANIMATION_TYPE_SCROLL_IN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.TextAnimationKind = $iEffect
		$iError = ($oObj.TextAnimationKind() = $iEffect) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iDirection <> Null) Then
		If Not __LO_IntIsBetween($iDirection, $LOI_ANIMATION_DIR_LEFT, $LOI_ANIMATION_DIR_DOWN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.TextAnimationDirection = $iDirection
		$iError = ($oObj.TextAnimationDirection() = $iDirection) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bStartInside <> Null) Then
		If Not IsBool($bStartInside) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.TextAnimationStartInside = $bStartInside
		$iError = ($oObj.TextAnimationStartInside() = $bStartInside) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bVisibleOnExit <> Null) Then
		If Not IsBool($bVisibleOnExit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.TextAnimationStopInside = $bVisibleOnExit
		$iError = ($oObj.TextAnimationStopInside() = $bVisibleOnExit) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iCycles <> Null) Then
		If Not __LO_IntIsBetween($iCycles, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.TextAnimationCount = $iCycles
		$iError = ($oObj.TextAnimationCount() = $iCycles) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iInc <> Null) Then
		If (($oObj.TextAnimationAmount() < 0) And ($bPixels <> False)) Or ($bPixels = True) Then ; Set in Pixels
			If Not __LO_IntIsBetween($iInc, 1, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oObj.TextAnimationAmount = ($iInc * -1) ; Multiply by -1 to change to negative, as Pixels are set in negative values.
			$iError = ($oObj.TextAnimationAmount() = ($iInc * -1)) ? ($iError) : (BitOR($iError, 32))

		Else ; Set in Hundredths of a Millimeter (HMM).
			If Not __LO_IntIsBetween($iInc, 25, 32766) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oObj.TextAnimationAmount = $iInc
			$iError = ($oObj.TextAnimationAmount() = $iInc) ? ($iError) : (BitOR($iError, 32))
		EndIf
	EndIf

	If ($bPixels <> Null) Then
		If Not IsBool($bPixels) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$iValue = $oObj.TextAnimationAmount()
		If Not IsInt($iValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If $bPixels Then
			If ($iValue > 0) Then ; Not set to pixels yet, convert to pixels
				$iValue = ($iValue > 100) ? (-100) : ($iValue * -1) ; If greater than 100 pixels (the max) set to 100, otherwise convert the value to negative for pixels.
				$oObj.TextAnimationAmount = $iValue
			EndIf

		Else
			If ($iValue < 0) Then ; Set to pixels, convert to Hundredths of a Millimeter (HMM).
				$iValue = ($iValue * -1) ; Convert the value to positive for Hundredths of a Millimeter (HMM).
				$oObj.TextAnimationAmount = $iValue
			EndIf
		EndIf
		$iError = ($oObj.TextAnimationAmount() = $iValue) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 30000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oObj.TextAnimationDelay = $iDelay
		$iError = ($oObj.TextAnimationDelay() = $iDelay) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeTextAttrAnimation

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeTextAttrFit
; Description ...: Set or Retrieve Shape or Shape Style Text Attribute Fit properties. See Remarks.
; Syntax ........: __LOImpress_ShapeTextAttrFit(ByRef $oObj[, $bFitWidth = Null[, $bFitHeight = Null[, $bFitToFrame = Null[, $bAdjustContour = Null[, $bWordWrap = Null[, $bResizeShape = Null]]]]]])
; Parameters ....: $oObj                - [in/out] an object. A Shape or Shape Style object returned by a previous _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $bFitWidth           - [optional] a boolean value. Default is Null. If True, Expands the width of the object to the width of the text.
;                  $bFitHeight          - [optional] a boolean value. Default is Null. If True, Expands the height of the object to the height of the text.
;                  $bFitToFrame         - [optional] a boolean value. Default is Null. If True, Resizes the text to fit the entire area of the drawing object.
;                  $bAdjustContour      - [optional] a boolean value. Default is Null. If True, Adapts the text flow so that it matches the contours of the drawing object.
;                  $bWordWrap           - [optional] a boolean value. Default is Null. If True, Wraps the text to fit inside the shape.
;                  $bResizeShape        - [optional] a boolean value. Default is Null. If True, Resizes a custom shape to fit the text that you enter.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFitWidth not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bFitHeight not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bFitToFrame not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bAdjustContour not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bWordWrap not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bResizeShape not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFitWidth
;                  |                               2 = Error setting $bFitHeight
;                  |                               4 = Error setting $bFitToFrame
;                  |                               8 = Error setting $bAdjustContour
;                  |                               16 = Error setting $bWordWrap
;                  |                               32 = Error setting $bResizeShape
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The various "Fit" properties are grouped together here. The user will need to select the appropriate properties to set and analyze based on the shape that is called.
;                  All lines, including measure lines and dimension lines have the properties $bFitToFrame and $bAdjustContour.
;                  All Text Boxes have the properties $bFitToFrame, $bFitWidth and $bFitHeight.
;                  All DrawShapes, (Smileys, Rectangles, Callouts, Fontwork, etc.) excluding lines, have the properties $bWordWrap and $bResizeShape.
;                  For any other shapes, the user will need to determine which properties the shape has in the TextAttributes properties.
;                  Properties as found in the UI, and their equivalent: "Word Wrap Text in Shape" = $bWordWrap. "Resize Shape to Fit Text" = $bResizeShape. "Fit Width to Text" = $bFitWidth. "Fit Height to Text" = $bFitHeight. "Fit to Frame" = $bFitToFrame. "Adjust to Contour" = $bAdjustContour.
;                  The returned Array when retrieving properties will contain values for all of the parameters, and the user will need to determine which ones they need.
;                  When setting the properties for a shape, it is the user's responsibility to ensure the correct properties are used for the corresponding shape type.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeTextAttrFit(ByRef $oObj, $bFitWidth = Null, $bFitHeight = Null, $bFitToFrame = Null, $bAdjustContour = Null, $bWordWrap = Null, $bResizeShape = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__LOI_TEXT_FIT_NONE = 0, $__LOI_TEXT_FIT_PROP = 1 ; com.sun.star.drawing.TextFitToSizeType
	Local $iError = 0
	Local $avTextAttr[6]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bFitWidth, $bFitHeight, $bFitToFrame, $bAdjustContour, $bWordWrap, $bResizeShape) Then
		__LO_ArrayFill($avTextAttr, $oObj.TextAutoGrowWidth(), $oObj.TextAutoGrowHeight(), ($oObj.TextFitToSize() = $__LOI_TEXT_FIT_PROP) ? (True) : (False), _
				$oObj.TextContourFrame(), $oObj.TextWordWrap(), $oObj.TextAutoGrowHeight())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTextAttr)
	EndIf

	If ($bFitWidth <> Null) Then
		If Not IsBool($bFitWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.TextAutoGrowWidth = $bFitWidth
		$iError = ($oObj.TextAutoGrowWidth() = $bFitWidth) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bFitHeight <> Null) Then
		If Not IsBool($bFitHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.TextAutoGrowHeight = $bFitHeight
		$iError = ($oObj.TextAutoGrowHeight() = $bFitHeight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bFitToFrame <> Null) Then
		If Not IsBool($bFitToFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.TextFitToSize = ($bFitToFrame) ? ($__LOI_TEXT_FIT_PROP) : ($__LOI_TEXT_FIT_NONE)
		$iError = ($oObj.TextFitToSize() = ($bFitToFrame) ? ($__LOI_TEXT_FIT_PROP) : ($__LOI_TEXT_FIT_NONE)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bAdjustContour <> Null) Then
		If Not IsBool($bAdjustContour) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.TextContourFrame = $bAdjustContour
		$iError = ($oObj.TextContourFrame() = $bAdjustContour) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bWordWrap <> Null) Then
		If Not IsBool($bWordWrap) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.TextWordWrap = $bWordWrap
		$iError = ($oObj.TextWordWrap() = $bWordWrap) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bResizeShape <> Null) Then
		If Not IsBool($bResizeShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.TextAutoGrowHeight = $bResizeShape
		$iError = ($oObj.TextAutoGrowHeight() = $bResizeShape) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeTextAttrFit

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeTextAttrSettings
; Description ...: Set or Retrieve Shape, Shape Style or Presentation Style text Attribute settings.
; Syntax ........: __LOImpress_ShapeTextAttrSettings(ByRef $oObj[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null[, $iAnchor = Null[, $bFullWidth = Null]]]]]])
; Parameters ....: $oObj                - [in/out] an object. A Shape, Shape Style or Presentation Style object returned by a previous _LOImpress_DrawShapeInsert, _LOImpress_ShapesGetList, _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iLeft               - [optional] an integer value (-100000-100000). Default is Null. The space between the left edge of the drawing object and the left border of the text, in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value (-100000-100000). Default is Null. The space between the right edge of the drawing object and the right border of the text, in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value (-100000-100000). Default is Null. The space between the top edge of the drawing object and the top border of the text, in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value (-100000-100000). Default is Null. The space between the bottom edge of the drawing object and the bottom border of the text, in Hundredths of a Millimeter (HMM).
;                  $iAnchor             - [optional] an integer value (0-8). Default is Null. The text anchor position. See Constants, $LOI_PAR_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bFullWidth          - [optional] a boolean value. Default is Null. If True, Anchors the text to the full width of the drawing object.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iLeft not an Integer, less than -100000 or greater than 100000.
;                  @Error 1 @Extended 3 Return 0 = $iRight not an Integer, less than -100000 or greater than 100000.
;                  @Error 1 @Extended 4 Return 0 = $iTop not an Integer, less than -100000 or greater than 100000.
;                  @Error 1 @Extended 5 Return 0 = $iBottom not an Integer, less than -100000 or greater than 100000.
;                  @Error 1 @Extended 6 Return 0 = $iAnchor  not an Integer, less than 0 or greater than 8. See Constants, $LOI_PAR_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bFullWidth not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
;                  |                               16 = Error setting $iAnchor
;                  |                               32 = Error setting $bFullWidth
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeTextAttrSettings(ByRef $oObj, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null, $iAnchor = Null, $bFullWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurAnchor
	Local $avTextAttr[6]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLeft, $iRight, $iTop, $iBottom, $iAnchor, $bFullWidth) Then
		Select
			Case ($oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_TOP) And ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_LEFT)
				$iCurAnchor = $LOI_PAR_TEXT_ANCHOR_TOP_LEFT

			Case ($oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_TOP) And (($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_CENTER) Or ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_BLOCK))
				$iCurAnchor = $LOI_PAR_TEXT_ANCHOR_TOP_CENTER

			Case ($oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_TOP) And ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_RIGHT)
				$iCurAnchor = $LOI_PAR_TEXT_ANCHOR_TOP_RIGHT

			Case ($oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_CENTER) And ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_LEFT)
				$iCurAnchor = $LOI_PAR_TEXT_ANCHOR_MIDDLE_LEFT

			Case ($oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_CENTER) And (($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_CENTER) Or ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_BLOCK))
				$iCurAnchor = $LOI_PAR_TEXT_ANCHOR_MIDDLE_CENTER

			Case ($oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_CENTER) And ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_RIGHT)
				$iCurAnchor = $LOI_PAR_TEXT_ANCHOR_MIDDLE_RIGHT

			Case ($oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_BOTTOM) And ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_LEFT)
				$iCurAnchor = $LOI_PAR_TEXT_ANCHOR_BOTTOM_LEFT

			Case ($oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_BOTTOM) And (($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_CENTER) Or ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_BLOCK))
				$iCurAnchor = $LOI_PAR_TEXT_ANCHOR_BOTTOM_CENTER

			Case ($oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_BOTTOM) And ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_RIGHT)
				$iCurAnchor = $LOI_PAR_TEXT_ANCHOR_BOTTOM_RIGHT
		EndSelect

		__LO_ArrayFill($avTextAttr, $oObj.TextLeftDistance(), $oObj.TextRightDistance(), $oObj.TextUpperDistance(), $oObj.TextLowerDistance(), _
				$iCurAnchor, ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_BLOCK) ? (True) : (False))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTextAttr)
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.TextLeftDistance = $iLeft
		$iError = (__LO_IntIsBetween($oObj.TextLeftDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oObj.TextRightDistance = $iRight
		$iError = (__LO_IntIsBetween($oObj.TextRightDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.TextUpperDistance = $iTop
		$iError = (__LO_IntIsBetween($oObj.TextUpperDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.TextLowerDistance = $iBottom
		$iError = (__LO_IntIsBetween($oObj.TextLowerDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iAnchor <> Null) Then
		If Not __LO_IntIsBetween($iAnchor, $LOI_PAR_TEXT_ANCHOR_TOP_LEFT, $LOI_PAR_TEXT_ANCHOR_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		Switch $iAnchor
			Case $LOI_PAR_TEXT_ANCHOR_TOP_LEFT
				$oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_TOP
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_LEFT
				$iError = ($oObj.TextVerticalAdjust() = $LOI_PAR_TEXT_ALIGN_VERT_TOP) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_LEFT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_PAR_TEXT_ANCHOR_TOP_CENTER
				$oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_TOP
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_CENTER
				$iError = ($oObj.TextVerticalAdjust() = $LOI_PAR_TEXT_ALIGN_VERT_TOP) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_CENTER) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_PAR_TEXT_ANCHOR_TOP_RIGHT
				$oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_TOP
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_RIGHT
				$iError = ($oObj.TextVerticalAdjust() = $LOI_PAR_TEXT_ALIGN_VERT_TOP) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_RIGHT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_PAR_TEXT_ANCHOR_MIDDLE_LEFT
				$oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_CENTER
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_LEFT
				$iError = ($oObj.TextVerticalAdjust() = $LOI_PAR_TEXT_ALIGN_VERT_CENTER) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_LEFT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_PAR_TEXT_ANCHOR_MIDDLE_CENTER
				$oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_CENTER
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_CENTER
				$iError = ($oObj.TextVerticalAdjust() = $LOI_PAR_TEXT_ALIGN_VERT_CENTER) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_CENTER) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_PAR_TEXT_ANCHOR_MIDDLE_RIGHT
				$oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_CENTER
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_RIGHT
				$iError = ($oObj.TextVerticalAdjust() = $LOI_PAR_TEXT_ALIGN_VERT_CENTER) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_RIGHT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_PAR_TEXT_ANCHOR_BOTTOM_LEFT
				$oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_BOTTOM
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_LEFT
				$iError = ($oObj.TextVerticalAdjust() = $LOI_PAR_TEXT_ALIGN_VERT_BOTTOM) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_LEFT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_PAR_TEXT_ANCHOR_BOTTOM_CENTER
				$oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_BOTTOM
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_CENTER
				$iError = ($oObj.TextVerticalAdjust() = $LOI_PAR_TEXT_ALIGN_VERT_BOTTOM) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_CENTER) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_PAR_TEXT_ANCHOR_BOTTOM_RIGHT
				$oObj.TextVerticalAdjust = $LOI_PAR_TEXT_ALIGN_VERT_BOTTOM
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_RIGHT
				$iError = ($oObj.TextVerticalAdjust() = $LOI_PAR_TEXT_ALIGN_VERT_BOTTOM) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_RIGHT) ? ($iError) : (BitOR($iError, 16))
		EndSwitch
	EndIf

	If ($bFullWidth <> Null) Then
		If Not IsBool($bFullWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		If $bFullWidth Then
			$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_BLOCK
			$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_BLOCK) ? ($iError) : (BitOR($iError, 32))

		Else
			If ($oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_BLOCK) Then ; Only set Horizontal Adjust to Center if it was set to Block already when setting $bFullWidth to False.
				$oObj.TextHorizontalAdjust = $LOI_PAR_TEXT_ALIGN_HORI_CENTER
				$iError = ($oObj.TextHorizontalAdjust() = $LOI_PAR_TEXT_ALIGN_HORI_CENTER) ? ($iError) : (BitOR($iError, 32))
			EndIf
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_ShapeTextAttrSettings

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_StyleCharFontColor
; Description ...: Set or retrieve the font color and highlighting values.
; Syntax ........: __LOImpress_StyleCharFontColor(ByRef $oObj[, $iFontColor = Null[, $iHighlight = Null]])
; Parameters ....: $oObj                - [in/out] an object. A Shape Style or Presentation Style object returned by a previous _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, or _LOImpress_ShapePresStyleGetObjByName function.
;                  $iFontColor          - [optional] an integer value (-1-16777215). Default is Null. The font Color value, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iHighlight          - [optional] an integer value (-1-16777215). Default is Null. The highlight Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 3 Return 0 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $FontColor
;                  |                               2 = Error setting $iHighlight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_StyleCharFontColor(ByRef $oObj, $iFontColor = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avColor[2]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iFontColor, $iHighlight) Then
		__LO_ArrayFill($avColor, __LOImpress_ColorRemoveAlpha($oObj.CharColor()), $oObj.CharBackColor())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iFontColor <> Null) Then
		If Not __LO_IntIsBetween($iFontColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oObj.CharColor = $iFontColor
		$iError = ($oObj.CharColor() = $iFontColor) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iHighlight <> Null) Then
		If Not __LO_IntIsBetween($iHighlight, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		; CharHighlight; same as CharBackColor---Libre seems to use back color for highlighting however, so using that for setting.
;~ 		If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 2, 0)
;~ 		$oObj.CharHighlight = $iHighlight ;-- keeping old method in case.
;~ 		$iError = ($oObj.CharHighlight() = $iHighlight) ? ($iError) : (BitOR($iError, 2)
		$oObj.CharBackColor = $iHighlight
		$iError = ($oObj.CharBackColor() = $iHighlight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_StyleCharFontColor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_Transition
; Description ...: Set or Retrieve the current transition effect of a Slide.
; Syntax ........: __LOImpress_Transition(ByRef $oSlide[, $iTransition = Null])
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iTransition         - [optional] an integer value (0-78). Default is Null. The Transition effect. See Constants, $LOI_SLIDE_TRANSITION_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTransition not an Integer, less then 0 or greater than 78. See Constants, $LOI_SLIDE_TRANSITION_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Effect value.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Transition Type value.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve current Transition SubType value.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify current Transition type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTransition
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current Transition effect type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_Transition(ByRef $oSlide, $iTransition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iEffect, $iTransitionType, $iTransitionSubType, $iCurrTransition

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransition) Then
		$iEffect = $oSlide.Effect()
		If Not IsInt($iEffect) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$iTransitionType = $oSlide.TransitionType()
		If Not IsInt($iTransitionType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$iTransitionSubType = $oSlide.TransitionSubType()
		If Not IsInt($iTransitionSubType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Switch $iEffect ; Determine current Transition Type
			Case 0
				Switch $iTransitionType
					Case 0
						If ($iTransitionSubType = 0) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_NONE
						; $iEffect = 0 $iTransitionType = 0 $iTransitionSubType = 0

					Case 1
						If ($iTransitionSubType = 104) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_CUT_THROUGH_BLACK
						; $iEffect = 0 $iTransitionType = 1 $iTransitionSubType = 104

					Case 17
						Switch $iTransitionSubType
							Case 13
								$iCurrTransition = $LOI_SLIDE_TRANSITION_SHAPE_OVAL_VERT
								; $iEffect = 0 $iTransitionType = 17 $iTransitionSubType = 13

							Case 14
								$iCurrTransition = $LOI_SLIDE_TRANSITION_SHAPE_OVAL_HORI
								; $iEffect = 0 $iTransitionType = 17 $iTransitionSubType = 14
						EndSwitch

					Case 21
						Switch $iTransitionSubType
							Case 1
								$iCurrTransition = $LOI_SLIDE_TRANSITION_FALL
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 1

							Case 2
								$iCurrTransition = $LOI_SLIDE_TRANSITION_TURN_AROUND
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 2

							Case 3
								$iCurrTransition = $LOI_SLIDE_TRANSITION_IRIS
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 3

							Case 4
								$iCurrTransition = $LOI_SLIDE_TRANSITION_TURN_DOWN
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 4

							Case 5
								$iCurrTransition = $LOI_SLIDE_TRANSITION_ROCHADE
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 5

							Case 6
								$iCurrTransition = $LOI_SLIDE_TRANSITION_3D_VENETIAN_VERT
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 6

							Case 7
								$iCurrTransition = $LOI_SLIDE_TRANSITION_3D_VENETIAN_HORI
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 7

							Case 8
								$iCurrTransition = $LOI_SLIDE_TRANSITION_STATIC
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 8

							Case 9
								$iCurrTransition = $LOI_SLIDE_TRANSITION_FINE_DISSOLVE
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 9

							Case 11
								$iCurrTransition = $LOI_SLIDE_TRANSITION_CUBE_INSIDE
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 11

							Case 12
								$iCurrTransition = $LOI_SLIDE_TRANSITION_CUBE_OUTSIDE
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 12

							Case 13
								$iCurrTransition = $LOI_SLIDE_TRANSITION_VORTEX
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 13

							Case 14
								$iCurrTransition = $LOI_SLIDE_TRANSITION_RIPPLE
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 14

							Case 26
								$iCurrTransition = $LOI_SLIDE_TRANSITION_GLITTER
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 26

							Case 27
								$iCurrTransition = $LOI_SLIDE_TRANSITION_CIRCLES
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 27

							Case 31
								$iCurrTransition = $LOI_SLIDE_TRANSITION_HONEYCOMB
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 31

							Case 55
								$iCurrTransition = $LOI_SLIDE_TRANSITION_HELIX
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 55

							Case 108
								$iCurrTransition = $LOI_SLIDE_TRANSITION_TILES
								; $iEffect = 0 $iTransitionType = 21 $iTransitionSubType = 108
						EndSwitch

					Case 37
						If ($iTransitionSubType = 104) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_FADE_THROUGH_WHITE
						; $iEffect = 0 $iTransitionType = 37 $iTransitionSubType = 104
				EndSwitch

			Case 1
				If ($iTransitionType = 1) And ($iTransitionSubType = 1) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_WIPE_LEFT_TO_RIGHT
				; $iEffect = 1 $iTransitionType = 1 $iTransitionSubType = 1

			Case 2
				If ($iTransitionType = 1) And ($iTransitionSubType = 2) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_WIPE_TOP_TO_BOTTOM
				; $iEffect = 2 $iTransitionType = 1 $iTransitionSubType = 2

			Case 3
				If ($iTransitionType = 1) And ($iTransitionSubType = 1) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_WIPE_RIGHT_TO_LEFT
				; $iEffect = 3 $iTransitionType = 1 $iTransitionSubType = 1

			Case 4
				If ($iTransitionType = 1) And ($iTransitionSubType = 2) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_WIPE_BOTTOM_TO_TOP
				; $iEffect = 4 $iTransitionType = 1 $iTransitionSubType = 2

			Case 5
				If ($iTransitionType = 12) And ($iTransitionSubType = 25) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_BOX_IN
				; $iEffect = 5 $iTransitionType = 12 $iTransitionSubType = 25

			Case 6
				Switch $iTransitionType
					Case 3
						If ($iTransitionSubType = 12) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_SHAPE_PLUS
						; $iEffect = 6 $iTransitionType = 3 $iTransitionSubType = 12

					Case 12
						Switch $iTransitionSubType
							Case 25
								$iCurrTransition = $LOI_SLIDE_TRANSITION_BOX_OUT
								; $iEffect = 6 $iTransitionType = 12 $iTransitionSubType = 25

							Case 26
								$iCurrTransition = $LOI_SLIDE_TRANSITION_SHAPE_DIAMOND
								; $iEffect = 6 $iTransitionType = 12 $iTransitionSubType = 26
						EndSwitch

					Case 17
						If ($iTransitionSubType = 27) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_SHAPE_CIRCLE
						; $iEffect = 6 $iTransitionType = 17 $iTransitionSubType = 27
				EndSwitch

			Case 7
				If ($iTransitionType = 36) And ($iTransitionSubType = 97) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COVER_LEFT_TO_RIGHT
				; $iEffect = 7 $iTransitionType = 36 $iTransitionSubType = 97

			Case 8
				If ($iTransitionType = 36) And ($iTransitionSubType = 98) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COVER_TOP_TO_BOTTOM
				; $iEffect = 8 $iTransitionType = 36 $iTransitionSubType = 98

			Case 9
				If ($iTransitionType = 36) And ($iTransitionSubType = 99) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COVER_RIGHT_TO_LEFT
				; $iEffect = 9 $iTransitionType = 36 $iTransitionSubType = 99

			Case 10
				If ($iTransitionType = 36) And ($iTransitionSubType = 100) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COVER_BOTTOM_TO_TOP
				; $iEffect = 10 $iTransitionType = 36 $iTransitionSubType = 100

			Case 11
				If ($iTransitionType = 35) And ($iTransitionSubType = 97) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_PUSH_LEFT_TO_RIGHT
				; $iEffect = 11 $iTransitionType = 35 $iTransitionSubType = 97

			Case 12
				If ($iTransitionType = 35) And ($iTransitionSubType = 98) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_PUSH_TOP_TO_BOTTOM
				; $iEffect = 12 $iTransitionType = 35 $iTransitionSubType = 98

			Case 13
				If ($iTransitionType = 35) And ($iTransitionSubType = 98) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_PUSH_RIGHT_TO_LEFT
				; $iEffect = 13 $iTransitionType = 35 $iTransitionSubType = 99

			Case 14
				If ($iTransitionType = 35) And ($iTransitionSubType = 100) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_PUSH_BOTTOM_TO_TOP
				; $iEffect = 14 $iTransitionType = 35 $iTransitionSubType = 100

			Case 15
				If ($iTransitionType = 41) And ($iTransitionSubType = 13) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_VENETIAN_VERT
				; $iEffect = 15 $iTransitionType = 41 $iTransitionSubType = 13

			Case 16
				If ($iTransitionType = 41) And ($iTransitionSubType = 14) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_VENETIAN_HORI
				; $iEffect = 16 $iTransitionType = 41 $iTransitionSubType = 14

			Case 17
				Switch $iTransitionType
					Case 23
						Switch $iTransitionSubType
							Case 37
								$iCurrTransition = $LOI_SLIDE_TRANSITION_WHEEL_2_SPOKE
								; $iEffect = 17 $iTransitionType = 23 $iTransitionSubType = 37

							Case 39
								$iCurrTransition = $LOI_SLIDE_TRANSITION_WHEEL_4_SPOKE
								; $iEffect = 17 $iTransitionType = 23 $iTransitionSubType = 39

							Case 105
								$iCurrTransition = $LOI_SLIDE_TRANSITION_WHEEL_3_SPOKE
								; $iEffect = 17 $iTransitionType = 23 $iTransitionSubType = 105

							Case 106
								$iCurrTransition = $LOI_SLIDE_TRANSITION_WHEEL_8_SPOKE
								; $iEffect = 17 $iTransitionType = 23 $iTransitionSubType = 106

							Case 107
								$iCurrTransition = $LOI_SLIDE_TRANSITION_WHEEL_1_SPOKE
								; $iEffect = 17 $iTransitionType = 23 $iTransitionSubType = 107
						EndSwitch

					Case 25
						If ($iTransitionSubType = 48) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_WEDGE
						; $iEffect = 17 $iTransitionType = 25 $iTransitionSubType = 48

					Case 43
						If ($iTransitionSubType = 114) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_NEWSFLASH
						; $iEffect = 17 $iTransitionType = 43 $iTransitionSubType = 114
				EndSwitch

			Case 19
				If ($iTransitionType = 34) And ($iTransitionSubType = 95) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_DIAGONAL_TOP_LEFT_TO_BOTTOM_RIGHT
				; $iEffect = 19 $iTransitionType = 34 $iTransitionSubType = 95

			Case 20
				If ($iTransitionType = 34) And ($iTransitionSubType = 96) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_DIAGONAL_TOP_RIGHT_TO_BOTTOM_LEFT
				; $iEffect = 20 $iTransitionType = 34 $iTransitionSubType = 96

			Case 21
				If ($iTransitionType = 34) And ($iTransitionSubType = 96) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_DIAGONAL_BOTTOM_LEFT_TO_TOP_RIGHT
				; $iEffect = 21 $iTransitionType = 34 $iTransitionSubType = 96

			Case 22
				If ($iTransitionType = 34) And ($iTransitionSubType = 95) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_DIAGONAL_BOTTOM_RIGHT_TO_TOP_LEFT
				; $iEffect = 22 $iTransitionType = 34 $iTransitionSubType = 95

			Case 23
				If ($iTransitionType = 4) And ($iTransitionSubType = 14) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_SPLIT_HORI_IN
				; $iEffect = 23 $iTransitionType = 4 $iTransitionSubType = 14

			Case 24
				If ($iTransitionType = 4) And ($iTransitionSubType = 13) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_SPLIT_VERT_IN
				; $iEffect = 24 $iTransitionType = 4 $iTransitionSubType = 13

			Case 25
				If ($iTransitionType = 4) And ($iTransitionSubType = 14) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_SPLIT_HORI_OUT
				; $iEffect = 25 $iTransitionType = 4 $iTransitionSubType = 14

			Case 26
				If ($iTransitionType = 4) And ($iTransitionSubType = 13) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_SPLIT_VERT_OUT
				; $iEffect = 26 $iTransitionType = 4 $iTransitionSubType = 13

			Case 31
				Switch $iTransitionType
					Case 37
						Switch $iTransitionSubType
							Case 101
								$iCurrTransition = $LOI_SLIDE_TRANSITION_FADE_SMOOTHLY
								; $iEffect = 31 $iTransitionType = 37 $iTransitionSubType = 101

							Case 104
								$iCurrTransition = $LOI_SLIDE_TRANSITION_FADE_THROUGH_BLACK
								; $iEffect = 31 $iTransitionType = 37 $iTransitionSubType = 104
						EndSwitch

					Case 40
						If ($iTransitionSubType = 0) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_DISSOLVE
						; $iEffect = 31 $iTransitionType = 40 $iTransitionSubType = 0
				EndSwitch

			Case 36
				If ($iTransitionType = 42) And ($iTransitionSubType = 0) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_RANDOM
				; $iEffect = 36 $iTransitionType = 42 $iTransitionSubType = 0

			Case 41
				Switch $iTransitionType
					Case 35
						If ($iTransitionSubType = 111) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COMB_VERT
						; $iEffect = 41 $iTransitionType = 35 $iTransitionSubType = 111

					Case 38
						If ($iTransitionSubType = 13) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_BARS_VERT
						; $iEffect = 41 $iTransitionType = 38 $iTransitionSubType = 13
				EndSwitch

			Case 42
				Switch $iTransitionType
					Case 35
						If ($iTransitionSubType = 110) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COMB_HORI
						; $iEffect = 42 $iTransitionType = 35 $iTransitionSubType = 110

					Case 38
						If ($iTransitionSubType = 14) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_BARS_HORI
						; $iEffect = 42 $iTransitionType = 38 $iTransitionSubType = 14
				EndSwitch

			Case 43
				If ($iTransitionType = 36) And ($iTransitionSubType = 116) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COVER_TOP_LEFT_TO_BOTTOM_RIGHT
				; $iEffect = 43 $iTransitionType = 36 $iTransitionSubType = 116

			Case 44
				If ($iTransitionType = 36) And ($iTransitionSubType = 117) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COVER_TOP_RIGHT_TO_BOTTOM_LEFT
				; $iEffect = 44 $iTransitionType = 36 $iTransitionSubType = 117

			Case 45
				If ($iTransitionType = 36) And ($iTransitionSubType = 119) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COVER_BOTTOM_RIGHT_TO_TOP_LEFT
				; $iEffect = 45 $iTransitionType = 36 $iTransitionSubType = 119

			Case 46
				If ($iTransitionType = 36) And ($iTransitionSubType = 118) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_COVER_BOTTOM_LEFT_TO_TOP_RIGHT
				; $iEffect = 46 $iTransitionType = 36 $iTransitionSubType = 118

			Case 47
				If ($iTransitionType = 36) And ($iTransitionSubType = 99) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_UNCOVER_RIGHT_TO_LEFT
				; $iEffect = 47 $iTransitionType = 36 $iTransitionSubType = 99

			Case 48
				If ($iTransitionType = 36) And ($iTransitionSubType = 119) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_UNCOVER_BOTTOM_RIGHT_TO_TOP_LEFT
				; $iEffect = 48 $iTransitionType = 36 $iTransitionSubType = 119

			Case 49
				If ($iTransitionType = 36) And ($iTransitionSubType = 100) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_UNCOVER_BOTTOM_TO_TOP
				; $iEffect = 49 $iTransitionType = 36 $iTransitionSubType = 100

			Case 50
				If ($iTransitionType = 36) And ($iTransitionSubType = 118) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_UNCOVER_BOTTOM_LEFT_TO_TOP_RIGHT
				; $iEffect = 50 $iTransitionType = 36 $iTransitionSubType = 118

			Case 51
				If ($iTransitionType = 36) And ($iTransitionSubType = 97) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_UNCOVER_LEFT_TO_RIGHT
				; $iEffect = 51 $iTransitionType = 36 $iTransitionSubType = 97

			Case 52
				If ($iTransitionType = 36) And ($iTransitionSubType = 116) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_UNCOVER_TOP_LEFT_TO_BOTTOM_RIGHT
				; $iEffect = 52 $iTransitionType = 36 $iTransitionSubType = 116

			Case 53
				If ($iTransitionType = 36) And ($iTransitionSubType = 98) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_UNCOVER_TOP_TO_BOTTOM
				; $iEffect = 53 $iTransitionType = 36 $iTransitionSubType = 98

			Case 54
				If ($iTransitionType = 36) And ($iTransitionSubType = 117) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_UNCOVER_TOP_RIGHT_TO_BOTTOM_LEFT
				; $iEffect = 54 $iTransitionType = 36 $iTransitionSubType = 117

			Case 55
				If ($iTransitionType = 39) And ($iTransitionSubType = 19) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_CHECKERS_DOWN
				; $iEffect = 55 $iTransitionType = 39 $iTransitionSubType = 19

			Case 56
				If ($iTransitionType = 39) And ($iTransitionSubType = 108) Then $iCurrTransition = $LOI_SLIDE_TRANSITION_CHECKERS_ACROSS
				; $iEffect = 56 $iTransitionType = 39 $iTransitionSubType = 108
		EndSwitch

		If Not IsInt($iCurrTransition) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurrTransition)
	EndIf

	If Not __LO_IntIsBetween($iTransition, $LOI_SLIDE_TRANSITION_3D_VENETIAN_VERT, $LOI_SLIDE_TRANSITION_WIPE_TOP_TO_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch $iTransition
		Case $LOI_SLIDE_TRANSITION_3D_VENETIAN_VERT
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 6

		Case $LOI_SLIDE_TRANSITION_3D_VENETIAN_HORI
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 7

		Case $LOI_SLIDE_TRANSITION_BARS_VERT
			$iEffect = 41
			$iTransitionType = 38
			$iTransitionSubType = 13

		Case $LOI_SLIDE_TRANSITION_BARS_HORI
			$iEffect = 42
			$iTransitionType = 38
			$iTransitionSubType = 14

		Case $LOI_SLIDE_TRANSITION_BOX_OUT
			$iEffect = 6
			$iTransitionType = 12
			$iTransitionSubType = 25

		Case $LOI_SLIDE_TRANSITION_BOX_IN
			$iEffect = 5
			$iTransitionType = 12
			$iTransitionSubType = 25

		Case $LOI_SLIDE_TRANSITION_CHECKERS_DOWN
			$iEffect = 55
			$iTransitionType = 39
			$iTransitionSubType = 19

		Case $LOI_SLIDE_TRANSITION_CHECKERS_ACROSS
			$iEffect = 56
			$iTransitionType = 39
			$iTransitionSubType = 108

		Case $LOI_SLIDE_TRANSITION_CIRCLES
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 27

		Case $LOI_SLIDE_TRANSITION_COMB_HORI
			$iEffect = 42
			$iTransitionType = 35
			$iTransitionSubType = 110

		Case $LOI_SLIDE_TRANSITION_COMB_VERT
			$iEffect = 41
			$iTransitionType = 35
			$iTransitionSubType = 111

		Case $LOI_SLIDE_TRANSITION_COVER_TOP_TO_BOTTOM
			$iEffect = 8
			$iTransitionType = 36
			$iTransitionSubType = 98

		Case $LOI_SLIDE_TRANSITION_COVER_RIGHT_TO_LEFT
			$iEffect = 9
			$iTransitionType = 36
			$iTransitionSubType = 99

		Case $LOI_SLIDE_TRANSITION_COVER_LEFT_TO_RIGHT
			$iEffect = 7
			$iTransitionType = 36
			$iTransitionSubType = 97

		Case $LOI_SLIDE_TRANSITION_COVER_BOTTOM_TO_TOP
			$iEffect = 10
			$iTransitionType = 36
			$iTransitionSubType = 100

		Case $LOI_SLIDE_TRANSITION_COVER_TOP_RIGHT_TO_BOTTOM_LEFT
			$iEffect = 44
			$iTransitionType = 36
			$iTransitionSubType = 117

		Case $LOI_SLIDE_TRANSITION_COVER_BOTTOM_RIGHT_TO_TOP_LEFT
			$iEffect = 45
			$iTransitionType = 36
			$iTransitionSubType = 119

		Case $LOI_SLIDE_TRANSITION_COVER_TOP_LEFT_TO_BOTTOM_RIGHT
			$iEffect = 43
			$iTransitionType = 36
			$iTransitionSubType = 116

		Case $LOI_SLIDE_TRANSITION_COVER_BOTTOM_LEFT_TO_TOP_RIGHT
			$iEffect = 46
			$iTransitionType = 36
			$iTransitionSubType = 118

		Case $LOI_SLIDE_TRANSITION_CUBE_OUTSIDE
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 12

		Case $LOI_SLIDE_TRANSITION_CUBE_INSIDE
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 11

		Case $LOI_SLIDE_TRANSITION_CUT_THROUGH_BLACK
			$iEffect = 0
			$iTransitionType = 1
			$iTransitionSubType = 104

		Case $LOI_SLIDE_TRANSITION_DIAGONAL_TOP_RIGHT_TO_BOTTOM_LEFT
			$iEffect = 20
			$iTransitionType = 34
			$iTransitionSubType = 96

		Case $LOI_SLIDE_TRANSITION_DIAGONAL_BOTTOM_RIGHT_TO_TOP_LEFT
			$iEffect = 22
			$iTransitionType = 34
			$iTransitionSubType = 95

		Case $LOI_SLIDE_TRANSITION_DIAGONAL_TOP_LEFT_TO_BOTTOM_RIGHT
			$iEffect = 19
			$iTransitionType = 34
			$iTransitionSubType = 95

		Case $LOI_SLIDE_TRANSITION_DIAGONAL_BOTTOM_LEFT_TO_TOP_RIGHT
			$iEffect = 21
			$iTransitionType = 34
			$iTransitionSubType = 96

		Case $LOI_SLIDE_TRANSITION_DISSOLVE
			$iEffect = 31
			$iTransitionType = 40
			$iTransitionSubType = 0

		Case $LOI_SLIDE_TRANSITION_FADE_THROUGH_BLACK
			$iEffect = 31
			$iTransitionType = 37
			$iTransitionSubType = 104

		Case $LOI_SLIDE_TRANSITION_FADE_THROUGH_WHITE
			$iEffect = 0
			$iTransitionType = 37
			$iTransitionSubType = 104

		Case $LOI_SLIDE_TRANSITION_FADE_SMOOTHLY
			$iEffect = 31
			$iTransitionType = 37
			$iTransitionSubType = 101

		Case $LOI_SLIDE_TRANSITION_FALL
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 1

		Case $LOI_SLIDE_TRANSITION_FINE_DISSOLVE
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 9

		Case $LOI_SLIDE_TRANSITION_GLITTER
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 26

		Case $LOI_SLIDE_TRANSITION_HELIX
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 55

		Case $LOI_SLIDE_TRANSITION_HONEYCOMB
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 31

		Case $LOI_SLIDE_TRANSITION_IRIS
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 3

		Case $LOI_SLIDE_TRANSITION_NEWSFLASH
			$iEffect = 17
			$iTransitionType = 43
			$iTransitionSubType = 114

		Case $LOI_SLIDE_TRANSITION_NONE
			$iEffect = 0
			$iTransitionType = 0
			$iTransitionSubType = 0

		Case $LOI_SLIDE_TRANSITION_PUSH_TOP_TO_BOTTOM
			$iEffect = 12
			$iTransitionType = 35
			$iTransitionSubType = 98

		Case $LOI_SLIDE_TRANSITION_PUSH_RIGHT_TO_LEFT
			$iEffect = 13
			$iTransitionType = 35
			$iTransitionSubType = 99

		Case $LOI_SLIDE_TRANSITION_PUSH_LEFT_TO_RIGHT
			$iEffect = 11
			$iTransitionType = 35
			$iTransitionSubType = 97

		Case $LOI_SLIDE_TRANSITION_PUSH_BOTTOM_TO_TOP
			$iEffect = 14
			$iTransitionType = 35
			$iTransitionSubType = 100

		Case $LOI_SLIDE_TRANSITION_RANDOM
			$iEffect = 36
			$iTransitionType = 42
			$iTransitionSubType = 0

		Case $LOI_SLIDE_TRANSITION_RIPPLE
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 14

		Case $LOI_SLIDE_TRANSITION_ROCHADE
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 5

		Case $LOI_SLIDE_TRANSITION_SHAPE_PLUS
			$iEffect = 6
			$iTransitionType = 3
			$iTransitionSubType = 12

		Case $LOI_SLIDE_TRANSITION_SHAPE_DIAMOND
			$iEffect = 6
			$iTransitionType = 12
			$iTransitionSubType = 26

		Case $LOI_SLIDE_TRANSITION_SHAPE_CIRCLE
			$iEffect = 6
			$iTransitionType = 17
			$iTransitionSubType = 27

		Case $LOI_SLIDE_TRANSITION_SHAPE_OVAL_HORI
			$iEffect = 0
			$iTransitionType = 17
			$iTransitionSubType = 14

		Case $LOI_SLIDE_TRANSITION_SHAPE_OVAL_VERT
			$iEffect = 0
			$iTransitionType = 17
			$iTransitionSubType = 13

		Case $LOI_SLIDE_TRANSITION_SPLIT_HORI_IN
			$iEffect = 23
			$iTransitionType = 4
			$iTransitionSubType = 14

		Case $LOI_SLIDE_TRANSITION_SPLIT_HORI_OUT
			$iEffect = 25
			$iTransitionType = 4
			$iTransitionSubType = 14

		Case $LOI_SLIDE_TRANSITION_SPLIT_VERT_IN
			$iEffect = 24
			$iTransitionType = 4
			$iTransitionSubType = 13

		Case $LOI_SLIDE_TRANSITION_SPLIT_VERT_OUT
			$iEffect = 26
			$iTransitionType = 4
			$iTransitionSubType = 13

		Case $LOI_SLIDE_TRANSITION_STATIC
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 8

		Case $LOI_SLIDE_TRANSITION_TILES
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 108

		Case $LOI_SLIDE_TRANSITION_TURN_AROUND
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 2

		Case $LOI_SLIDE_TRANSITION_TURN_DOWN
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 4

		Case $LOI_SLIDE_TRANSITION_UNCOVER_TOP_TO_BOTTOM
			$iEffect = 53
			$iTransitionType = 36
			$iTransitionSubType = 98

		Case $LOI_SLIDE_TRANSITION_UNCOVER_RIGHT_TO_LEFT
			$iEffect = 47
			$iTransitionType = 36
			$iTransitionSubType = 99

		Case $LOI_SLIDE_TRANSITION_UNCOVER_LEFT_TO_RIGHT
			$iEffect = 51
			$iTransitionType = 36
			$iTransitionSubType = 97

		Case $LOI_SLIDE_TRANSITION_UNCOVER_BOTTOM_TO_TOP
			$iEffect = 49
			$iTransitionType = 36
			$iTransitionSubType = 100

		Case $LOI_SLIDE_TRANSITION_UNCOVER_TOP_RIGHT_TO_BOTTOM_LEFT
			$iEffect = 54
			$iTransitionType = 36
			$iTransitionSubType = 117

		Case $LOI_SLIDE_TRANSITION_UNCOVER_BOTTOM_RIGHT_TO_TOP_LEFT
			$iEffect = 48
			$iTransitionType = 36
			$iTransitionSubType = 119

		Case $LOI_SLIDE_TRANSITION_UNCOVER_TOP_LEFT_TO_BOTTOM_RIGHT
			$iEffect = 52
			$iTransitionType = 36
			$iTransitionSubType = 116

		Case $LOI_SLIDE_TRANSITION_UNCOVER_BOTTOM_LEFT_TO_TOP_RIGHT
			$iEffect = 50
			$iTransitionType = 36
			$iTransitionSubType = 118

		Case $LOI_SLIDE_TRANSITION_VENETIAN_VERT
			$iEffect = 15
			$iTransitionType = 41
			$iTransitionSubType = 13

		Case $LOI_SLIDE_TRANSITION_VENETIAN_HORI
			$iEffect = 16
			$iTransitionType = 41
			$iTransitionSubType = 14

		Case $LOI_SLIDE_TRANSITION_VORTEX
			$iEffect = 0
			$iTransitionType = 21
			$iTransitionSubType = 13

		Case $LOI_SLIDE_TRANSITION_WEDGE
			$iEffect = 17
			$iTransitionType = 25
			$iTransitionSubType = 48

		Case $LOI_SLIDE_TRANSITION_WHEEL_1_SPOKE
			$iEffect = 17
			$iTransitionType = 23
			$iTransitionSubType = 107

		Case $LOI_SLIDE_TRANSITION_WHEEL_2_SPOKE
			$iEffect = 17
			$iTransitionType = 23
			$iTransitionSubType = 37

		Case $LOI_SLIDE_TRANSITION_WHEEL_3_SPOKE
			$iEffect = 17
			$iTransitionType = 23
			$iTransitionSubType = 105

		Case $LOI_SLIDE_TRANSITION_WHEEL_4_SPOKE
			$iEffect = 17
			$iTransitionType = 23
			$iTransitionSubType = 39

		Case $LOI_SLIDE_TRANSITION_WHEEL_8_SPOKE
			$iEffect = 17
			$iTransitionType = 23
			$iTransitionSubType = 106

		Case $LOI_SLIDE_TRANSITION_WIPE_BOTTOM_TO_TOP
			$iEffect = 4
			$iTransitionType = 1
			$iTransitionSubType = 2

		Case $LOI_SLIDE_TRANSITION_WIPE_LEFT_TO_RIGHT
			$iEffect = 1
			$iTransitionType = 1
			$iTransitionSubType = 1

		Case $LOI_SLIDE_TRANSITION_WIPE_RIGHT_TO_LEFT
			$iEffect = 3
			$iTransitionType = 1
			$iTransitionSubType = 1

		Case $LOI_SLIDE_TRANSITION_WIPE_TOP_TO_BOTTOM
			$iEffect = 2
			$iTransitionType = 1
			$iTransitionSubType = 2
	EndSwitch

	$oSlide.Effect = $iEffect
	$oSlide.TransitionType = $iTransitionType
	$oSlide.TransitionSubType = $iTransitionSubType

	$iError = (($oSlide.Effect() = $iEffect) And _
			($oSlide.TransitionType() = $iTransitionType) And _
			($oSlide.TransitionSubType() = $iTransitionSubType)) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOImpress_Transition

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_TransparencyGradientConvert
; Description ...: Convert a Transparency Gradient percentage value to a color value or from a color value to a percentage.
; Syntax ........: __LOImpress_TransparencyGradientConvert([$iPercentToLong = Null[, $iLongToPercent = Null]])
; Parameters ....: $iPercentToLong      - [optional] an integer value. Default is Null. The percentage to convert to a RGB Color Integer.
;                  $iLongToPercent      - [optional] an integer value. Default is Null. The RGB Color Integer to convert to percentage.
; Return values .: Success: Integer.
;                  Failure: Null and sets the @Error and @Extended flags to non-zero.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return Null = No values called in parameters.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. The requested Integer value converted from percentage to a RGB Color Integer.
;                  @Error 0 @Extended 1 Return Integer = Success. The requested Integer value from a RGB Color Integer to percentage.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_TransparencyGradientConvert($iPercentToLong = Null, $iLongToPercent = Null)
	Local $iReturn

	If ($iPercentToLong <> Null) Then
		$iReturn = ((255 * ($iPercentToLong / 100)) + .50) ; Change percentage to decimal and times by White color (255 RGB) Add . 50 to round up if applicable.
		$iReturn = _LO_ConvertColorToLong(Int($iReturn), Int($iReturn), Int($iReturn))

		Return SetError($__LO_STATUS_SUCCESS, 0, $iReturn)

	ElseIf ($iLongToPercent <> Null) Then
		$iReturn = _LO_ConvertColorFromLong(Null, $iLongToPercent)
		$iReturn = Int((($iReturn[0] / 255) * 100) + .50) ; All return color values will be the same, so use only one. Add . 50 to round up if applicable.

		Return SetError($__LO_STATUS_SUCCESS, 1, $iReturn)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, Null)
	EndIf
EndFunc   ;==>__LOImpress_TransparencyGradientConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_TransparencyGradientNameInsert
; Description ...: Create and insert a new Transparency Gradient name.
; Syntax ........: __LOImpress_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $tTGradient          - a dll struct value. A Gradient Structure to copy settings from.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tTGradient not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.TransparencyGradientTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error creating Transparency Gradient Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. A new transparency Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Transparency Gradient name is blank, I need to create a new name and apply it. I think I could re-use an old one without problems, but I'm not sure, so to be safe, I will create a new one.
;                  If there are no names that have been already created, then I need to create and apply one before the transparency gradient will be displayed.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tNewTGradient
	Local $oTGradTable
	Local $iCount = 1
	Local $sGradient = "com.sun.star.awt.Gradient2"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tTGradient) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not __LO_VersionCheck(7.6) Then $sGradient = "com.sun.star.awt.Gradient"

	$oTGradTable = $oDoc.createInstance("com.sun.star.drawing.TransparencyGradientTable")
	If Not IsObj($oTGradTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oTGradTable.hasByName("Transparency " & $iCount)
		$iCount += 1
		Sleep((IsInt($iCount / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
	WEnd

	$tNewTGradient = __LO_CreateStruct($sGradient)
	If Not IsObj($tNewTGradient) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; Copy the settings over from the input Style Gradient to my new one. This may not be necessary? But just in case.
	With $tNewTGradient
		.Style = $tTGradient.Style()
		.XOffset = $tTGradient.XOffset()
		.YOffset = $tTGradient.YOffset()
		.Angle = $tTGradient.Angle()
		.Border = $tTGradient.Border()
		.StartColor = $tTGradient.StartColor()
		.EndColor = $tTGradient.EndColor()

		If __LO_VersionCheck(7.6) Then .ColorStops = $tTGradient.ColorStops()
	EndWith

	$oTGradTable.insertByName("Transparency " & $iCount, $tNewTGradient)
	If Not ($oTGradTable.hasByName("Transparency " & $iCount)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, "Transparency " & $iCount)
EndFunc   ;==>__LOImpress_TransparencyGradientNameInsert
