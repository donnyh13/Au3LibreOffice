#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
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
#include "LibreOfficeWriter_Par.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Applying L.O. Writer Page Styles.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_PageStyleAreaColor
; _LOWriter_PageStyleAreaFillStyle
; _LOWriter_PageStyleAreaGradient
; _LOWriter_PageStyleAreaGradientMulticolor
; _LOWriter_PageStyleAreaTransparency
; _LOWriter_PageStyleAreaTransparencyGradient
; _LOWriter_PageStyleAreaTransparencyGradientMulti
; _LOWriter_PageStyleBorderColor
; _LOWriter_PageStyleBorderPadding
; _LOWriter_PageStyleBorderStyle
; _LOWriter_PageStyleBorderWidth
; _LOWriter_PageStyleColumnSeparator
; _LOWriter_PageStyleColumnSettings
; _LOWriter_PageStyleColumnSize
; _LOWriter_PageStyleCreate
; _LOWriter_PageStyleDelete
; _LOWriter_PageStyleExists
; _LOWriter_PageStyleFooter
; _LOWriter_PageStyleFooterAreaColor
; _LOWriter_PageStyleFooterAreaFillStyle
; _LOWriter_PageStyleFooterAreaGradient
; _LOWriter_PageStyleFooterAreaGradientMulticolor
; _LOWriter_PageStyleFooterAreaTransparency
; _LOWriter_PageStyleFooterAreaTransparencyGradient
; _LOWriter_PageStyleFooterAreaTransparencyGradientMulti
; _LOWriter_PageStyleFooterBorderColor
; _LOWriter_PageStyleFooterBorderPadding
; _LOWriter_PageStyleFooterBorderStyle
; _LOWriter_PageStyleFooterBorderWidth
; _LOWriter_PageStyleFooterShadow
; _LOWriter_PageStyleFootnoteArea
; _LOWriter_PageStyleFootnoteLine
; _LOWriter_PageStyleGetObj
; _LOWriter_PageStyleHeader
; _LOWriter_PageStyleHeaderAreaColor
; _LOWriter_PageStyleHeaderAreaFillStyle
; _LOWriter_PageStyleHeaderAreaGradient
; _LOWriter_PageStyleHeaderAreaGradientMulticolor
; _LOWriter_PageStyleHeaderAreaTransparency
; _LOWriter_PageStyleHeaderAreaTransparencyGradient
; _LOWriter_PageStyleHeaderAreaTransparencyGradientMulti
; _LOWriter_PageStyleHeaderBorderColor
; _LOWriter_PageStyleHeaderBorderPadding
; _LOWriter_PageStyleHeaderBorderStyle
; _LOWriter_PageStyleHeaderBorderWidth
; _LOWriter_PageStyleHeaderShadow
; _LOWriter_PageStyleLayout
; _LOWriter_PageStyleMargins
; _LOWriter_PageStyleOrganizer
; _LOWriter_PageStylePaperFormat
; _LOWriter_PageStyleSet
; _LOWriter_PageStylesGetNames
; _LOWriter_PageStyleShadow
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleAreaColor
; Description ...: Set or Retrieve background color settings for a Page style.
; Syntax ........: _LOWriter_PageStyleAreaColor(ByRef $oPageStyle[, $iBackColor = Null[, $bBackTransparent = Null]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The color to make the background. Set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for "None".
;                  $bBackTransparent    - [optional] a boolean value. Default is Null. If True, the background color is transparent.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iBackColor not an integer, less than -1, or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $bBackTransparent not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  |                               2 = Error setting $bBackTransparent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleAreaColor(ByRef $oPageStyle, $iBackColor = Null, $bBackTransparent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency
	Local $avColor[2]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iBackColor, $bBackTransparent) Then
		__LO_ArrayFill($avColor, __LOWriter_ColorRemoveAlpha($oPageStyle.BackColor()), $oPageStyle.BackTransparent())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$iOldTransparency = $oPageStyle.FillTransparence()
		If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oPageStyle.BackColor = $iBackColor
		$iError = ($oPageStyle.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))

		$oPageStyle.FillTransparence = $iOldTransparency
	EndIf

	If ($bBackTransparent <> Null) Then
		If Not IsBool($bBackTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.BackTransparent = $bBackTransparent
		$iError = ($oPageStyle.BackTransparent() = $bBackTransparent) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOWriter_PageStyleAreaFillStyle(ByRef $oPageStyle)
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Fill Style.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning current background fill style. Return will be one of the constants $LOW_AREA_FILL_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleAreaFillStyle(ByRef $oPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oPageStyle.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOWriter_PageStyleAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleAreaGradient
; Description ...: Modify or retrieve the settings for Page Style Background color Gradient.
; Syntax ........: _LOWriter_PageStyleAreaGradient(ByRef $oDoc, ByRef $oPageStyle[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See Constants, $LOW_GRAD_NAME_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient that you want to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iIncrement          - [optional] an integer value (0,3-256). Default is Null. Specifies the number of steps of change color. 0 = Automatic.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] an integer value (0-16777215). Default is Null. A color for the beginning point of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] an integer value (0-16777215). Default is Null. A color for the endpoint of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 4 Return 0 = $sGradientName Not a String.
;                  @Error 1 @Extended 5 Return 0 = $iType Not an Integer, less than -1, or greater than 5. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iIncrement Not an Integer, less than 3 but not 0, or greater than 256.
;                  @Error 1 @Extended 7 Return 0 = $iXCenter Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iYCenter Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iAngle Not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 10 Return 0 = $iTransitionStart Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 11 Return 0 = $iFromColor Not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 12 Return 0 = $iToColor Not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 13 Return 0 = $iFromIntense Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 14 Return 0 = $iToIntense Not an Integer, less than 0, or greater than 100.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 11 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleAreaGradient(ByRef $oDoc, ByRef $oPageStyle, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName
	Local $atColorStop[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tStyleGradient = $oPageStyle.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oPageStyle.FillGradientName(), $tStyleGradient.Style(), _
				$oPageStyle.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oPageStyle.FillStyle() <> $LOW_AREA_FILL_STYLE_GRADIENT) Then $oPageStyle.FillStyle = $LOW_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		__LOWriter_GradientPresets($oDoc, $oPageStyle, $tStyleGradient, $sGradientName)
		$iError = ($oPageStyle.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oPageStyle.FillStyle = $LOW_AREA_FILL_STYLE_OFF

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oPageStyle.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iFromColor <> Null) Then
		If Not __LO_IntIsBetween($iFromColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

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
		If Not __LO_IntIsBetween($iToColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

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
		If Not __LO_IntIsBetween($iFromIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$tStyleGradient.StartIntensity = $iFromIntense
	EndIf

	If ($iToIntense <> Null) Then
		If Not __LO_IntIsBetween($iToIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$tStyleGradient.EndIntensity = $iToIntense
	EndIf

	If ($oPageStyle.FillGradientName() = "") Then
		$sGradName = __LOWriter_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oPageStyle.FillGradientName = $sGradName
		If ($oPageStyle.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oPageStyle.FillGradient = $tStyleGradient

	; Error checking
	$iError = ($iType = Null) ? ($iError) : (($oPageStyle.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oPageStyle.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oPageStyle.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oPageStyle.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oPageStyle.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = ($iFromColor = Null) ? ($iError) : (($oPageStyle.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = ($iToColor = Null) ? ($iError) : (($oPageStyle.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = ($iFromIntense = Null) ? ($iError) : (($oPageStyle.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = ($iToIntense = Null) ? ($iError) : (($oPageStyle.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleAreaGradientMulticolor
; Description ...: Set or Retrieve a Page Style's Multicolor Gradient settings. See remarks.
; Syntax ........: _LOWriter_PageStyleAreaGradientMulticolor(ByRef $oPageStyle[, $avColorStops = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 4 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 5 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 6 Return ? = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
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
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were set to Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, in Long integer format.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOWriter_GradientMulticolorAdd, _LOWriter_GradientMulticolorDelete, _LOWriter_GradientMulticolorModify, _LOWriter_PageStyleAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleAreaGradientMulticolor(ByRef $oPageStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.FillGradient()
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
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)

		$tStopColor.Red = (BitAND(BitShift($avColorStops[$i][1], 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($avColorStops[$i][1], 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($avColorStops[$i][1], 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oPageStyle.FillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oPageStyle.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleAreaTransparency
; Description ...: Modify or retrieve Transparency settings for a page style.
; Syntax ........: _LOWriter_PageStyleAreaTransparency(ByRef $oPageStyle[, $iTransparency = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTransparency not an Integer, less than 0, or greater than 100.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current setting for Transparency as an integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleAreaTransparency(ByRef $oPageStyle, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oPageStyle.FillTransparence())

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPageStyle.FillTransparenceGradientName = ""
	$oPageStyle.FillTransparence = $iTransparency
	$iError = ($oPageStyle.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleAreaTransparencyGradient
; Description ...: Modify or retrieve the transparency gradient settings.
; Syntax ........: _LOWriter_PageStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oPageStyle[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 4 Return 0 = $iType not an Integer, less than -1, or greater than 5, see constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iXCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iYCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iAngle not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 8 Return 0 = $iTransitionStart not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iStart not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 10 Return 0 = $iEnd not an Integer, less than 0, or greater than 100.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oPageStyle, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop
	Local $fValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tStyleGradient = $oPageStyle.FillTransparenceGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		__LO_ArrayFill($aiTransparent, $tStyleGradient.Style(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), _
				Int($tStyleGradient.Angle() / 10), $tStyleGradient.Border(), __LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.StartColor()), _
				__LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oPageStyle.FillTransparenceGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iStart <> Null) Then
		If Not __LO_IntIsBetween($iStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tStyleGradient.StartColor = __LOWriter_TransparencyGradientConvert($iStart)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tColorStop = $atColorStop[0] ; StopOffset 0 is the "Start" Value.

			$tStopColor = $tColorStop.StopColor()

			$fValue = $iStart / 100 ; Value is a decimal percentage value.

			$tStopColor.Red = $fValue
			$tStopColor.Green = $fValue
			$tStopColor.Blue = $fValue

			$tColorStop.StopColor = $tStopColor

			$atColorStop[0] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($iEnd <> Null) Then
		If Not __LO_IntIsBetween($iEnd, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tStyleGradient.EndColor = __LOWriter_TransparencyGradientConvert($iEnd)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; StopOffset 0 is the "End" Value.

			$tStopColor = $tColorStop.StopColor()

			$fValue = $iEnd / 100 ; Value is a decimal percentage value.

			$tStopColor.Red = $fValue
			$tStopColor.Green = $fValue
			$tStopColor.Blue = $fValue

			$tColorStop.StopColor = $tStopColor

			$atColorStop[UBound($atColorStop) - 1] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($oPageStyle.FillTransparenceGradientName() = "") Then
		$sTGradName = __LOWriter_TransparencyGradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oPageStyle.FillTransparenceGradientName = $sTGradName
		If ($oPageStyle.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oPageStyle.FillTransparenceGradient = $tStyleGradient

	$iError = ($iType = Null) ? ($iError) : (($oPageStyle.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oPageStyle.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oPageStyle.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oPageStyle.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oPageStyle.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iStart = Null) ? ($iError) : (($oPageStyle.FillTransparenceGradient.StartColor() = __LOWriter_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iEnd = Null) ? ($iError) : (($oPageStyle.FillTransparenceGradient.EndColor() = __LOWriter_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Page Style's Multi Transparency Gradient settings. See remarks.
; Syntax ........: _LOWriter_PageStyleAreaTransparencyGradientMulti(ByRef $oPageStyle[, $avColorStops = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 4 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 5 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 6 Return ? = ColorStop color not an Integer, less than 0 or greater than 100. Returning problem element index.
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
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were set to Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple Transparency stops in a Gradient rather than just a beginning and an ending value, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the Transparency percentage value between 0 and 100%.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOWriter_TransparencyGradientMultiModify, _LOWriter_TransparencyGradientMultiDelete, _LOWriter_TransparencyGradientMultiAdd, _LOWriter_PageStyleAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleAreaTransparencyGradientMulti(ByRef $oPageStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.FillTransparenceGradient()
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
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)

		$tStopColor.Red = ($avColorStops[$i][1] / 100)
		$tStopColor.Green = ($avColorStops[$i][1] / 100)
		$tStopColor.Blue = ($avColorStops[$i][1] / 100)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oPageStyle.FillTransparenceGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oPageStyle.FillTransparenceGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleBorderColor
; Description ...: Set the Page Style Border Line Color. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_PageStyleBorderColor(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. Set the Top Border Line Color of the Page in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. Set the Bottom Border Line Color of the Page in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. Set the Left Border Line Color of the Page in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. Set the Right Border Line Color of the Page in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0, or greater than 16,777,215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Color when Right Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_PageStyleBorderWidth, _LOWriter_PageStyleBorderStyle, _LOWriter_PageStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleBorderColor(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oPageStyle, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_PageStyleBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleBorderPadding
; Description ...: Set or retrieve the Page Style Border Padding settings.
; Syntax ........: _LOWriter_PageStyleBorderPadding(ByRef $oPageStyle[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Micrometers (uM).
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Distance between the Border and Page contents in Micrometers(uM).
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Distance between the Border and Page contents in Micrometers(uM).
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Distance between the Border and Page contents in Micrometers(uM).
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Distance between the Border and Page contents in Micrometers(uM).
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer, _LOWriter_PageStyleBorderWidth, _LOWriter_PageStyleBorderStyle, _LOWriter_PageStyleBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleBorderPadding(ByRef $oPageStyle, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_PageStyleBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleBorderStyle
; Description ...: Set or Retrieve the Page Style Border Line style. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_PageStyleBorderStyle(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Top Border Line Style of the Page using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Bottom Border Line Style of the Page using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Left Border Line Style of the Page using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Right Border Line Style of the Page using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style when Right Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LOWriter_PageStyleBorderWidth, _LOWriter_PageStyleBorderColor, _LOWriter_PageStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleBorderStyle(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oPageStyle, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_PageStyleBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleBorderWidth
; Description ...: Set or Retrieve the Page Style Border Line Width. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_PageStyleBorderWidth(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Border Line width of the Page in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Border Line Width of the Page in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Border Line width of the Page in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Border Line Width of the Page in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set Width to 0
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer, _LOWriter_PageStyleBorderStyle, _LOWriter_PageStyleBorderColor, _LOWriter_PageStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleBorderWidth(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oPageStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_PageStyleBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleColumnSeparator
; Description ...: Modify or retrieve Page Style Column Separator line settings.
; Syntax ........: _LOWriter_PageStyleColumnSeparator(ByRef $oPageStyle[, $bSeparatorOn = Null[, $iStyle = Null[, $iWidth = Null[, $iColor = Null[, $iHeight = Null[, $iPosition = Null]]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $bSeparatorOn        - [optional] a boolean value. Default is Null. If true, add a separator line between two or more columns.
;                  $iStyle              - [optional] an integer value (0-3). Default is Null. The formatting style for the column separator line. See Constants, $LOW_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iWidth              - [optional] an integer value (5-180). Default is Null. The width of the separator line. Set in Micrometers.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The color for the separator line. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iHeight             - [optional] an integer value (0-100). Default is Null. The length of the separator line as a percentage of the height of the column area.
;                  $iPosition           - [optional] an integer value (0-2). Default is Null. The vertical alignment of the separator line. This option is only available if Height value of the line is less than 100%. See Constants, $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $bSeparatorOn not a Boolean value.
;                  @Error 1 @Extended 4 Return 0 = $iStyle not an Integer, less than 0, or greater than 3. See constants, $LOW_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iWidth not an Integer, less than 5 or greater than 180.
;                  @Error 1 @Extended 6 Return 0 = $iColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 7 Return 0 = $iHeight not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iPosition not an Integer, less than 0, or greater than 2. See constants, $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Text Columns Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bSeparatorOn
;                  |                               2 = Error setting $iStyle
;                  |                               4 = Error setting $iWidth
;                  |                               8 = Error setting $iColor
;                  |                               16 = Error setting $iHeight
;                  |                               32 = Error setting $iPosition
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleColumnSeparator(ByRef $oPageStyle, $bSeparatorOn = Null, $iStyle = Null, $iWidth = Null, $iColor = Null, $iHeight = Null, $iPosition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns
	Local $iError = 0
	Local $avColumnLine[6]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextColumns = $oPageStyle.TextColumns()
	If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bSeparatorOn, $iStyle, $iWidth, $iColor, $iHeight, $iPosition) Then
		__LO_ArrayFill($avColumnLine, $oTextColumns.SeparatorLineIsOn(), $oTextColumns.SeparatorLineStyle(), $oTextColumns.SeparatorLineWidth(), _
				$oTextColumns.SeparatorLineColor(), $oTextColumns.SeparatorLineRelativeHeight(), $oTextColumns.SeparatorLineVerticalAlignment())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColumnLine)
	EndIf

	If ($bSeparatorOn <> Null) Then
		If Not IsBool($bSeparatorOn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTextColumns.SeparatorLineIsOn = $bSeparatorOn
		$iError = ($oTextColumns.SeparatorLineIsOn() = $bSeparatorOn) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iStyle <> Null) Then
		If Not __LO_IntIsBetween($iStyle, $LOW_LINE_STYLE_NONE, $LOW_LINE_STYLE_DASHED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTextColumns.SeparatorLineStyle = $iStyle
		$iError = ($oTextColumns.SeparatorLineStyle() = $iStyle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 5, 180) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTextColumns.SeparatorLineWidth = $iWidth
		$iError = (__LO_IntIsBetween($oTextColumns.SeparatorLineWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTextColumns.SeparatorLineColor = $iColor
		$iError = ($oTextColumns.SeparatorLineColor() = $iColor) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iHeight <> Null) Then
		If Not __LO_IntIsBetween($iHeight, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oTextColumns.SeparatorLineRelativeHeight = $iHeight
		$iError = ($oTextColumns.SeparatorLineRelativeHeight() = $iHeight) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iPosition <> Null) Then
		If Not __LO_IntIsBetween($iPosition, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oTextColumns.SeparatorLineVerticalAlignment = $iPosition
		$iError = ($oTextColumns.SeparatorLineVerticalAlignment() = $iPosition) ? ($iError) : (BitOR($iError, 32))
	EndIf

	$oPageStyle.TextColumns = $oTextColumns

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleColumnSeparator

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleColumnSettings
; Description ...: Modify or retrieve page style Column count.
; Syntax ........: _LOWriter_PageStyleColumnSettings(ByRef $oPageStyle[, $iColumns = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iColumns            - [optional] an integer value. Default is Null. The number of columns that you want in the page. Minimum 1.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iColumns not an Integer or less than 1.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Text Columns Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iColumns
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current column count.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleColumnSettings(ByRef $oPageStyle, $iColumns = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns
	Local $iError = 0

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextColumns = $oPageStyle.TextColumns()
	If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iColumns) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oTextColumns.ColumnCount())

	If Not __LO_IntIsBetween($iColumns, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextColumns.ColumnCount = $iColumns
	$oPageStyle.TextColumns = $oTextColumns

	$iError = ($oPageStyle.TextColumns.ColumnCount() = $iColumns) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleColumnSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleColumnSize
; Description ...: Modify or retrieve Column sizing settings. See remarks.
; Syntax ........: _LOWriter_PageStyleColumnSize(ByRef $oPageStyle, $iColumn[, $bAutoWidth = Null[, $iGlobalSpacing = Null[, $iSpacing = Null[, $iWidth = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iColumn             - an integer value. The column to modify the settings on. See Remarks.
;                  $bAutoWidth          - [optional] a boolean value. Default is Null. If True, Column Width is automatically adjusted.
;                  $iGlobalSpacing      - [optional] an integer value. Default is Null. Set a spacing value for between all columns. Set in Micrometers. See remarks.
;                  $iSpacing            - [optional] an integer value. Default is Null. The Space between two columns, in Micrometers. Cannot be set for the last column.
;                  $iWidth              - [optional] an integer value. Default is Null. If $iGlobalSpacing is set to other than 0, enter the width of the column. Set in Micrometers.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iColumn not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iColumn higher than number of columns in the document or less than 1.
;                  @Error 1 @Extended 5 Return 0 = $bAutoWidth not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iGlobalSpacing not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iSpacing not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iWidth not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Text Columns Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Page Style Column Object Array.
;                  @Error 3 @Extended 3 Return 0 = No columns present for requested Page Style.
;                  @Error 3 @Extended 4 Return 0 = No columns present for requested Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoWidth
;                  |                               2 = Error setting $iGlobalSpacing
;                  |                               4 = Error setting $iSpacing
;                  |                               8 = Error setting $iWidth
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work fine for setting AutoWidth, and Spacing values, however Width will not work the best, Spacing etc is set in plain micrometer values, however width is set in a relative value, and I am unable to find a way to be able to convert a specific value, such as 1" (2540 Micrometers) etc, to the appropriate relative value, especially when spacing is set.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  To set $bAutoWidth or $iGlobalSpacing you may enter any number in $iColumn as long as you are not setting width or spacing, as AutoWidth is not column specific. If you set a value for $iGlobalSpacing with $bAutoWidth set to false, the value is applied to all the columns still.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleColumnSize(ByRef $oPageStyle, $iColumn, $bAutoWidth = Null, $iGlobalSpacing = Null, $iSpacing = Null, $iWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns
	Local $atColumns
	Local $iError = 0, $iRightMargin, $iLeftMargin
	Local $avColumnSize[4]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextColumns = $oPageStyle.TextColumns()
	If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$atColumns = $oTextColumns.Columns()
	If Not IsArray($atColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($oTextColumns.ColumnCount() <= 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If ($iColumn > UBound($atColumns)) Or ($iColumn < 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iColumn = $iColumn - 1 ; Libre Columns Array is 0 based -- Minus one to compensate

	If __LO_VarsAreNull($bAutoWidth, $iGlobalSpacing, $iSpacing, $iWidth) Then
		If ($iColumn = (UBound($atColumns) - 1)) Then ; If last column is called, there is no spacing value, so return the outer margin, which will be 0.
			__LO_ArrayFill($avColumnSize, $oTextColumns.IsAutomatic, $oTextColumns.AutomaticDistance(), _
					$atColumns[$iColumn].RightMargin(), $atColumns[$iColumn].Width())

		Else
			__LO_ArrayFill($avColumnSize, $oTextColumns.IsAutomatic, $oTextColumns.AutomaticDistance(), _
					$atColumns[$iColumn].RightMargin() + $atColumns[$iColumn + 1].LeftMargin(), $atColumns[$iColumn].Width())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColumnSize)
	EndIf

	If ($bAutoWidth <> Null) Then
		If Not IsBool($bAutoWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		If ($bAutoWidth <> $oTextColumns.IsAutomatic()) Then ; If Auto Width not already the same setting, then modify it.

			If ($bAutoWidth = True) Then
				; retrieve both outside column inner margin settings to add together for determining AutoWidth value.
				$iGlobalSpacing = ($iGlobalSpacing = Null) ? ($atColumns[0].RightMargin() + $atColumns[UBound($atColumns) - 1].LeftMargin()) : ($iGlobalSpacing)
				; If $iGlobalSpacing is not called with a value, set my own, else use the called value.

				$oTextColumns.ColumnCount = $oTextColumns.ColumnCount()
				$oPageStyle.TextColumns = $oTextColumns
				; Setting the number of columns activates the AutoWidth option, so set it to the same number of columns.

			Else ; If False
				; If GlobalSpacing isn't set, then set it myself to the current automatic distance.
				$iGlobalSpacing = ($iGlobalSpacing = Null) ? ($oTextColumns.AutomaticDistance()) : ($iGlobalSpacing)
				$oTextColumns.setColumns($atColumns) ; Inserting the Column Array(Sequence) again, even without changes, deactivates AutoWidth.
			EndIf
		EndIf

		$oPageStyle.TextColumns = $oTextColumns
		$iError = ($oPageStyle.TextColumns.IsAutomatic() = $bAutoWidth) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iGlobalSpacing <> Null) Then
		If Not IsInt($iGlobalSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTextColumns.AutomaticDistance = $iGlobalSpacing
		$oPageStyle.TextColumns = $oTextColumns

		If ($oPageStyle.TextColumns.IsAutomatic() = True) Then ; If AutoWidth is on (True) Then error test, else don't, because I use $iGlobalSpacing
			; for setting the width internally also.
			$iError = (__LO_IntIsBetween($oPageStyle.TextColumns.AutomaticDistance(), $iGlobalSpacing - 2, $iGlobalSpacing + 2)) ? ($iError) : (BitOR($iError, 2))
		EndIf
	EndIf

	If ($iSpacing <> Null) Then
		If Not IsInt($iSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		If ($iColumn = (UBound($atColumns) - 1)) Then ; If the requested column is the last column (furthest right), then set property setting error.
			; because spacing can't be set for the last column.
			$iError = BitOR($iError, 4)

		Else
			; Spacing is equally divided between the two adjoining columns, so set the first columns right margin,
			; and the next column's left margin to half of the spacing value each.
			$iRightMargin = Int($iSpacing / 2)
			$atColumns[$iColumn].RightMargin = $iRightMargin

			$iLeftMargin = Int($iSpacing - ($iSpacing / 2))
			$atColumns[$iColumn + 1].LeftMargin = $iLeftMargin

			; Set the settings into the document.
			$oTextColumns.setColumns($atColumns)
			$oPageStyle.TextColumns = $oTextColumns

			; Retrieve Array of columns again for testing.
			$atColumns = $oTextColumns.Columns()
			If Not IsArray($atColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			; See if setting spacing worked. Spacing is equally divided between the two adjoining columns, so retrieve the first columns right
			; margin, and the next column's left margin.
			$iError = (__LO_IntIsBetween($atColumns[$iColumn].RightMargin() + $atColumns[$iColumn + 1].LeftMargin(), $iSpacing - 1, $iSpacing + 1)) ? ($iError) : (BitOR($iError, 4))
		EndIf
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$atColumns[$iColumn].Width = $iWidth

		; Set the settings into the document.
		$oTextColumns.setColumns($atColumns)
		$oPageStyle.TextColumns = $oTextColumns

		; Retrieve Array of columns again for testing.
		$atColumns = $oPageStyle.TextColumns.Columns()
		If Not IsArray($atColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$iError = ($iWidth = Null) ? ($iError) : ((__LO_IntIsBetween($atColumns[$iColumn].Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 8)))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleColumnSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleCreate
; Description ...: Create a new Page Style in a Document.
; Syntax ........: _LOWriter_PageStyleCreate(ByRef $oDoc, $sPageStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Related .......: _LOWriter_PageStyleDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleCreate(ByRef $oDoc, $sPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageStyles, $oStyle, $oPageStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oPageStyles = $oDoc.StyleFamilies().getByName("PageStyles")
	If Not IsObj($oPageStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If _LOWriter_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oStyle = $oDoc.createInstance("com.sun.star.style.PageStyle")
	If Not IsObj($oStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oPageStyles.insertByName($sPageStyle, $oStyle)

	If Not $oPageStyles.hasByName($sPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oPageStyle = $oPageStyles.getByName($sPageStyle)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPageStyle)
EndFunc   ;==>_LOWriter_PageStyleCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleDelete
; Description ...: Delete a User-Created Page Style from a Document.
; Syntax ........: _LOWriter_PageStyleDelete(ByRef $oDoc, $oPageStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function. Must be User-Created, not a built-in Style native to LibreOffice.
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
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleDelete(ByRef $oDoc, ByRef $oPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_PageStyleDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleExists
; Description ...: Check whether a document contains the requested Page Style by Name.
; Syntax ........: _LOWriter_PageStyleExists(ByRef $oDoc, $sPageStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
Func _LOWriter_PageStyleExists(ByRef $oDoc, $sPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oDoc.StyleFamilies.getByName("PageStyles").hasByName($sPageStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOWriter_PageStyleExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooter
; Description ...: Modify or retrieve Footer settings for a page style.
; Syntax ........: _LOWriter_PageStyleFooter(ByRef $oPageStyle[, $bFooterOn = Null[, $bSameLeftRight = Null[, $bSameOnFirst = Null[, $iLeftMargin = Null[, $iRightMargin = Null[, $iSpacing = Null[, $bDynamicSpacing = Null[, $iHeight = Null[, $bAutoHeight = Null]]]]]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $bFooterOn           - [optional] a boolean value. Default is Null. If True, adds a footer to the page style.
;                  $bSameLeftRight      - [optional] a boolean value. Default is Null. If True, Even and odd pages share the same content.
;                  $bSameOnFirst        - [optional] a boolean value. Default is Null. If True, First and even/odd pages share the same content. LibreOffice 4.0 and up.
;                  $iLeftMargin         - [optional] an integer value. Default is Null. The amount of space to leave between the left edge of the page and the left edge of the footer. Set in Micrometers.
;                  $iRightMargin        - [optional] an integer value. Default is Null. The amount of space to leave between the right edge of the page and the right edge of the footer. Set in Micrometers.
;                  $iSpacing            - [optional] an integer value. Default is Null. The amount of space that you want to maintain between the bottom edge of the document text and the top edge of the footer. Set in Micrometers.
;                  $bDynamicSpacing     - [optional] a boolean value. Default is Null. If True, Overrides the Spacing setting and allows the footer to expand into the area between the footer and document text.
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the footer. Set in Micrometers.
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
;                  @Error 1 @Extended 9 Return 0 = $bDynamicSpacing not a Boolean value.
;                  @Error 1 @Extended 10 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 11 Return 0 = $bAutoHeight not a Boolean value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bFooterOn
;                  |                               2 = Error setting $bSameLeftRight
;                  |                               4 = Error setting $bSameOnFirst
;                  |                               8 = Error setting $iLeftMargin
;                  |                               16 = Error setting $iRightMargin
;                  |                               32 = Error setting $iSpacing
;                  |                               64 = Error setting $bDynamicSpacing
;                  |                               128 = Error setting $iHeight
;                  |                               256 = Error setting $bAutoHeight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 or 9 Element Array with values in order of function parameters. If Libre Office version is less than 4.0, then the Array returned will contain 8 elements, because $bSameOnFirst will not be available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooter(ByRef $oPageStyle, $bFooterOn = Null, $bSameLeftRight = Null, $bSameOnFirst = Null, $iLeftMargin = Null, $iRightMargin = Null, $iSpacing = Null, $bDynamicSpacing = Null, $iHeight = Null, $bAutoHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFooter[8]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bFooterOn, $bSameLeftRight, $bSameOnFirst, $iLeftMargin, $iRightMargin, $iSpacing, $bDynamicSpacing, $iHeight, $bAutoHeight) Then
		If __LO_VersionCheck(4.0) Then
			__LO_ArrayFill($avFooter, $oPageStyle.FooterIsOn(), $oPageStyle.FooterIsShared(), $oPageStyle.FirstIsShared(), $oPageStyle.FooterLeftMargin(), _
					$oPageStyle.FooterRightMargin(), $oPageStyle.FooterBodyDistance(), $oPageStyle.FooterDynamicSpacing(), $oPageStyle.FooterHeight(), _
					$oPageStyle.FooterIsDynamicHeight())

		Else
			__LO_ArrayFill($avFooter, $oPageStyle.FooterIsOn(), $oPageStyle.FooterIsShared(), $oPageStyle.FooterLeftMargin(), _
					$oPageStyle.FooterRightMargin(), $oPageStyle.FooterBodyDistance(), $oPageStyle.FooterDynamicSpacing(), $oPageStyle.FooterHeight(), _
					$oPageStyle.FooterIsDynamicHeight())
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

		$oPageStyle.FirstIsShared = $bSameOnFirst
		$iError = ($oPageStyle.FirstIsShared() = $bSameOnFirst) ? ($iError) : (BitOR($iError, 4))
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

	If ($bDynamicSpacing <> Null) Then
		If Not IsBool($bDynamicSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oPageStyle.FooterDynamicSpacing = $bDynamicSpacing
		$iError = ($oPageStyle.FooterDynamicSpacing() = $bDynamicSpacing) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iHeight <> Null) Then
		If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oPageStyle.FooterHeight = $iHeight
		$iError = (__LO_IntIsBetween($oPageStyle.FooterHeight(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bAutoHeight <> Null) Then
		If Not IsBool($bAutoHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oPageStyle.FooterIsDynamicHeight = $bAutoHeight
		$iError = ($oPageStyle.FooterIsDynamicHeight() = $bAutoHeight) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterAreaColor
; Description ...: Set or Retrieve background color settings for a Page style Footer.
; Syntax ........: _LOWriter_PageStyleFooterAreaColor(ByRef $oPageStyle[, $iBackColor = Null[, $bBackTransparent = Null]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The background color. Set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for "None".
;                  $bBackTransparent    - [optional] a boolean value. Default is Null. If True, the background color is transparent.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iBackColor not an integer, less than -1, or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $bBackTransparent not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Footers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  |                               2 = Error setting $bBackTransparent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterAreaColor(ByRef $oPageStyle, $iBackColor = Null, $bBackTransparent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency
	Local $avColor[2]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iBackColor, $bBackTransparent) Then
		__LO_ArrayFill($avColor, __LOWriter_ColorRemoveAlpha($oPageStyle.FooterBackColor()), $oPageStyle.FooterBackTransparent())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$iOldTransparency = $oPageStyle.FooterFillTransparence()
		If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oPageStyle.FooterBackColor = $iBackColor
		$iError = ($oPageStyle.FooterBackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))

		$oPageStyle.FooterFillTransparence = $iOldTransparency
	EndIf

	If ($bBackTransparent <> Null) Then
		If Not IsBool($bBackTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.FooterBackTransparent = $bBackTransparent
		$iError = ($oPageStyle.FooterBackTransparent() = $bBackTransparent) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFooterAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOWriter_PageStyleFooterAreaFillStyle(ByRef $oPageStyle)
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Fill Style.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning current background fill style. Return will be one of the constants $LOW_AREA_FILL_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterAreaFillStyle(ByRef $oPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oPageStyle.FooterFillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOWriter_PageStyleFooterAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterAreaGradient
; Description ...: Modify or retrieve the settings for Page Style Footer Background color Gradient.
; Syntax ........: _LOWriter_PageStyleFooterAreaGradient(ByRef $oDoc, ByRef $oPageStyle[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See Constants, $LOW_GRAD_NAME_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient that you want to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iIncrement          - [optional] an integer value (0,3-256). Default is Null. Specifies the number of steps of change color. 0 = Automatic.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] an integer value (0-16777215). Default is Null. A color for the beginning point of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] an integer value (0-16777215). Default is Null. A color for the endpoint of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 4 Return 0 = $sGradientName not a String.
;                  @Error 1 @Extended 5 Return 0 = $iType Not an Integer, less than -1, or greater than 5. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error 1 @Extended 7 Return 0 = $iXCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iYCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iAngle not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 10 Return 0 = $iTransitionStart not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 11 Return 0 = $iFromColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 12 Return 0 = $iToColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 13 Return 0 = $iFromIntense not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 14 Return 0 = $iToIntense not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Footers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving "FillGradient" Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 4 Return 0 = Error creating Gradient Name.
;                  @Error 3 @Extended 5 Return 0 = Error setting Gradient Name.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 11 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterAreaGradient(ByRef $oDoc, ByRef $oPageStyle, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName
	Local $atColorStop[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.FooterFillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oPageStyle.FooterFillGradientName(), $tStyleGradient.Style(), _
				$oPageStyle.FooterFillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oPageStyle.FooterFillStyle() <> $LOW_AREA_FILL_STYLE_GRADIENT) Then $oPageStyle.FooterFillStyle = $LOW_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		__LOWriter_GradientPresets($oDoc, $oPageStyle, $tStyleGradient, $sGradientName, True)
		$iError = ($oPageStyle.FooterFillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oPageStyle.FooterFillStyle = $LOW_AREA_FILL_STYLE_OFF

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.FooterFillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oPageStyle.FooterFillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iFromColor <> Null) Then
		If Not __LO_IntIsBetween($iFromColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tStyleGradient.StartColor = $iFromColor

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

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
		If Not __LO_IntIsBetween($iToColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$tStyleGradient.EndColor = $iToColor

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

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
		If Not __LO_IntIsBetween($iFromIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$tStyleGradient.StartIntensity = $iFromIntense
	EndIf

	If ($iToIntense <> Null) Then
		If Not __LO_IntIsBetween($iToIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$tStyleGradient.EndIntensity = $iToIntense
	EndIf

	If ($oPageStyle.FooterFillGradientName = "") Then
		$sGradName = __LOWriter_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oPageStyle.FooterFillGradientName = $sGradName
		If ($oPageStyle.FooterFillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	$oPageStyle.FooterFillGradient = $tStyleGradient

	; Error checking
	$iError = ($iType = Null) ? ($iError) : (($oPageStyle.FooterFillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oPageStyle.FooterFillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oPageStyle.FooterFillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oPageStyle.FooterFillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oPageStyle.FooterFillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = ($iFromColor = Null) ? ($iError) : (($oPageStyle.FooterFillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = ($iToColor = Null) ? ($iError) : (($oPageStyle.FooterFillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = ($iFromIntense = Null) ? ($iError) : (($oPageStyle.FooterFillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = ($iToIntense = Null) ? ($iError) : (($oPageStyle.FooterFillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFooterAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterAreaGradientMulticolor
; Description ...: Set or Retrieve a Page Style's Footer Multicolor Gradient settings. See remarks.
; Syntax ........: _LOWriter_PageStyleFooterAreaGradientMulticolor(ByRef $oPageStyle[, $avColorStops = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 4 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 5 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 6 Return ? = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve FooterFillGradient Struct.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve StopColor Struct.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current version less than 7.6.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were set to Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, in Long integer format.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOWriter_GradientMulticolorAdd, _LOWriter_GradientMulticolorDelete, _LOWriter_GradientMulticolorModify, _LOWriter_PageStyleFooterAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterAreaGradientMulticolor(ByRef $oPageStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.FooterFillGradient()
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
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)

		$tStopColor.Red = (BitAND(BitShift($avColorStops[$i][1], 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($avColorStops[$i][1], 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($avColorStops[$i][1], 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oPageStyle.FooterFillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oPageStyle.FooterFillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFooterAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterAreaTransparency
; Description ...: Modify or retrieve Transparency settings for a page style Footer.
; Syntax ........: _LOWriter_PageStyleFooterAreaTransparency(ByRef $oPageStyle[, $iTransparency = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency percentage. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTransparency not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Footers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current setting for Transparency in integer format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterAreaTransparency(ByRef $oPageStyle, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oPageStyle.FooterFillTransparence())

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPageStyle.FooterFillTransparenceGradientName = ""
	$oPageStyle.FooterFillTransparence = $iTransparency
	$iError = ($oPageStyle.FooterFillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFooterAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterAreaTransparencyGradient
; Description ...: Modify or retrieve the Page Style Footer transparency gradient settings.
; Syntax ........: _LOWriter_PageStyleFooterAreaTransparencyGradient(ByRef $oDoc, ByRef $oPageStyle[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient that you want to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 4 Return 0 = $iType Not an Integer, less than -1, or greater than 5, see constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iXCenter Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iYCenter Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iAngle Not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 8 Return 0 = $iTransitionStart Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iStart Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 10 Return 0 = $iEnd Not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Footers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving "FillTransparenceGradient" Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 4 Return 0 = Error creating Transparency Gradient Name.
;                  @Error 3 @Extended 5 Return 0 = Error setting Transparency Gradient Name.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterAreaTransparencyGradient(ByRef $oDoc, ByRef $oPageStyle, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop
	Local $fValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.FooterFillTransparenceGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		__LO_ArrayFill($aiTransparent, $tStyleGradient.Style(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), _
				Int($tStyleGradient.Angle() / 10), $tStyleGradient.Border(), __LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.StartColor()), _
				__LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oPageStyle.FooterFillTransparenceGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iStart <> Null) Then
		If Not __LO_IntIsBetween($iStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tStyleGradient.StartColor = __LOWriter_TransparencyGradientConvert($iStart)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$tColorStop = $atColorStop[0] ; StopOffset 0 is the "Start" Value.

			$tStopColor = $tColorStop.StopColor()

			$fValue = $iStart / 100 ; Value is a decimal percentage value.

			$tStopColor.Red = $fValue
			$tStopColor.Green = $fValue
			$tStopColor.Blue = $fValue

			$tColorStop.StopColor = $tStopColor

			$atColorStop[0] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($iEnd <> Null) Then
		If Not __LO_IntIsBetween($iEnd, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tStyleGradient.EndColor = __LOWriter_TransparencyGradientConvert($iEnd)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; StopOffset 0 is the "End" Value.

			$tStopColor = $tColorStop.StopColor()

			$fValue = $iEnd / 100 ; Value is a decimal percentage value.

			$tStopColor.Red = $fValue
			$tStopColor.Green = $fValue
			$tStopColor.Blue = $fValue

			$tColorStop.StopColor = $tStopColor

			$atColorStop[UBound($atColorStop) - 1] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($oPageStyle.FooterFillTransparenceGradientName() = "") Then
		$sTGradName = __LOWriter_TransparencyGradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oPageStyle.FooterFillTransparenceGradientName = $sTGradName
		If ($oPageStyle.FooterFillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	$oPageStyle.FooterFillTransparenceGradient = $tStyleGradient

	$iError = ($iType = Null) ? ($iError) : (($oPageStyle.FooterFillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oPageStyle.FooterFillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oPageStyle.FooterFillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oPageStyle.FooterFillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oPageStyle.FooterFillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iStart = Null) ? ($iError) : (($oPageStyle.FooterFillTransparenceGradient.StartColor() = __LOWriter_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iEnd = Null) ? ($iError) : (($oPageStyle.FooterFillTransparenceGradient.EndColor() = __LOWriter_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFooterAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Page Style's Footer Multi Transparency Gradient settings. See remarks.
; Syntax ........: _LOWriter_PageStyleFooterAreaTransparencyGradientMulti(ByRef $oPageStyle[, $avColorStops = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 4 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 5 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 6 Return ? = ColorStop color not an Integer, less than 0 or greater than 100. Returning problem element index.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve FooterFillTransparenceGradient Struct.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve StopColor Struct.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current version less than 7.6.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were set to Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple Transparency stops in a Gradient rather than just a beginning and an ending value, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the Transparency percentage value between 0 and 100%.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOWriter_TransparencyGradientMultiModify, _LOWriter_TransparencyGradientMultiDelete, _LOWriter_TransparencyGradientMultiAdd, _LOWriter_PageStyleFooterAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterAreaTransparencyGradientMulti(ByRef $oPageStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.FooterFillTransparenceGradient()
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
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)

		$tStopColor.Red = ($avColorStops[$i][1] / 100)
		$tStopColor.Green = ($avColorStops[$i][1] / 100)
		$tStopColor.Blue = ($avColorStops[$i][1] / 100)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oPageStyle.FooterFillTransparenceGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oPageStyle.FooterFillTransparenceGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFooterAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterBorderColor
; Description ...: Set and Retrieve the Page Style Footer Border Line Color.
; Syntax ........: _LOWriter_PageStyleFooterBorderColor(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. Set the Top Border Line Color of the Page Style in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. Set the Bottom Border Line Color of the Page Style in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. Set the Left Border Line Color of the Page Style in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. Set the Right Border Line Color of the Page Style in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or less than 0, or greater than 16,777,215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Footers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Color when Right Border width not set.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_PageStyleFooterBorderWidth, _LOWriter_PageStyleFooterBorderStyle, _LOWriter_PageStyleFooterBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterBorderColor(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_FooterBorder($oPageStyle, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_PageStyleFooterBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterBorderPadding
; Description ...: Set or retrieve the Footer Border Padding settings.
; Syntax ........: _LOWriter_PageStyleFooterBorderPadding(ByRef $oPageStyle[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Micrometers (uM).
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Distance between the Border and Page contents in Micrometers(uM).
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Distance between the Border and Page contents in Micrometers(uM).
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Distance between the Border and Page contents in Micrometers(uM).
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Distance between the Border and Page contents in Micrometers(uM).
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer, _LOWriter_PageStyleFooterBorderWidth, _LOWriter_PageStyleFooterBorderStyle, _LOWriter_PageStyleFooterBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterBorderPadding(ByRef $oPageStyle, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_PageStyleFooterBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterBorderStyle
; Description ...: Set and retrieve the Page Style Footer Border Line style.
; Syntax ........: _LOWriter_PageStyleFooterBorderStyle(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Top Border Line Style of the Page Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Bottom Border Line Style of the Page Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Left Border Line Style of the Page Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Right Border Line Style of the Page Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, set to higher than 17, and not equal to 0x7FFF, Or is set to less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, set to higher than 17, and not equal to 0x7FFF, Or is set to less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, set to higher than 17, and not equal to 0x7FFF, Or is set to less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, set to higher than 17, and not equal to 0x7FFF, Or is set to less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Footers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style when Right Border width not set.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LOWriter_PageStyleFooterBorderWidth, _LOWriter_PageStyleFooterBorderColor, _LOWriter_PageStyleFooterBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterBorderStyle(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_FooterBorder($oPageStyle, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_PageStyleFooterBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterBorderWidth
; Description ...: Set and retrieve the Page Style Footer Border Line Width.
; Syntax ........: _LOWriter_PageStyleFooterBorderWidth(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Border Line width of the Page Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Border Line Width of the Page Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Border Line width of the Page Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Border Line Width of the Page Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Footers are not enabled for this Page Style.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set Width to 0.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer, _LOWriter_PageStyleFooterBorderStyle, _LOWriter_PageStyleFooterBorderColor, _LOWriter_PageStyleFooterBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterBorderWidth(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_FooterBorder($oPageStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_PageStyleFooterBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFooterShadow
; Description ...: Set or Retrieve the shadow settings for a Page Style Footer.
; Syntax ........: _LOWriter_PageStyleFooterShadow(ByRef $oPageStyle[, $iWidth = Null[, $iColor = Null[, $bTransparent = Null[, $iLocation = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Shadow Width of the footer, set in Micrometers.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Color of the Footer shadow, set in Long Integer format, can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bTransparent        - [optional] a boolean value. Default is Null. If True, the Footer Shadow is transparent.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Footer Shadow. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $bTransparent not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iLocation not an Integer, less than 0, or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Footers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving ShadowFormat Object.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving ShadowFormat Object for Error checking.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
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
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFooterShadow(ByRef $oPageStyle, $iWidth = Null, $iColor = Null, $bTransparent = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tShdwFrmt
	Local $iError = 0
	Local $avShadow[4]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.FooterIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tShdwFrmt = $oPageStyle.FooterShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.IsTransparent(), $tShdwFrmt.Location())

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

	If ($bTransparent <> Null) Then
		If Not IsBool($bTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tShdwFrmt.IsTransparent = $bTransparent
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LOW_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oPageStyle.FooterShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oPageStyle.FooterShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$iError = ($iWidth = Null) ? ($iError) : ((__LO_IntIsBetween($tShdwFrmt.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iColor = Null) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($bTransparent = Null) ? ($iError) : (($tShdwFrmt.IsTransparent() = $bTransparent) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iLocation = Null) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFooterShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFootnoteArea
; Description ...: Modify or retrieve Page Style Footnote Size settings.
; Syntax ........: _LOWriter_PageStyleFootnoteArea(ByRef $oPageStyle[, $iFootnoteHeight = Null[, $iSpaceToText = Null]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iFootnoteHeight     - [optional] an integer value. Default is Null. The maximum height for the footnote area. Set in Micrometers. Enter 0 for "Not larger than page", else minimum 508 uM.
;                  $iSpaceToText        - [optional] an integer value. Default is Null. The amount of space to leave between the bottom page margin and the first line of text in the footnote area. Set in Micrometers.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iFootnoteHeight not an Integer, less than 508, but not 0.
;                  @Error 1 @Extended 4 Return 0 = $iSpaceToText not an Integer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iFootnoteHeight
;                  |                               2 = Error setting $iSpaceToText
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFootnoteArea(ByRef $oPageStyle, $iFootnoteHeight = Null, $iSpaceToText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiFootnote[2]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iFootnoteHeight, $iSpaceToText) Then
		__LO_ArrayFill($aiFootnote, $oPageStyle.FootnoteHeight(), $oPageStyle.FootnoteLineTextDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiFootnote)
	EndIf

	If ($iFootnoteHeight <> Null) Then
		If Not __LO_IntIsBetween($iFootnoteHeight, 508, $iFootnoteHeight, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.FootnoteHeight = $iFootnoteHeight
		$iError = (__LO_IntIsBetween($oPageStyle.FootnoteHeight(), $iFootnoteHeight - 1, $iFootnoteHeight + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iSpaceToText <> Null) Then
		If Not IsInt($iSpaceToText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.FootnoteLineTextDistance = $iSpaceToText
		$iError = (__LO_IntIsBetween($oPageStyle.FootnoteLineTextDistance(), $iSpaceToText - 1, $iSpaceToText + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFootnoteArea

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleFootnoteLine
; Description ...: Modify or retrieve the page style footnote separator line settings.
; Syntax ........: _LOWriter_PageStyleFootnoteLine(ByRef $oPageStyle[, $iPosition = Null[, $iStyle = Null[, $nThickness = Null[, $iColor = Null[, $iLength = Null[, $iSpacing = Null]]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iPosition           - [optional] an integer value (0-2). Default is Null. The horizontal alignment for the line that separates the main text from the footnote area. See Constants, $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iStyle              - [optional] an integer value (0-3). Default is Null. The formatting style for the separator line. See Constants, $LOW_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $nThickness          - [optional] a general number value (0-9). Default is Null. The thickness of the separator line. Set in Printer's Points.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The color of the separator line. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLength             - [optional] an integer value (0-100). Default is Null. The length of the separator line as a percentage of the page width area.
;                  $iSpacing            - [optional] an integer value. Default is Null. The amount of space to leave between the separator line and the first line of the footnote area. Set in Micrometers.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iPosition not an Integer, less than 0, or greater than 2. See Constants, $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iStyle not an Integer, less than 0, or greater than 3. See Constants, $LOW_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $nThickness not a Number, less than 0, or greater than 9.
;                  @Error 1 @Extended 6 Return 0 = $iColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 7 Return 0 = $iLength not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iSpacing not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting from Printer's Points to Micrometers.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iStyle
;                  |                               4 = Error setting $nThickness
;                  |                               8 = Error setting $iColor
;                  |                               16 = Error setting $iLength
;                  |                               32 = Error setting $iSpacing
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleFootnoteLine(ByRef $oPageStyle, $iPosition = Null, $iStyle = Null, $nThickness = Null, $iColor = Null, $iLength = Null, $iSpacing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFootnoteLine[6]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iPosition, $iStyle, $nThickness, $iColor, $iLength, $iSpacing) Then
		__LO_ArrayFill($avFootnoteLine, $oPageStyle.FootnoteLineAdjust(), $oPageStyle.FootnoteLineStyle(), _
				__LO_UnitConvert($oPageStyle.FootnoteLineWeight(), $__LOCONST_CONVERT_UM_PT), _ ; Convert Thickness from uM to Point.
				$oPageStyle.FootnoteLineColor(), $oPageStyle.FootnoteLineRelativeWidth(), $oPageStyle.FootnoteLineDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFootnoteLine)
	EndIf

	If ($iPosition <> Null) Then
		If Not __LO_IntIsBetween($iPosition, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPageStyle.FootnoteLineAdjust = $iPosition
		$iError = ($oPageStyle.FootnoteLineAdjust() = $iPosition) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iStyle <> Null) Then
		If Not __LO_IntIsBetween($iStyle, $LOW_LINE_STYLE_NONE, $LOW_LINE_STYLE_DASHED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.FootnoteLineStyle = $iStyle
		$iError = ($oPageStyle.FootnoteLineStyle() = $iStyle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($nThickness <> Null) Then
		If Not __LO_NumIsBetween($nThickness, 0, 9) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$nThickness = __LO_UnitConvert($nThickness, $__LOCONST_CONVERT_PT_UM) ; Convert Thickness from Point to uM
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oPageStyle.FootnoteLineWeight = $nThickness
		$iError = (__LO_IntIsBetween($oPageStyle.FootnoteLineWeight, $nThickness - 1, $nThickness + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.FootnoteLineColor = $iColor
		$iError = ($oPageStyle.FootnoteLineColor() = $iColor) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iLength <> Null) Then
		If Not __LO_IntIsBetween($iLength, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPageStyle.FootnoteLineRelativeWidth = $iLength
		$iError = ($oPageStyle.FootnoteLineRelativeWidth() = $iLength) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iSpacing <> Null) Then
		If Not IsInt($iSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPageStyle.FootnoteLineDistance = $iSpacing
		$iError = (__LO_IntIsBetween($oPageStyle.FootnoteLineDistance, $iSpacing - 1, $iSpacing + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleFootnoteLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleGetObj
; Description ...: Retrieve a Page Style Object for use with other Page Style functions.
; Syntax ........: _LOWriter_PageStyleGetObj(ByRef $oDoc, $sPageStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Related .......: _LOWriter_PageStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleGetObj(ByRef $oDoc, $sPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPageStyle = $oDoc.StyleFamilies().getByName("PageStyles").getByName($sPageStyle)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPageStyle)
EndFunc   ;==>_LOWriter_PageStyleGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeader
; Description ...: Modify or retrieve Header settings for a page style.
; Syntax ........: _LOWriter_PageStyleHeader(ByRef $oPageStyle[, $bHeaderOn = Null[, $bSameLeftRight = Null[, $bSameOnFirst = Null[, $iLeftMargin = Null[, $iRightMargin = Null[, $iSpacing = Null[, $bDynamicSpacing = Null[, $iHeight = Null[, $bAutoHeight = Null]]]]]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $bHeaderOn           - [optional] a boolean value. Default is Null. If True, adds a Header to the page style.
;                  $bSameLeftRight      - [optional] a boolean value. Default is Null. If True, Even and odd pages share the same content.
;                  $bSameOnFirst        - [optional] a boolean value. Default is Null. If True, First and even/odd pages share the same content. LibreOffice 4.0 and up.
;                  $iLeftMargin         - [optional] an integer value. Default is Null. The amount of space to leave between the left edge of the page and the left edge of the Header. Set in Micrometers.
;                  $iRightMargin        - [optional] an integer value. Default is Null. The amount of space to leave between the right edge of the page and the right edge of the Header. Set in Micrometers.
;                  $iSpacing            - [optional] an integer value. Default is Null. The amount of space to maintain between the Top edge of the document text and the bottom edge of the Header. Set in Micrometers.
;                  $bDynamicSpacing     - [optional] a boolean value. Default is Null. If True, Overrides the Spacing setting and allows the Header to expand into the area between the Header and document text.
;                  $iHeight             - [optional] an integer value. Default is Null. The height for the Header. Set in Micrometers.
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
;                  @Error 1 @Extended 9 Return 0 = $bDynamicSpacing not a Boolean value.
;                  @Error 1 @Extended 10 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 11 Return 0 = $bAutoHeight not a Boolean value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHeaderOn
;                  |                               2 = Error setting $bSameLeftRight
;                  |                               4 = Error setting $bSameOnFirst
;                  |                               8 = Error setting $iLeftMargin
;                  |                               16 = Error setting $iRightMargin
;                  |                               32 = Error setting $iSpacing
;                  |                               64 = Error setting $bDynamicSpacing
;                  |                               128 = Error setting $iHeight
;                  |                               256 = Error setting $bAutoHeight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 or 9 Element Array with values in order of function parameters. If Libre Office version is less than 4.0, then the Array returned will contain 8 elements, because $bSameOnFirst will not be available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeader(ByRef $oPageStyle, $bHeaderOn = Null, $bSameLeftRight = Null, $bSameOnFirst = Null, $iLeftMargin = Null, $iRightMargin = Null, $iSpacing = Null, $bDynamicSpacing = Null, $iHeight = Null, $bAutoHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHeader[8]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bHeaderOn, $bSameLeftRight, $bSameOnFirst, $iLeftMargin, $iRightMargin, $iSpacing, $bDynamicSpacing, $iHeight, $bAutoHeight) Then
		If __LO_VersionCheck(4.0) Then
			__LO_ArrayFill($avHeader, $oPageStyle.HeaderIsOn(), $oPageStyle.HeaderIsShared(), $oPageStyle.FirstIsShared(), $oPageStyle.HeaderLeftMargin(), _
					$oPageStyle.HeaderRightMargin(), $oPageStyle.HeaderBodyDistance(), $oPageStyle.HeaderDynamicSpacing(), $oPageStyle.HeaderHeight(), _
					$oPageStyle.HeaderIsDynamicHeight())

		Else
			__LO_ArrayFill($avHeader, $oPageStyle.HeaderIsOn(), $oPageStyle.HeaderIsShared(), $oPageStyle.HeaderLeftMargin(), _
					$oPageStyle.HeaderRightMargin(), $oPageStyle.HeaderBodyDistance(), $oPageStyle.HeaderDynamicSpacing(), $oPageStyle.HeaderHeight(), _
					$oPageStyle.HeaderIsDynamicHeight())
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

		$oPageStyle.FirstIsShared = $bSameOnFirst
		$iError = ($oPageStyle.FirstIsShared() = $bSameOnFirst) ? ($iError) : (BitOR($iError, 4))
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

	If ($bDynamicSpacing <> Null) Then
		If Not IsBool($bDynamicSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oPageStyle.HeaderDynamicSpacing = $bDynamicSpacing
		$iError = ($oPageStyle.HeaderDynamicSpacing() = $bDynamicSpacing) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iHeight <> Null) Then
		If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oPageStyle.HeaderHeight = $iHeight
		$iError = (__LO_IntIsBetween($oPageStyle.HeaderHeight(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bAutoHeight <> Null) Then
		If Not IsBool($bAutoHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oPageStyle.HeaderIsDynamicHeight = $bAutoHeight
		$iError = ($oPageStyle.HeaderIsDynamicHeight() = $bAutoHeight) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleHeader

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderAreaColor
; Description ...: Set or Retrieve background color settings for a Page style header.
; Syntax ........: _LOWriter_PageStyleHeaderAreaColor(ByRef $oPageStyle[, $iBackColor = Null[, $bBackTransparent = Null]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The background color. Set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for "None".
;                  $bBackTransparent    - [optional] a boolean value. Default is Null. If True the background color is transparent.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iBackColor not an integer, less than -1, or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $bBackTransparent not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Headers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  |                               2 = Error setting $bBackTransparent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderAreaColor(ByRef $oPageStyle, $iBackColor = Null, $bBackTransparent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency
	Local $avColor[2]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iBackColor, $bBackTransparent) Then
		__LO_ArrayFill($avColor, __LOWriter_ColorRemoveAlpha($oPageStyle.HeaderBackColor()), $oPageStyle.HeaderBackTransparent())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$iOldTransparency = $oPageStyle.HeaderFillTransparence()
		If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oPageStyle.HeaderBackColor = $iBackColor
		$iError = ($oPageStyle.HeaderBackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))

		$oPageStyle.HeaderFillTransparence = $iOldTransparency
	EndIf

	If ($bBackTransparent <> Null) Then
		If Not IsBool($bBackTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.HeaderBackTransparent = $bBackTransparent
		$iError = ($oPageStyle.HeaderBackTransparent() = $bBackTransparent) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleHeaderAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOWriter_PageStyleHeaderAreaFillStyle(ByRef $oPageStyle)
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Fill Style.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning current background fill style. Return will be one of the constants $LOW_AREA_FILL_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderAreaFillStyle(ByRef $oPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oPageStyle.HeaderFillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOWriter_PageStyleHeaderAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderAreaGradient
; Description ...: Modify or retrieve settings for Page Style Header Background color Gradient.
; Syntax ........: _LOWriter_PageStyleHeaderAreaGradient(ByRef $oDoc, ByRef $oPageStyle[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See Constants, $LOW_GRAD_NAME_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient type that you want to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iIncrement          - [optional] an integer value (0,3-256). Default is Null. Specifies the number of steps of change color. 0 = Automatic.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] an integer value (0-16777215). Default is Null. A color for the beginning point of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] an integer value (0-16777215). Default is Null. A color for the endpoint of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 4 Return 0 = $sGradientName Not a String.
;                  @Error 1 @Extended 5 Return 0 = $iType Not an Integer, less than -1, or greater than 5. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iIncrement Not an Integer, less than 3 but not 0, or greater than 256.
;                  @Error 1 @Extended 7 Return 0 = $iXCenter Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iYCenter Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iAngle Not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 10 Return 0 = $iTransitionStart Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 11 Return 0 = $iFromColor Not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 12 Return 0 = $iToColor Not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 13 Return 0 = $iFromIntense Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 14 Return 0 = $iToIntense Not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Headers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving "FillGradient" Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 4 Return 0 = Error creating Gradient Name.
;                  @Error 3 @Extended 5 Return 0 = Error setting Gradient Name.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 11 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderAreaGradient(ByRef $oDoc, ByRef $oPageStyle, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName
	Local $atColorStop[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.HeaderFillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oPageStyle.HeaderFillGradientName(), $tStyleGradient.Style(), _
				$oPageStyle.HeaderFillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oPageStyle.HeaderFillStyle() <> $LOW_AREA_FILL_STYLE_GRADIENT) Then $oPageStyle.HeaderFillStyle = $LOW_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		__LOWriter_GradientPresets($oDoc, $oPageStyle, $tStyleGradient, $sGradientName, False, True)
		$iError = ($oPageStyle.HeaderFillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oPageStyle.HeaderFillStyle = $LOW_AREA_FILL_STYLE_OFF

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPageStyle.HeaderFillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oPageStyle.HeaderFillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iFromColor <> Null) Then
		If Not __LO_IntIsBetween($iFromColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tStyleGradient.StartColor = $iFromColor

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

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
		If Not __LO_IntIsBetween($iToColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$tStyleGradient.EndColor = $iToColor

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

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
		If Not __LO_IntIsBetween($iFromIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$tStyleGradient.StartIntensity = $iFromIntense
	EndIf

	If ($iToIntense <> Null) Then
		If Not __LO_IntIsBetween($iToIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$tStyleGradient.EndIntensity = $iToIntense
	EndIf

	If ($oPageStyle.HeaderFillGradientName = "") Then
		$sGradName = __LOWriter_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oPageStyle.HeaderFillGradientName = $sGradName
		If ($oPageStyle.HeaderFillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	$oPageStyle.HeaderFillGradient = $tStyleGradient

	; Error checking
	$iError = ($iType = Null) ? ($iError) : (($oPageStyle.HeaderFillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oPageStyle.HeaderFillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oPageStyle.HeaderFillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oPageStyle.HeaderFillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oPageStyle.HeaderFillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = ($iFromColor = Null) ? ($iError) : (($oPageStyle.HeaderFillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = ($iToColor = Null) ? ($iError) : (($oPageStyle.HeaderFillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = ($iFromIntense = Null) ? ($iError) : (($oPageStyle.HeaderFillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = ($iToIntense = Null) ? ($iError) : (($oPageStyle.HeaderFillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleHeaderAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderAreaGradientMulticolor
; Description ...: Set or Retrieve a Page Style's Header Multicolor Gradient settings. See remarks.
; Syntax ........: _LOWriter_PageStyleHeaderAreaGradientMulticolor(ByRef $oPageStyle[, $avColorStops = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 4 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 5 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 6 Return ? = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve HeaderFillGradient Struct.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve StopColor Struct.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current version less than 7.6.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were set to Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, in Long integer format.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOWriter_GradientMulticolorAdd, _LOWriter_GradientMulticolorDelete, _LOWriter_GradientMulticolorModify, _LOWriter_PageStyleHeaderAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderAreaGradientMulticolor(ByRef $oPageStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.HeaderFillGradient()
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
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)

		$tStopColor.Red = (BitAND(BitShift($avColorStops[$i][1], 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($avColorStops[$i][1], 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($avColorStops[$i][1], 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oPageStyle.HeaderFillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oPageStyle.HeaderFillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleHeaderAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderAreaTransparency
; Description ...: Modify or retrieve Transparency settings for a page style Header.
; Syntax ........: _LOWriter_PageStyleHeaderAreaTransparency(ByRef $oPageStyle[, $iTransparency = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTransparency not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Headers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current setting for Transparency in integer format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderAreaTransparency(ByRef $oPageStyle, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oPageStyle.HeaderFillTransparence())

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPageStyle.HeaderFillTransparenceGradientName = ""
	$oPageStyle.HeaderFillTransparence = $iTransparency
	$iError = ($oPageStyle.HeaderFillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleHeaderAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderAreaTransparencyGradient
; Description ...: Modify or retrieve the Page Style Header transparency gradient settings.
; Syntax ........: _LOWriter_PageStyleHeaderAreaTransparencyGradient(ByRef $oDoc, ByRef $oPageStyle[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_GRAD_TYPE_OFF to turn Transparency Gradient off.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iStart              - [optional] an integer value (0-100). Default is Null. The transparency value for the beginning point of the gradient, where 0% is fully opaque and 100% is fully transparent.
;                  $iEnd                - [optional] an integer value (0-100). Default is Null. The transparency value for the endpoint of the gradient, where 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 4 Return 0 = $iType not an Integer, less than -1, or greater than 5, see constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iXCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iYCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iAngle not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 8 Return 0 = $iTransitionStart not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iStart not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 10 Return 0 = $iEnd not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Headers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving "FillTransparenceGradient" Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 4 Return 0 = Error creating Transparency Gradient Name.
;                  @Error 3 @Extended 5 Return 0 = Error setting Transparency Gradient Name.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderAreaTransparencyGradient(ByRef $oDoc, ByRef $oPageStyle, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop
	Local $fValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.HeaderFillTransparenceGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		__LO_ArrayFill($aiTransparent, $tStyleGradient.Style(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), _
				Int($tStyleGradient.Angle() / 10), $tStyleGradient.Border(), __LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.StartColor()), _
				__LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oPageStyle.HeaderFillTransparenceGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iStart <> Null) Then
		If Not __LO_IntIsBetween($iStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tStyleGradient.StartColor = __LOWriter_TransparencyGradientConvert($iStart)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$tColorStop = $atColorStop[0] ; StopOffset 0 is the "Start" Value.

			$tStopColor = $tColorStop.StopColor()

			$fValue = $iStart / 100 ; Value is a decimal percentage value.

			$tStopColor.Red = $fValue
			$tStopColor.Green = $fValue
			$tStopColor.Blue = $fValue

			$tColorStop.StopColor = $tStopColor

			$atColorStop[0] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($iEnd <> Null) Then
		If Not __LO_IntIsBetween($iEnd, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tStyleGradient.EndColor = __LOWriter_TransparencyGradientConvert($iEnd)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; StopOffset 0 is the "End" Value.

			$tStopColor = $tColorStop.StopColor()

			$fValue = $iEnd / 100 ; Value is a decimal percentage value.

			$tStopColor.Red = $fValue
			$tStopColor.Green = $fValue
			$tStopColor.Blue = $fValue

			$tColorStop.StopColor = $tStopColor

			$atColorStop[UBound($atColorStop) - 1] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($oPageStyle.HeaderFillTransparenceGradientName() = "") Then
		$sTGradName = __LOWriter_TransparencyGradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oPageStyle.HeaderFillTransparenceGradientName = $sTGradName
		If ($oPageStyle.HeaderFillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	$oPageStyle.HeaderFillTransparenceGradient = $tStyleGradient

	$iError = ($iType = Null) ? ($iError) : (($oPageStyle.HeaderFillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oPageStyle.HeaderFillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oPageStyle.HeaderFillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oPageStyle.HeaderFillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oPageStyle.HeaderFillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iStart = Null) ? ($iError) : (($oPageStyle.HeaderFillTransparenceGradient.StartColor() = __LOWriter_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iEnd = Null) ? ($iError) : (($oPageStyle.HeaderFillTransparenceGradient.EndColor() = __LOWriter_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleHeaderAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Page Style's Header Multi Transparency Gradient settings. See remarks.
; Syntax ........: _LOWriter_PageStyleHeaderAreaTransparencyGradientMulti(ByRef $oPageStyle[, $avColorStops = Null])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 4 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 5 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 6 Return ? = ColorStop color not an Integer, less than 0 or greater than 100. Returning problem element index.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve HeaderFillTransparenceGradient Struct.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve StopColor Struct.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current version less than 7.6.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were set to Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple Transparency stops in a Gradient rather than just a beginning and an ending value, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the Transparency percentage value between 0 and 100%.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOWriter_TransparencyGradientMultiModify, _LOWriter_TransparencyGradientMultiDelete, _LOWriter_TransparencyGradientMultiAdd, _LOWriter_PageStyleHeaderAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderAreaTransparencyGradientMulti(ByRef $oPageStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oPageStyle.HeaderFillTransparenceGradient()
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
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)

		$tStopColor.Red = ($avColorStops[$i][1] / 100)
		$tStopColor.Green = ($avColorStops[$i][1] / 100)
		$tStopColor.Blue = ($avColorStops[$i][1] / 100)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oPageStyle.HeaderFillTransparenceGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oPageStyle.HeaderFillTransparenceGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleHeaderAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderBorderColor
; Description ...: Set and Retrieve the Page Style Header Border Line Color.
; Syntax ........: _LOWriter_PageStyleHeaderBorderColor(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. Set the Top Border Line Color of the Page Style in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. Set the Bottom Border Line Color of the Page Style in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. Set the Left Border Line Color of the Page Style in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. Set the Right Border Line Color of the Page Style in Long Color code format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0, or greater than 16,777,215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Headers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Color when Right Border width not set.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_PageStyleHeaderBorderWidth, _LOWriter_PageStyleHeaderBorderStyle, _LOWriter_PageStyleHeaderBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderBorderColor(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_HeaderBorder($oPageStyle, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_PageStyleHeaderBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderBorderPadding
; Description ...: Set or retrieve the Header Border Padding settings.
; Syntax ........: _LOWriter_PageStyleHeaderBorderPadding(ByRef $oPageStyle[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Micrometers (uM).
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Distance between the Border and Page Header contents in Micrometers(uM).
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Distance between the Border and Page Header contents in Micrometers(uM).
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Distance between the Border and Page Header contents in Micrometers(uM).
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Distance between the Border and Page Header contents in Micrometers(uM).
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer, _LOWriter_PageStyleHeaderBorderWidth, _LOWriter_PageStyleHeaderBorderStyle, _LOWriter_PageStyleHeaderBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderBorderPadding(ByRef $oPageStyle, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_PageStyleHeaderBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderBorderStyle
; Description ...: Set and retrieve the Page Style Header Border Line style.
; Syntax ........: _LOWriter_PageStyleHeaderBorderStyle(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. Sets the Top Border Line Style of the Page Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. Sets the Bottom Border Line Style of the Page Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. Sets the Left Border Line Style of the Page Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. Sets the Right Border Line Style of the Page Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or greater than 17, and not equal to 0x7FFF, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or greater than 17, and not equal to 0x7FFF, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or greater than 17, and not equal to 0x7FFF, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or greater than 17, and not equal to 0x7FFF, or less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Headers are not enabled for this Page Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style Top when Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style Bottom when Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style when Right Border width not set.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LOWriter_PageStyleHeaderBorderWidth, _LOWriter_PageStyleHeaderBorderColor, _LOWriter_PageStyleHeaderBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderBorderStyle(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_HeaderBorder($oPageStyle, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_PageStyleHeaderBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderBorderWidth
; Description ...: Set and retrieve the Page Style Header Border Line Width.
; Syntax ........: _LOWriter_PageStyleHeaderBorderWidth(ByRef $oPageStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Border Line width of the Page Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Border Line Width of the Page Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Border Line width of the Page Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Border Line Width of the Page Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Headers are not enabled for this Page Style.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set Width to 0.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer, _LOWriter_PageStyleHeaderBorderStyle, _LOWriter_PageStyleHeaderBorderColor, _LOWriter_PageStyleHeaderBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderBorderWidth(ByRef $oPageStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_HeaderBorder($oPageStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_PageStyleHeaderBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleHeaderShadow
; Description ...: Set or Retrieve the shadow settings for a Page Style Header.
; Syntax ........: _LOWriter_PageStyleHeaderShadow(ByRef $oPageStyle[, $iWidth = Null[, $iColor = Null[, $bTransparent = Null[, $iLocation = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Shadow Width of the Header, set in Micrometers.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Color of the Header shadow, set in Long Integer format, can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bTransparent        - [optional] a boolean value. Default is Null. If True, the Header Shadow is transparent.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Header Shadow. See constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $bTransparent not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iLocation not an Integer, less than 0, or greater than 4. See Constants.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Headers are not enabled for this Page Style.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving ShadowFormat Object.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving ShadowFormat Object for Error Checking.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
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
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleHeaderShadow(ByRef $oPageStyle, $iWidth = Null, $iColor = Null, $bTransparent = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tShdwFrmt
	Local $iError = 0
	Local $avShadow[4]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oPageStyle.HeaderIsOn() = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tShdwFrmt = $oPageStyle.HeaderShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.IsTransparent(), $tShdwFrmt.Location())

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

	If ($bTransparent <> Null) Then
		If Not IsBool($bTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tShdwFrmt.IsTransparent = $bTransparent
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oPageStyle.HeaderShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oPageStyle.HeaderShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$iError = ($iWidth = Null) ? ($iError) : ((__LO_IntIsBetween($tShdwFrmt.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iColor = Null) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($bTransparent = Null) ? ($iError) : (($tShdwFrmt.IsTransparent() = $bTransparent) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iLocation = Null) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleHeaderShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleLayout
; Description ...: Modify or retrieve the Layout settings for a Page Style.
; Syntax ........: _LOWriter_PageStyleLayout(ByRef $oDoc, $oPageStyle[, $iLayout = Null[, $iNumFormat = Null[, $sRefStyle = Null[, $bGutterOnRight = Null[, $bGutterAtTop = Null[, $bBackCoversMargins = Null[, $sPaperTray = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iLayout             - [optional] an integer value (0-4). Default is Null. Specify the current Page layout style, either Left(Even) pages, Right(Odd) pages, or both Left(Even) and Right(Odd) pages or mirrored. See Constants, $LOW_PAGE_LAYOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The page numbering format to use for this Page Style. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sRefStyle           - [optional] a string value. Default is Null. The Paragraph Style to use as a reference for lining up the text on the selected Page style. To disable Page Spacing alignment, set to "".
;                  $bGutterOnRight      - [optional] a boolean value. Default is Null. If True, the page gutter will be placed on the right side of the page. Libre 7.2 and up.
;                  $bGutterAtTop        - [optional] a boolean value. Default is Null. If False, the current document's gutter will be positioned at the left of the document's pages (L.O. default) or If True, at top of the document's pages when the document is displayed.
;                  $bBackCoversMargins  - [optional] a boolean value. Default is Null. If true, the background covers the full page, Else only inside the margins. Libre 7.2 and up.
;                  $sPaperTray          - [optional] a string value. Default is Null. The paper source for your printer. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 4 Return 0 = $iLayout not an Integer, less than 0, or greater than 4. See Constants, $LOW_PAGE_LAYOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iNumFormat not an Integer, less than 0, or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $sRefStyle not a String.
;                  @Error 1 @Extended 7 Return 0 = Paragraph style referenced in $sRefStyle not found in document and $sRefStyle not equal to "".
;                  @Error 1 @Extended 8 Return 0 = $bGutterOnRight not a Boolean value.
;                  @Error 1 @Extended 9 Return 0 = $bGutterAtTop not a Boolean value.
;                  @Error 1 @Extended 10 Return 0 = $bBackCoversMargins not a Boolean value.
;                  @Error 1 @Extended 11 Return 0 = $sPaperTray not a string.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating Document Settings Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLayout
;                  |                               2 = Error setting $iNumFormat
;                  |                               4 = Error setting $sRefStyle
;                  |                               8 = Error setting $bGutterOnRight
;                  |                               16 = Error setting $bGutterAtTop
;                  |                               32 = Error setting $bBackCoversMargins
;                  |                               64 = Error setting $sPaperTray
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 7.2.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 or 7 Element Array with values in order of function parameters. If the current Libre Office version is less than 7.2, the Array will be a 4 element Array, because $bGutterOnRight, $bGutterAtTop, and $bBackCoversMargins will not be available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I have no way to retrieve possible values for the Paper Tray parameter, at least that I can find. You may still use it if you know the appropriate value.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleLayout(ByRef $oDoc, ByRef $oPageStyle, $iLayout = Null, $iNumFormat = Null, $sRefStyle = Null, $bGutterOnRight = Null, $bGutterAtTop = Null, $bBackCoversMargins = Null, $sPaperTray = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSettings
	Local $iError = 0
	Local $avLayout[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSettings = $oDoc.createInstance("com.sun.star.text.DocumentSettings")
	If Not IsObj($oSettings) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLayout, $iNumFormat, $sRefStyle, $bGutterOnRight, $bGutterAtTop, $bBackCoversMargins, $sPaperTray) Then
		If __LO_VersionCheck(7.2) Then
			__LO_ArrayFill($avLayout, $oPageStyle.PageStyleLayout(), $oPageStyle.NumberingType(), _
					__LOWriter_ParStyleNameToggle($oPageStyle.RegisterParagraphStyle(), True), _
					$oPageStyle.RtlGutter(), $oSettings.getPropertyValue("GutterAtTop"), $oPageStyle.BackgroundFullSize(), _
					$oPageStyle.PrinterPaperTray())

		Else
			__LO_ArrayFill($avLayout, $oPageStyle.PageStyleLayout(), $oPageStyle.NumberingType(), _
					__LOWriter_ParStyleNameToggle($oPageStyle.RegisterParagraphStyle(), True), $oPageStyle.PrinterPaperTray())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avLayout)
	EndIf

	If ($iLayout <> Null) Then
		If Not __LO_IntIsBetween($iLayout, $LOW_PAGE_LAYOUT_ALL, $LOW_PAGE_LAYOUT_MIRRORED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPageStyle.PageStyleLayout = $iLayout
		$iError = ($oPageStyle.PageStyleLayout() = $iLayout) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LO_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.NumberingType = $iNumFormat
		$iError = ($oPageStyle.NumberingType() = $iNumFormat) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sRefStyle <> Null) Then
		If Not IsString($sRefStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_ParStyleExists($oDoc, $sRefStyle) And Not ($sRefStyle = "") Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$sRefStyle = __LOWriter_ParStyleNameToggle($sRefStyle)
		$oPageStyle.RegisterParagraphStyle = $sRefStyle
		$iError = ($oPageStyle.RegisterParagraphStyle() = $sRefStyle) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bGutterOnRight <> Null) Then
		If Not IsBool($bGutterOnRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		If Not __LO_VersionCheck(7.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oPageStyle.RtlGutter = $bGutterOnRight
		$iError = ($oPageStyle.RtlGutter() = $bGutterOnRight) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bGutterAtTop <> Null) Then
		If Not IsBool($bGutterAtTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not __LO_VersionCheck(7.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oSettings.setPropertyValue("GutterAtTop", $bGutterAtTop)
		$iError = ($oSettings.getPropertyValue("GutterAtTop") = $bGutterAtTop) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bBackCoversMargins <> Null) Then
		If Not IsBool($bBackCoversMargins) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		If Not __LO_VersionCheck(7.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oPageStyle.BackgroundFullSize = $bBackCoversMargins
		$iError = ($oPageStyle.BackgroundFullSize() = $bBackCoversMargins) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($sPaperTray <> Null) Then
		If Not IsString($sPaperTray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oPageStyle.PrinterPaperTray = $sPaperTray
		$iError = ($oPageStyle.PrinterPaperTray() = $sPaperTray) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleLayout

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleMargins
; Description ...: Modify or retrieve the margin settings for a Page Style.
; Syntax ........: _LOWriter_PageStyleMargins(ByRef $oPageStyle[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null[, $iGutter = Null]]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iLeft               - [optional] an integer value. Default is Null. The amount of space to leave between the left edge of the page and the document text. If you are using the Mirrored page layout, enter the amount of space to leave between the inner text margin and the inner edge of the page. Set in Micrometers.
;                  $iRight              - [optional] an integer value. Default is Null. The amount of space to leave between the right edge of the page and the document text. If you are using the Mirrored page layout, enter the amount of space to leave between the outer text margin and the outer edge of the page. Set in Micrometers.
;                  $iTop                - [optional] an integer value. Default is Null. The amount of space to leave between the upper edge of the page and the document text. Set in Micrometers.
;                  $iBottom             - [optional] an integer value. Default is Null. The amount of space to leave between the lower edge of the page and the document text. Set in Micrometers.
;                  $iGutter             - [optional] an integer value. Default is Null. The amount of space to leave between the left edge of the page and the left margin. If you are using the Mirrored page layout, enter the amount of space to leave between the inner page margin and the inner edge of the page. Set in Micrometers. Libre 7.2 and up.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iLeft not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iRight not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iGutter not an Integer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
;                  |                               16 = Error setting $iGutter
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 7.2.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 or 5 Element Array with values in order of function parameters. If the current Libre Office version is less than 7.2, then the array will have 4 elements as Gutter Margin will not be available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleMargins(ByRef $oPageStyle, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null, $iGutter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiMargins[4]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iLeft, $iRight, $iTop, $iBottom, $iGutter) Then
		If __LO_VersionCheck(7.2) Then
			__LO_ArrayFill($aiMargins, $oPageStyle.LeftMargin(), $oPageStyle.RightMargin(), $oPageStyle.TopMargin(), $oPageStyle.BottomMargin(), _
					$oPageStyle.GutterMargin())

		Else
			__LO_ArrayFill($aiMargins, $oPageStyle.LeftMargin(), $oPageStyle.RightMargin(), $oPageStyle.TopMargin(), $oPageStyle.BottomMargin())
		EndIf

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

	If ($iGutter <> Null) Then
		If Not IsInt($iGutter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If Not __LO_VersionCheck(7.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oPageStyle.GutterMargin = $iGutter
		$iError = (__LO_IntIsBetween($oPageStyle.GutterMargin(), $iGutter - 1, $iGutter + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleOrganizer
; Description ...: Set or retrieve the Organizer settings of a Page Style.
; Syntax ........: _LOWriter_PageStyleOrganizer(ByRef $oDoc, $oPageStyle[, $sNewPageStyleName = Null[, $bHidden = Null[, $sFollowStyle = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $sNewPageStyleName   - [optional] a string value. Default is Null. The new name to set the Page Style called in $oPageStyle to.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, the style is hidden in L.O. UI. Libre Office 4.0 and Up.
;                  $sFollowStyle        - [optional] a string value. Default is Null. The name of the Page style that is applied After this Page Style.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 4 Return 0 = $sNewPageStyleName not a String.
;                  @Error 1 @Extended 5 Return 0 = Page Style name called in $sNewPageStyleName already exists in document.
;                  @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $sFollowStyle not a String.
;                  @Error 1 @Extended 8 Return 0 = Page Style called in $sFollowStyle doesn't exist in this document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sNewParStyleName
;                  |                               2 = Error setting $bHidden
;                  |                               4 = Error setting $sFollowStyle
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 or 3 Element Array with values in order of function parameters. If the Libre Office version is below 4.0, the Array will contain 2 elements because $bHidden is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LOWriter_PageStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleOrganizer(ByRef $oDoc, ByRef $oPageStyle, $sNewPageStyleName = Null, $bHidden = Null, $sFollowStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOrganizer[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($sNewPageStyleName, $bHidden, $sFollowStyle) Then
		If __LO_VersionCheck(4.0) Then
			__LO_ArrayFill($avOrganizer, $oPageStyle.Name(), $oPageStyle.Hidden(), __LOWriter_PageStyleNameToggle($oPageStyle.FollowStyle(), True))

		Else
			__LO_ArrayFill($avOrganizer, $oPageStyle.Name(), __LOWriter_PageStyleNameToggle($oPageStyle.FollowStyle(), True))
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOrganizer)
	EndIf

	If ($sNewPageStyleName <> Null) Then
		If Not IsString($sNewPageStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If _LOWriter_PageStyleExists($oDoc, $sNewPageStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPageStyle.Name = $sNewPageStyleName
		$iError = ($oPageStyle.Name() = $sNewPageStyleName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not __LO_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oPageStyle.Hidden = $bHidden
		$iError = ($oPageStyle.Hidden() = $bHidden) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sFollowStyle <> Null) Then
		If Not IsString($sFollowStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If Not _LOWriter_PageStyleExists($oDoc, $sFollowStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$sFollowStyle = __LOWriter_PageStyleNameToggle($sFollowStyle)
		$oPageStyle.FollowStyle = $sFollowStyle
		$iError = ($oPageStyle.FollowStyle() = $sFollowStyle) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleOrganizer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStylePaperFormat
; Description ...: Modify or retrieve the paper format settings for a Page Style.
; Syntax ........: _LOWriter_PageStylePaperFormat(ByRef $oPageStyle[, $iWidth = Null[, $iHeight = Null[, $bLandscape = Null]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Width of the page, may be a custom value in Micrometers, or one of the constants, $LOW_PAPER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iHeight             - [optional] an integer value. Default is Null. The Height of the page, may be a custom value in Micrometers, or one of the constants, $LOW_PAPER_HEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bLandscape          - [optional] a boolean value. Default is Null. If true, displays the page in Landscape layout.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStylePaperFormat(ByRef $oPageStyle, $iWidth = Null, $iHeight = Null, $bLandscape = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_PageStylePaperFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleSet
; Description ...: Set a Page style for a paragraph by Cursor or paragraph Object.
; Syntax ........: _LOWriter_PageStyleSet(ByRef $oDoc, ByRef $oObj, $sPageStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj                - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object returned from _LOWriter_ParObjCreateList function.
;                  $sPageStyle          - a string value. The Page Style name to set the Page to.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oObj does not support Paragraph Properties Service.
;                  @Error 1 @Extended 4 Return 0 = $sPageStyle not a String.
;                  @Error 1 @Extended 5 Return 0 = Page Style called in $sPageStyle doesn't exist in Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Error setting Page Style.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Page Style successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ParObjCreateList, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_PageStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleSet(ByRef $oDoc, ByRef $oObj, $sPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oObj.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not _LOWriter_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$sPageStyle = __LOWriter_PageStyleNameToggle($sPageStyle)
	$oObj.PageDescName = $sPageStyle

	Return ($oObj.PageStyleName() = $sPageStyle) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOWriter_PageStyleSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStylesGetNames
; Description ...: Retrieve an array of all Page Style names available for a document.
; Syntax ........: _LOWriter_PageStylesGetNames(ByRef $oDoc[, $bUserOnly = False[, $bAppliedOnly = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
;                  @Error 0 @Extended ? Return Array = Success. An Array containing all Page Styles matching the input parameters. See remarks. @Extended contains the count of results returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If Only a Document object is input, all available Page styles will be returned.
;                  Else if $bUserOnly is set to True, only User-Created Page Styles are returned.
;                  Else if $bAppliedOnly is set to True, only Applied Page Styles are returned.
;                  If Both are true then only User-Created Page styles that are applied are returned.
;                  One Page style has two separate names, Default Page Style is also internally called "Standard".
;                  Either name works when setting a Page Style, but on certain functions that return a Page Style Name, you may see the alternative name.
; Related .......: _LOWriter_PageStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStylesGetNames(ByRef $oDoc, $bUserOnly = False, $bAppliedOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local $sExecute = ""
	Local $aStyles[0]
	Local $oStyles

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oStyles = $oDoc.StyleFamilies.getByName("PageStyles")
	If Not IsObj($oStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $aStyles[$oStyles.getCount()]

	If Not $bUserOnly And Not $bAppliedOnly Then
		For $i = 0 To $oStyles.getCount() - 1
			$aStyles[$i] = $oStyles.getByIndex($i).DisplayName
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
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
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next
	ReDim $aStyles[$iCount]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aStyles)
EndFunc   ;==>_LOWriter_PageStylesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PageStyleShadow
; Description ...: Set or Retrieve the shadow settings for a Page Style.
; Syntax ........: _LOWriter_PageStyleShadow(ByRef $oPageStyle[, $iWidth = Null[, $iColor = Null[, $bTransparent = Null[, $iLocation = Null]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Shadow Width of the Page, set in Micrometers.
;                  $iColor              - [optional] an integer value. Default is Null. The shadow Color of the Page, set in Long Integer format, can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bTransparent        - [optional] a boolean value. Default is Null. If True, the Page Shadow is transparent.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Page Shadow. See constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not a Page Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $bTransparent not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iLocation not an Integer, less than 0, or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ShadowFormat Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving ShadowFormat Object for Error checking.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
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
; Related .......: _LOWriter_PageStyleCreate, _LOWriter_PageStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PageStyleShadow(ByRef $oPageStyle, $iWidth = Null, $iColor = Null, $bTransparent = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tShdwFrmt
	Local $iError = 0
	Local $avShadow[4]

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oPageStyle.supportsService("com.sun.star.style.PageStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tShdwFrmt = $oPageStyle.ShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.IsTransparent(), $tShdwFrmt.Location())

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

	If ($bTransparent <> Null) Then
		If Not IsBool($bTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tShdwFrmt.IsTransparent = $bTransparent
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LOW_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oPageStyle.ShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oPageStyle.ShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iError = ($iWidth = Null) ? ($iError) : ((__LO_IntIsBetween($tShdwFrmt.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iColor = Null) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($bTransparent = Null) ? ($iError) : (($tShdwFrmt.IsTransparent() = $bTransparent) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iLocation = Null) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_PageStyleShadow
