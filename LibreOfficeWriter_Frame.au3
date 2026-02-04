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
#include "LibreOfficeWriter_Doc.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Applying L.O. Writer Frame Styles. Also for Creating, and Modifying Frames.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_FrameAreaColor
; _LOWriter_FrameAreaFillStyle
; _LOWriter_FrameAreaGradient
; _LOWriter_FrameAreaGradientMulticolor
; _LOWriter_FrameAreaTransparency
; _LOWriter_FrameAreaTransparencyGradient
; _LOWriter_FrameAreaTransparencyGradientMulti
; _LOWriter_FrameBorderColor
; _LOWriter_FrameBorderPadding
; _LOWriter_FrameBorderStyle
; _LOWriter_FrameBorderWidth
; _LOWriter_FrameColumnSeparator
; _LOWriter_FrameColumnSettings
; _LOWriter_FrameColumnSize
; _LOWriter_FrameCreate
; _LOWriter_FrameCreateTextCursor
; _LOWriter_FrameDelete
; _LOWriter_FrameExists
; _LOWriter_FrameGetAnchor
; _LOWriter_FrameGetObjByCursor
; _LOWriter_FrameGetObjByName
; _LOWriter_FrameHyperlink
; _LOWriter_FrameOptions
; _LOWriter_FrameOptionsName
; _LOWriter_FramesGetNames
; _LOWriter_FrameShadow
; _LOWriter_FrameStyleAreaColor
; _LOWriter_FrameStyleAreaFillStyle
; _LOWriter_FrameStyleAreaGradient
; _LOWriter_FrameStyleAreaGradientMulticolor
; _LOWriter_FrameStyleAreaTransparency
; _LOWriter_FrameStyleAreaTransparencyGradient
; _LOWriter_FrameStyleAreaTransparencyGradientMulti
; _LOWriter_FrameStyleBorderColor
; _LOWriter_FrameStyleBorderPadding
; _LOWriter_FrameStyleBorderStyle
; _LOWriter_FrameStyleBorderWidth
; _LOWriter_FrameStyleColumnSeparator
; _LOWriter_FrameStyleColumnSettings
; _LOWriter_FrameStyleColumnSize
; _LOWriter_FrameStyleCreate
; _LOWriter_FrameStyleCurrent
; _LOWriter_FrameStyleDelete
; _LOWriter_FrameStyleExists
; _LOWriter_FrameStyleGetObj
; _LOWriter_FrameStyleOptions
; _LOWriter_FrameStyleOrganizer
; _LOWriter_FrameStylesGetNames
; _LOWriter_FrameStyleShadow
; _LOWriter_FrameStyleTypePosition
; _LOWriter_FrameStyleTypeSize
; _LOWriter_FrameStyleWrap
; _LOWriter_FrameStyleWrapOptions
; _LOWriter_FrameTypePosition
; _LOWriter_FrameTypeSize
; _LOWriter_FrameWrap
; _LOWriter_FrameWrapOptions
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameAreaColor
; Description ...: Set or Retrieve background color settings for a Frame.
; Syntax ........: _LOWriter_FrameAreaColor(ByRef $oFrame[, $iBackColor = Null])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Background color.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameAreaColor(ByRef $oFrame, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency
	Local $iColor

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iBackColor) Then
		$iColor = __LOWriter_ColorRemoveAlpha($oFrame.BackColor())
		If Not IsInt($iColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iOldTransparency = $oFrame.FillTransparence()
	If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oFrame.BackColor = $iBackColor
	$iError = ($oFrame.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))

	$oFrame.FillTransparence = $iOldTransparency

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOWriter_FrameAreaFillStyle(ByRef $oFrame)
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
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
Func _LOWriter_FrameAreaFillStyle(ByRef $oFrame)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oFrame.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOWriter_FrameAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameAreaGradient
; Description ...: Modify or retrieve the settings for Frame Background color Gradient.
; Syntax ........: _LOWriter_FrameAreaGradient(ByRef $oDoc, ByRef $oFrame[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See remarks. See constants, $LOW_GRAD_NAME_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient that you want to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iIncrement          - [optional] an integer value (0,3-256). Default is Null. Specifies the number of steps of change color. Allowed values are: 0, 3 to 256. 0 = Automatic.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] an integer value (0-16777215). Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] an integer value (0-16777215). Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sGradientName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iIncrement not an Integer, less than 3 but not 0, or greater than 256.
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
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameAreaGradient(ByRef $oDoc, ByRef $oFrame, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName
	Local $atColorStop[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tStyleGradient = $oFrame.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oFrame.FillGradientName(), $tStyleGradient.Style(), _
				$oFrame.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oFrame.FillStyle() <> $LOW_AREA_FILL_STYLE_GRADIENT) Then $oFrame.FillStyle = $LOW_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		__LOWriter_GradientPresets($oDoc, $oFrame, $tStyleGradient, $sGradientName)
		$iError = ($oFrame.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oFrame.FillStyle = $LOW_AREA_FILL_STYLE_OFF
			$oFrame.FillGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrame.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oFrame.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
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

	If ($oFrame.FillGradientName() = "") Or __LOWriter_GradientIsModified($tStyleGradient, $oFrame.FillGradientName()) Then
		$sGradName = __LOWriter_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oFrame.FillGradientName = $sGradName
		If ($oFrame.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oFrame.FillGradient = $tStyleGradient

	; Error checking
	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oFrame.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oFrame.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oFrame.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oFrame.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oFrame.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = (__LO_VarsAreNull($iFromColor)) ? ($iError) : (($oFrame.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = (__LO_VarsAreNull($iToColor)) ? ($iError) : (($oFrame.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = (__LO_VarsAreNull($iFromIntense)) ? ($iError) : (($oFrame.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = (__LO_VarsAreNull($iToIntense)) ? ($iError) : (($oFrame.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameAreaGradientMulticolor
; Description ...: Set or Retrieve a Frame's Multicolor Gradient settings. See remarks.
; Syntax ........: _LOWriter_FrameAreaGradientMulticolor(ByRef $oFrame[, $avColorStops = Null])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
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
; Related .......: _LOWriter_GradientMulticolorAdd, _LOWriter_GradientMulticolorDelete, _LOWriter_GradientMulticolorModify, _LOWriter_FrameAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameAreaGradientMulticolor(ByRef $oFrame, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oFrame.FillGradient()
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

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oFrame.FillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oFrame.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameAreaTransparency
; Description ...: Set or retrieve Transparency settings for a Frame.
; Syntax ........: _LOWriter_FrameAreaTransparency(ByRef $oFrame[, $iTransparency = Null])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
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
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByCursor, _LOWriter_FrameGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameAreaTransparency(ByRef $oFrame, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oFrame.FillTransparence())

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oFrame.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oFrame.FillTransparence = $iTransparency
	$iError = ($oFrame.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameAreaTransparencyGradient
; Description ...: Set or retrieve the Frame transparency gradient settings.
; Syntax ........: _LOWriter_FrameAreaTransparencyGradient(ByRef $oDoc, ByRef $oFrame[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Call with $LOW_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error 1 @Extended 2 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iType Not an Integer, less than -1 or greater than 5. See constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iXCenter Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 5 Return 0 = $iYCenter Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iAngle Not an Integer, less than 0 or greater than 359.
;                  @Error 1 @Extended 7 Return 0 = $iTransitionStart Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iStart Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iEnd Not an Integer, less than 0 or greater than 100.
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
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByCursor, _LOWriter_FrameGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameAreaTransparencyGradient(ByRef $oDoc, ByRef $oFrame, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop
	Local $fValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tGradient = $oFrame.FillTransparenceGradient()
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		__LO_ArrayFill($aiTransparent, $tGradient.Style(), $tGradient.XOffset(), $tGradient.YOffset(), _
				Int($tGradient.Angle() / 10), $tGradient.Border(), __LOWriter_TransparencyGradientConvert(Null, $tGradient.StartColor()), _
				__LOWriter_TransparencyGradientConvert(Null, $tGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oFrame.FillTransparenceGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

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

		$tGradient.StartColor = __LOWriter_TransparencyGradientConvert($iStart)

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

		$tGradient.EndColor = __LOWriter_TransparencyGradientConvert($iEnd)

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

	If ($oFrame.FillTransparenceGradientName() = "") Then
		$sTGradName = __LOWriter_TransparencyGradientNameInsert($oDoc, $tGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oFrame.FillTransparenceGradientName = $sTGradName
		If ($oFrame.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	EndIf

	$oFrame.FillTransparenceGradient = $tGradient

	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oFrame.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oFrame.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oFrame.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oFrame.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oFrame.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iStart)) ? ($iError) : (($oFrame.FillTransparenceGradient.StartColor() = __LOWriter_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iEnd)) ? ($iError) : (($oFrame.FillTransparenceGradient.EndColor() = __LOWriter_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Frame's Multi Transparency Gradient settings. See remarks.
; Syntax ........: _LOWriter_FrameAreaTransparencyGradientMulti(ByRef $oFrame[, $avColorStops = Null])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
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
; Related .......: _LOWriter_TransparencyGradientMultiModify, _LOWriter_TransparencyGradientMultiDelete, _LOWriter_TransparencyGradientMultiAdd, _LOWriter_FrameAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameAreaTransparencyGradientMulti(ByRef $oFrame, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oFrame.FillTransparenceGradient()
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

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oFrame.FillTransparenceGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oFrame.FillTransparenceGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameBorderColor
; Description ...: Set or retrieve the Frame Border Line Color. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_FrameBorderColor(ByRef $oFrame[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. The Top Border Line Color of the Frame, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. The Bottom Border Line Color of the Frame, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. The Left Border Line Color of the Frame, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. The Right Border Line Color of the Frame, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTop not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 3 Return 0 = $iBottom not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $iLeft not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iRight not an Integer, less than 0 or greater than 16777215.
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
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_FrameBorderWidth, _LOWriter_FrameBorderStyle, _LOWriter_FrameBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameBorderColor(ByRef $oFrame, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oFrame, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FrameBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameBorderPadding
; Description ...: Set or retrieve the Frame Border Padding settings.
; Syntax ........: _LOWriter_FrameBorderPadding(ByRef $oFrame[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The Top Distance between the Border and Frame contents in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Distance between the Border and Frame contents in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Distance between the Border and Frame contents in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The Right Distance between the Border and Frame contents in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iAll not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $Left not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer.
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
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LO_UnitConvert, _LOWriter_FrameBorderWidth, _LOWriter_FrameBorderStyle, _LOWriter_FrameBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameBorderPadding(ByRef $oFrame, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oFrame.BorderDistance(), $oFrame.TopBorderDistance(), $oFrame.BottomBorderDistance(), _
				$oFrame.LeftBorderDistance(), $oFrame.RightBorderDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LO_IntIsBetween($iAll, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oFrame.BorderDistance = $iAll
		$iError = (__LO_IntIsBetween($oFrame.BorderDistance(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrame.TopBorderDistance = $iTop
		$iError = (__LO_IntIsBetween($oFrame.TopBorderDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrame.BottomBorderDistance = $iBottom
		$iError = (__LO_IntIsBetween($oFrame.BottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrame.LeftBorderDistance = $iLeft
		$iError = (__LO_IntIsBetween($oFrame.LeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrame.RightBorderDistance = $iRight
		$iError = (__LO_IntIsBetween($oFrame.RightBorderDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameBorderStyle
; Description ...: Set or Retrieve the Frame Border Line style. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_FrameBorderStyle(ByRef $oFrame[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. The Top Border Line Style of the Frame. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. The Bottom Border Line Style of the Frame. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. The Left Border Line Style of the Frame. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. The Right Border Line Style of the Frame. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 3 Return 0 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 4 Return 0 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 5 Return 0 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
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
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LOWriter_FrameBorderWidth, _LOWriter_FrameBorderColor, _LOWriter_FrameBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameBorderStyle(ByRef $oFrame, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oFrame, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FrameBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameBorderWidth
; Description ...: Set or Retrieve the Frame Border Line Width. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_FrameBorderWidth(ByRef $oFrame[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iTop                - [optional] an integer value. Default is Null. The Top Border Line width of the Frame in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Border Line Width of the Frame in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Border Line width of the Frame in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. The Right Border Line Width of the Frame in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTop not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = $iBottom not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iLeft not an Integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iRight not an Integer, or less than 0.
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
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set Width to 0
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LO_UnitConvert, _LOWriter_FrameBorderStyle, _LOWriter_FrameBorderColor, _LOWriter_FrameBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameBorderWidth(ByRef $oFrame, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oFrame, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FrameBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameColumnSeparator
; Description ...: Set or retrieve Frame Column Separator line settings.
; Syntax ........: _LOWriter_FrameColumnSeparator(ByRef $oFrame[, $bSeparatorOn = Null[, $iStyle = Null[, $iWidth = Null[, $iColor = Null[, $iHeight = Null[, $iPosition = Null]]]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $bSeparatorOn        - [optional] a boolean value. Default is Null. If True, add a separator line between two or more columns.
;                  $iStyle              - [optional] an integer value (0-3). Default is Null. The formatting style for the column separator line. See Constants, $LOW_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iWidth              - [optional] an integer value (5-180). Default is Null. The width of the separator line. Set in Hundredths of a Millimeter (HMM).
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The separator line color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iHeight             - [optional] an integer value (0-100). Default is Null. The length of the separator line as a percentage of the height of the column area.
;                  $iPosition           - [optional] an integer value (0-2). Default is Null. Select the vertical alignment of the separator line. This option is only available if Height value of the line is less than 100%. See Constants, $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bSeparatorOn not a Boolean value.
;                  @Error 1 @Extended 3 Return 0 = $iStyle not an Integer, less than 0 or greater than 3. See constants.
;                  @Error 1 @Extended 4 Return 0 = $iWidth not an Integer, less than 5 or greater than 180.
;                  @Error 1 @Extended 5 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 6 Return 0 = $iHeight not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iPosition not an Integer, less than 0 or greater than 2. See constants, $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameColumnSeparator(ByRef $oFrame, $bSeparatorOn = Null, $iStyle = Null, $iWidth = Null, $iColor = Null, $iHeight = Null, $iPosition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns
	Local $iError = 0
	Local $avColumnLine[6]

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oTextColumns = $oFrame.TextColumns()
	If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bSeparatorOn, $iStyle, $iWidth, $iColor, $iHeight, $iPosition) Then
		__LO_ArrayFill($avColumnLine, $oTextColumns.SeparatorLineIsOn(), $oTextColumns.SeparatorLineStyle(), $oTextColumns.SeparatorLineWidth(), _
				$oTextColumns.SeparatorLineColor(), $oTextColumns.SeparatorLineRelativeHeight(), $oTextColumns.SeparatorLineVerticalAlignment())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColumnLine)
	EndIf

	If ($bSeparatorOn <> Null) Then
		If Not IsBool($bSeparatorOn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oTextColumns.SeparatorLineIsOn = $bSeparatorOn
		$iError = ($oTextColumns.SeparatorLineIsOn() = $bSeparatorOn) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iStyle <> Null) Then
		If Not __LO_IntIsBetween($iStyle, $LOW_LINE_STYLE_NONE, $LOW_LINE_STYLE_DASHED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTextColumns.SeparatorLineStyle = $iStyle
		$iError = ($oTextColumns.SeparatorLineStyle() = $iStyle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 5, 180) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTextColumns.SeparatorLineWidth = $iWidth
		$iError = (__LO_IntIsBetween($oTextColumns.SeparatorLineWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTextColumns.SeparatorLineColor = $iColor
		$iError = ($oTextColumns.SeparatorLineColor() = $iColor) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iHeight <> Null) Then
		If Not __LO_IntIsBetween($iHeight, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTextColumns.SeparatorLineRelativeHeight = $iHeight
		$iError = ($oTextColumns.SeparatorLineRelativeHeight() = $iHeight) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iPosition <> Null) Then
		If Not __LO_IntIsBetween($iPosition, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oTextColumns.SeparatorLineVerticalAlignment = $iPosition
		$iError = ($oTextColumns.SeparatorLineVerticalAlignment() = $iPosition) ? ($iError) : (BitOR($iError, 32))
	EndIf

	$oFrame.TextColumns = $oTextColumns

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameColumnSeparator

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameColumnSettings
; Description ...: Set or retrieve Frame Column count.
; Syntax ........: _LOWriter_FrameColumnSettings(ByRef $oFrame[, $iColumns = Null])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iColumns            - [optional] an integer value. Default is Null. The number of columns that you want in the Frame. Min. 1.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumns not an Integer, or less than 1.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Text Columns Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iColumns
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current column count.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameColumnSettings(ByRef $oFrame, $iColumns = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns
	Local $iError = 0

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oTextColumns = $oFrame.TextColumns()
	If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iColumns) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oTextColumns.ColumnCount())

	If Not __LO_IntIsBetween($iColumns, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextColumns.ColumnCount = $iColumns
	$oFrame.TextColumns = $oTextColumns

	$iError = ($oFrame.TextColumns.ColumnCount() = $iColumns) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameColumnSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameColumnSize
; Description ...: Set or retrieve Frame Column sizing settings.
; Syntax ........: _LOWriter_FrameColumnSize(ByRef $oFrame, $iColumn[, $bAutoWidth = Null[, $iGlobalSpacing = Null[, $iSpacing = Null[, $iWidth = Null]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iColumn             - an integer value. The column to modify the settings on. See Remarks.
;                  $bAutoWidth          - [optional] a boolean value. Default is Null. If True Column Width is automatically adjusted.
;                  $iGlobalSpacing      - [optional] an integer value. Default is Null. Set a spacing value for between all columns. Set in Hundredths of a Millimeter (HMM). See remarks.
;                  $iSpacing            - [optional] an integer value. Default is Null. The Space between two columns, in Hundredths of a Millimeter (HMM). Cannot be set for the last column.
;                  $iWidth              - [optional] an integer value. Default is Null. If $iGlobalSpacing is set to other than 0, enter the width of the column. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iColumn greater than number of columns in the document or less than 1.
;                  @Error 1 @Extended 4 Return 0 = $bAutoWidth not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iGlobalSpacing not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iSpacing not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iWidth not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Text Columns Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Frame Style Column Object Array.
;                  @Error 3 @Extended 3 Return 0 = No columns present for requested Frame.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Array of Columns.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoWidth
;                  |                               2 = Error setting $iGlobalSpacing
;                  |                               4 = Error setting $iSpacing
;                  |                               8 = Error setting $iWidth
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work fine for setting AutoWidth, and Spacing values, however Width will not work the best, Spacing etc is set in plain Hundredths of a Millimeter (HMM) values, however width is set in a relative value, and I am unable to find a way to be able to convert a specific value, such as 1" (2540 HMM) etc, to the appropriate relative value, especially when spacing is set.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  To set $bAutoWidth or $iGlobalSpacing you may enter any number in $iColumn as long as you are not setting width or spacing, as AutoWidth is not column specific. If you set a value for $iGlobalSpacing with $bAutoWidth set to False, the value is applied to all the columns still.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameColumnSize(ByRef $oFrame, $iColumn, $bAutoWidth = Null, $iGlobalSpacing = Null, $iSpacing = Null, $iWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns
	Local $atColumns
	Local $iError = 0, $iRightMargin, $iLeftMargin
	Local $avColumnSize[4]

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextColumns = $oFrame.TextColumns()
	If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$atColumns = $oTextColumns.Columns()
	If Not IsArray($atColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($oTextColumns.ColumnCount() <= 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If ($iColumn > UBound($atColumns)) Or ($iColumn < 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

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
		If Not IsBool($bAutoWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If ($bAutoWidth <> $oTextColumns.IsAutomatic()) Then
			If ($bAutoWidth = True) Then
				; retrieve both outside column inner margin settings to add together for determining AutoWidth value.
				$iGlobalSpacing = ($iGlobalSpacing = Null) ? ($atColumns[0].RightMargin() + $atColumns[UBound($atColumns) - 1].LeftMargin()) : ($iGlobalSpacing)
				; If $iGlobalSpacing is not called with a value, set my own, else use the called value.
				$oTextColumns.ColumnCount = $oTextColumns.ColumnCount()
				$oFrame.TextColumns = $oTextColumns
				; Setting the number of columns activates the AutoWidth option, so set it to the same number of columns.

			Else ; If False
				; If GlobalSpacing isn't set, then set it myself to the current automatic distance.
				$iGlobalSpacing = ($iGlobalSpacing = Null) ? ($oTextColumns.AutomaticDistance()) : ($iGlobalSpacing)
				$oTextColumns.setColumns($atColumns) ; Inserting the Column Array(Sequence) again, even without changes, deactivates AutoWidth.
			EndIf
		EndIf

		$oFrame.TextColumns = $oTextColumns
		$iError = ($oFrame.TextColumns.IsAutomatic() = $bAutoWidth) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iGlobalSpacing <> Null) Then
		If Not IsInt($iGlobalSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTextColumns.AutomaticDistance = $iGlobalSpacing
		$oFrame.TextColumns = $oTextColumns

		If ($oFrame.TextColumns.IsAutomatic() = True) Then ; If AutoWidth is on (True) Then error test, else don't, because I use $iGlobalSpacing
			; for setting the width internally also.
			$iError = (__LO_IntIsBetween($oFrame.TextColumns.AutomaticDistance(), $iGlobalSpacing - 2, $iGlobalSpacing + 2)) ? ($iError) : (BitOR($iError, 2))
		EndIf
	EndIf

	If ($iSpacing <> Null) Then
		If Not IsInt($iSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

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
			$oFrame.TextColumns = $oTextColumns

			; Retrieve Array of columns again for testing.
			$atColumns = $oTextColumns.Columns()
			If Not IsArray($atColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			; See if setting spacing worked. Spacing is equally divided between the two adjoining columns, so retrieve the first columns right
			; margin, and the next column's left margin.
			$iError = (__LO_IntIsBetween($atColumns[$iColumn].RightMargin() + $atColumns[$iColumn + 1].LeftMargin(), $iSpacing - 1, $iSpacing + 1)) ? ($iError) : (BitOR($iError, 4))
		EndIf
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$atColumns[$iColumn].Width = $iWidth

		; Set the settings into the document.
		$oTextColumns.setColumns($atColumns)
		$oFrame.TextColumns = $oTextColumns

		; Retrieve Array of columns again for testing.
		$atColumns = $oFrame.TextColumns.Columns()
		If Not IsArray($atColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($atColumns[$iColumn].Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 8)))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameColumnSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameCreate
; Description ...: Create and insert a TextFrame.
; Syntax ........: _LOWriter_FrameCreate(ByRef $oDoc, ByRef $oCursor[, $sFrameName = Null[, $iWidth = Null[, $iHeight = Null[, $bOverwrite = False]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions. Must not be a Table Cursor.
;                  $sFrameName          - [optional] a string value. Default is Null. The Name of the Frame to Create.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Frame to create. Set in Hundredths of a Millimeter (HMM). Min. 51 (HMM).
;                  $iHeight             - [optional] an integer value. Default is Null. The Height of the Frame to create. Set in Hundredths of a Millimeter (HMM). Min. 51 (HMM).
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, content selected by Cursor is overwritten., Else Frame is inserted after the selection.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, cannot insert a Frame using a Table Cursor.
;                  @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $sFrameName not a String.
;                  @Error 1 @Extended 6 Return 0 = Document already contains a Frame with same name as $sFrameName.
;                  @Error 1 @Extended 7 Return 0 = $iWidth not an Integer, or less than 51 Hundredths of a Millimeter (HMM).
;                  @Error 1 @Extended 8 Return 0 = $iHeight not an Integer, or less than 51 Hundredths of a Millimeter (HMM).
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextFrame" Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Frame was created successfully and inserted at cursor position. Returning Frame Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LO_UnitConvert, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameCreate(ByRef $oDoc, ByRef $oCursor, $sFrameName = Null, $iWidth = Null, $iHeight = Null, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iCONST_AutoHW_OFF = 1
	Local $oFrame

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oFrame = $oDoc.createInstance("com.sun.star.text.TextFrame")
	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sFrameName <> Null) Then
		If Not IsString($sFrameName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If _LOWriter_FrameExists($oDoc, $sFrameName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrame.Name = $sFrameName
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrame.WidthType = $iCONST_AutoHW_OFF
		$oFrame.Width = $iWidth
	EndIf

	If ($iHeight <> Null) Then
		If Not __LO_IntIsBetween($iHeight, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFrame.SizeType = $iCONST_AutoHW_OFF
		$oFrame.Height = $iHeight
	EndIf

	$oDoc.Text.insertTextContent($oCursor, $oFrame, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFrame)
EndFunc   ;==>_LOWriter_FrameCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameCreateTextCursor
; Description ...: Create a Text Cursor in a Frame for inserting text etc.
; Syntax ........: _LOWriter_FrameCreateTextCursor(ByRef $oFrame)
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. A Text Cursor Object located in the Frame.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LOWriter_CursorMove _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameCreateTextCursor(ByRef $oFrame)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFrame.Text.createTextCursor())
EndFunc   ;==>_LOWriter_FrameCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameDelete
; Description ...: Delete a Frame from the document.
; Syntax ........: _LOWriter_FrameDelete(ByRef $oDoc, ByRef $oFrame)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrame not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Frame was attempted to be deleted, but the document still contains a frame named the same.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Frame was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameDelete(ByRef $oDoc, ByRef $oFrame)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sFrameName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sFrameName = $oFrame.getName()
	$oFrame.dispose()
	If ($oDoc.TextFrames.hasByName($sFrameName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Document still contains Frame named the same.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FrameDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameExists
; Description ...: Check if a Document contains a Frame with the specified name.
; Syntax ........: _LOWriter_FrameExists(ByRef $oDoc, $sFrameName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFrameName          - a string value. The Frame name to search for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFrameName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Text Frames Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Shapes Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Search was successful, If Frame was found matching $sFrameName True is Returned, else False
;                  @Error 0 @Extended 1 Return Boolean = Success. Search was successful, Frame found matching $sFrameName listed as a shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Some document types, such as docx, list frames as Shapes instead of TextFrames, so this function searches both.
; Related .......: _LOWriter_FrameDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameExists(ByRef $oDoc, $sFrameName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFrames, $oShapes

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFrameName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oFrames = $oDoc.TextFrames()
	If Not IsObj($oFrames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oFrames.hasByName($sFrameName)) Then Return SetError($__LO_STATUS_SUCCESS, 1, True)

	; If No results, then search Shapes.
	$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If $oShapes.hasElements() Then
		For $i = 0 To $oShapes.getCount() - 1
			If ($oShapes.getByIndex($i).Name() = $sFrameName) Then
				If ($oShapes.getByIndex($i).supportsService("com.sun.star.drawing.Text")) And _
						($oShapes.getByIndex($i).Text.ImplementationName() = "SwXTextFrame") And Not _
						$oShapes.getByIndex($i).getPropertySetInfo().hasPropertyByName("ActualSize") Then Return SetError($__LO_STATUS_SUCCESS, 2, True)
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; No matches
EndFunc   ;==>_LOWriter_FrameExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameGetAnchor
; Description ...: Create a Text Cursor at the Frame Anchor position.
; Syntax ........: _LOWriter_FrameGetAnchor(ByRef $oFrame)
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to retrieve Frame anchor Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully returned the Frame Anchor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameGetAnchor(ByRef $oFrame)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnchor

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnchor = $oFrame.Anchor.Text.createTextCursorByRange($oFrame.Anchor())
	If Not IsObj($oAnchor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oAnchor)
EndFunc   ;==>_LOWriter_FrameGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameGetObjByCursor
; Description ...: Returns a Frame Object, for later Frame related functions.
; Syntax ........: _LOWriter_FrameGetObjByCursor(ByRef $oDoc, ByRef $oCursor)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions. Cursor object must be located in a Frame.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oCursor not located in a Frame.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning an Object for the requested Frame.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_FrameDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameGetObjByCursor(ByRef $oDoc, ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetDataType($oDoc, $oCursor) <> $LOW_CURDATA_FRAME) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Cursor not in Frame

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.TextFrames.getByName($oCursor.TextFrame.Name))
EndFunc   ;==>_LOWriter_FrameGetObjByCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameGetObjByName
; Description ...: Retrieve a Frame Object by Name.
; Syntax ........: _LOWriter_FrameGetObjByName(ByRef $oDoc, $sFrameName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFrameName          - a string value. The frame name to search for.
; Return values .: Success: 0 or Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFrameName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving TextFrame Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Shapes Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. No matches found.
;                  @Error 0 @Extended 1 Return Object = Success. Successfully found requested Frame by name, returning Frame Object.
;                  @Error 0 @Extended 2 Return Object = Success. Successfully found requested Frame by name in Shapes list, returning Frame Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FramesGetNames, _LOWriter_FrameDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameGetObjByName(ByRef $oDoc, $sFrameName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFrames, $oShapes

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFrameName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oFrames = $oDoc.TextFrames()
	If Not IsObj($oFrames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oFrames.hasByName($sFrameName)) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oFrames.getByName($sFrameName))

	; If No results, then search Shapes.
	$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If $oShapes.hasElements() Then
		For $i = 0 To $oShapes.getCount() - 1
			If ($oShapes.getByIndex($i).Name() = $sFrameName) Then
				If ($oShapes.getByIndex($i).Text.ImplementationName() = "SwXTextFrame") Then Return SetError($__LO_STATUS_SUCCESS, 2, $oShapes.getByIndex($i))
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1) ; No matches
EndFunc   ;==>_LOWriter_FrameGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameHyperlink
; Description ...: Set or Retrieve Frame Hyperlink settings.
; Syntax ........: _LOWriter_FrameHyperlink(ByRef $oFrame[, $sURL = Null[, $sName = Null[, $sFrameTarget = Null[, $bServerSideMap = Null]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $sURL                - [optional] a string value. Default is Null. The complete path to the file that you want to open.
;                  $sName               - [optional] a string value. Default is Null. Name for the hyperlink.
;                  $sFrameTarget        - [optional] a string value. Default is Null. Specify the name of the frame where you want to open the targeted file. See Constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bServerSideMap      - [optional] a boolean value. Default is Null. If True, Uses a server-side image map.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sURL not a String
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sFrameTarget not a String.
;                  @Error 1 @Extended 5 Return 0 = $sFrameTarget not equal to one of the Constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bServerSideMap not a boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sURL
;                  |                               2 = Error setting $sName
;                  |                               4 = Error setting $sFrameTarget
;                  |                               8 = Error setting $bServerSideMap
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameHyperlink(ByRef $oFrame, $sURL = Null, $sName = Null, $sFrameTarget = Null, $bServerSideMap = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHyperlink[4]

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sURL, $sName, $sFrameTarget, $bServerSideMap) Then
		__LO_ArrayFill($avHyperlink, $oFrame.HyperLinkURL(), $oFrame.HyperLinkName(), $oFrame.HyperLinkTarget(), $oFrame.ServerMap())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avHyperlink)
	EndIf

	If ($sURL <> Null) Then
		If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oFrame.HyperLinkURL = $sURL
		$iError = ($oFrame.HyperLinkURL() = $sURL) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrame.HyperLinkName = $sName
		$iError = ($oFrame.HyperLinkName = $sName) ? ($iError) : (BitOR($iError, 2))
	EndIf

	; "" ; "_top" ; "_parent" ; "_blank" ; "_self"
	If ($sFrameTarget <> Null) Then
		If Not IsString($sFrameTarget) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If ($sFrameTarget <> "") Then
			If ($sFrameTarget <> $LOW_FRAME_TARGET_TOP) And _
					($sFrameTarget <> $LOW_FRAME_TARGET_PARENT) And _
					($sFrameTarget <> $LOW_FRAME_TARGET_BLANK) And _
					($sFrameTarget <> $LOW_FRAME_TARGET_SELF) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		EndIf
		$oFrame.HyperLinkTarget = $sFrameTarget
		$iError = ($oFrame.HyperLinkTarget() = $sFrameTarget) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bServerSideMap <> Null) Then
		If Not IsBool($bServerSideMap) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrame.ServerMap = $bServerSideMap
		$iError = ($oFrame.ServerMap() = $bServerSideMap) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameHyperlink

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameOptions
; Description ...: Set or Retrieve Frame Options.
; Syntax ........: _LOWriter_FrameOptions(ByRef $oFrame[, $bProtectContent = Null[, $bProtectPos = Null[, $bProtectSize = Null[, $iVertAlign = Null[, $bEditInRead = Null[, $bPrint = Null[, $iTxtDirection = Null]]]]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $bProtectContent     - [optional] a boolean value. Default is Null. If True, Prevents changes to the contents of the frame.
;                  $bProtectPos         - [optional] a boolean value. Default is Null. If True, Locks the position of the frame in the current document.
;                  $bProtectSize        - [optional] a boolean value. Default is Null. If True, Locks the size of the frame.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. Specifies the vertical alignment of the frame's content. See Constants, $LOW_TXT_ADJ_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEditInRead         - [optional] a boolean value. Default is Null. If True, Allows you to edit the contents of a frame in a document that is read-only.
;                  $bPrint              - [optional] a boolean value. Default is Null. If True, Includes the frame when you print the document.
;                  $iTxtDirection       - [optional] an integer value (0-5). Default is Null. Specifies the preferred text flow direction in a frame. See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bProtectContent not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bProtectPos not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bProtectSize not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants, $LOW_TXT_ADJ_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bEditInRead not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bPrint not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iTxtDirection not an Integer, less than 0 or greater than 5. See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bProtectContent
;                  |                               2 = Error setting $bProtectPos
;                  |                               4 = Error setting $bProtectSize
;                  |                               8 = Error setting $iVertAlign
;                  |                               16 = Error setting $bEditInRead
;                  |                               32 = Error setting $bPrint
;                  |                               64 = Error setting $iTxtDirection
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameOptions(ByRef $oFrame, $bProtectContent = Null, $bProtectPos = Null, $bProtectSize = Null, $iVertAlign = Null, $bEditInRead = Null, $bPrint = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOptions[7]

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bProtectContent, $bProtectPos, $bProtectSize, $iVertAlign, $bEditInRead, $bPrint, $iTxtDirection) Then
		__LO_ArrayFill($avOptions, $oFrame.ContentProtected(), $oFrame.PositionProtected(), $oFrame.SizeProtected(), _
				$oFrame.TextVerticalAdjust(), $oFrame.EditInReadOnly(), $oFrame.Print(), $oFrame.WritingMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOptions)
	EndIf

	If ($bProtectContent <> Null) Then
		If Not IsBool($bProtectContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oFrame.ContentProtected = $bProtectContent
		$iError = ($oFrame.ContentProtected() = $bProtectContent) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bProtectPos <> Null) Then
		If Not IsBool($bProtectPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrame.PositionProtected = $bProtectPos
		$iError = ($oFrame.PositionProtected() = $bProtectPos) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bProtectSize <> Null) Then
		If Not IsBool($bProtectSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrame.SizeProtected = $bProtectSize
		$iError = ($oFrame.SizeProtected() = $bProtectSize) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_TXT_ADJ_VERT_TOP, $LOW_TXT_ADJ_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrame.TextVerticalAdjust = $iVertAlign
		$iError = ($oFrame.TextVerticalAdjust() = $iVertAlign) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEditInRead <> Null) Then
		If Not IsBool($bEditInRead) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrame.EditInReadOnly = $bEditInRead
		$iError = ($oFrame.EditInReadOnly() = $bEditInRead) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bPrint <> Null) Then
		If Not IsBool($bPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrame.Print = $bPrint
		$iError = ($oFrame.Print() = $bPrint) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iTxtDirection <> Null) Then
		If Not __LO_IntIsBetween($iTxtDirection, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFrame.WritingMode = $iTxtDirection
		$iError = ($oFrame.WritingMode() = $iTxtDirection) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameOptions

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameOptionsName
; Description ...: Set or Retrieve Frame Name settings.
; Syntax ........: _LOWriter_FrameOptionsName(ByRef $oDoc, ByRef $oFrame[, $sName = Null[, $sDesc = Null[, $sPrevLink = Null[, $sNextLink = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $sName               - [optional] a string value. Default is Null. Name for the Frame.
;                  $sDesc               - [optional] a string value. Default is Null. Description of the Frame.
;                  $sPrevLink           - [optional] a string value. Default is Null. The Frame Name that comes before the current Frame in a linked sequence. The current frame and the target frame must be empty. Call with "" to remove a linked frame.
;                  $sNextLink           - [optional] a string value. Default is Null. The Frame Name that comes after the current Frame in a linked sequence. The current frame and the target frame must be empty. Call with "" to remove a linked frame.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = Document already contains Frame with same name as $sName.
;                  @Error 1 @Extended 5 Return 0 = $sDesc not a string.
;                  @Error 1 @Extended 6 Return 0 = $sPrevLink not a String.
;                  @Error 1 @Extended 7 Return 0 = Document does not contain Frame matching $sPrevLink.
;                  @Error 1 @Extended 8 Return 0 = $sNextLink not a String.
;                  @Error 1 @Extended 9 Return 0 = Document does not contain Frame matching $sNextLink
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sDesc
;                  |                               4 = Error setting $sPrevLink
;                  |                               8 = Error setting $sNextLink
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameOptionsName(ByRef $oDoc, ByRef $oFrame, $sName = Null, $sDesc = Null, $sPrevLink = Null, $sNextLink = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asName[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sName, $sDesc, $sPrevLink, $sNextLink) Then
		__LO_ArrayFill($asName, $oFrame.Name(), $oFrame.Description(), $oFrame.ChainPrevName(), $oFrame.ChainNextName())

		Return SetError($__LO_STATUS_SUCCESS, 1, $asName)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If _LOWriter_FrameExists($oDoc, $sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrame.Name = $sName
		$iError = ($oFrame.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sDesc <> Null) Then
		If Not IsString($sDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrame.Description = $sDesc
		$iError = ($oFrame.Description = $sDesc) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sPrevLink <> Null) Then
		If Not IsString($sPrevLink) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If ($sPrevLink <> "") And Not _LOWriter_FrameExists($oDoc, $sPrevLink) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrame.ChainPrevName = $sPrevLink
		$iError = ($oFrame.ChainPrevName() = $sPrevLink) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($sNextLink <> Null) Then
		If Not IsString($sNextLink) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		If ($sNextLink <> "") And Not _LOWriter_FrameExists($oDoc, $sNextLink) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oFrame.ChainNextName = $sNextLink
		$iError = ($oFrame.ChainNextName() = $sNextLink) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameOptionsName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FramesGetNames
; Description ...: Retrieve an array of names of all frames contained in a document.
; Syntax ........: _LOWriter_FramesGetNames(ByRef $oDoc[, $bSearchShapes = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bSearchShapes       - [optional] a boolean value. Default is False. If True, function searches and adds any Frames listed as "Shapes" in the document to the array of Frame names. See remarks.
; Return values .: Success: Array of Strings.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bSearchShapes not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failure retrieving Frame objects.
;                  @Error 3 @Extended 2 Return 0 = Failure retrieving Shape objects.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Returning Array of Frame names. @Extended set to number of Frame Names returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In Docx (and possibly other formats) Frames seem to be saved as "Shapes" instead of "Frames", if this function returns no results, or not the ones you expect, try setting $bSearchShapes to True.
; Related .......: _LOWriter_FrameGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FramesGetNames(ByRef $oDoc, $bSearchShapes = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asFrameNames[0], $asShapes[0]
	Local $oFrames, $oShapes
	Local $iCount = 0, $iEndofArray

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bSearchShapes) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oFrames = $oDoc.TextFrames()
	If Not IsObj($oFrames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oFrames.hasElements() Then
		ReDim $asFrameNames[$oFrames.getCount()]
		For $i = 0 To $oFrames.getCount() - 1
			$asFrameNames[$i] = $oFrames.getByIndex($i).Name()
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	If ($bSearchShapes = True) Then
		$oShapes = $oDoc.DrawPage()
		If Not IsObj($oShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		If $oShapes.hasElements() Then
			ReDim $asShapes[$oShapes.getCount()]
			For $i = 0 To $oShapes.getCount() - 1
				If $oShapes.getByIndex($i).supportsService("com.sun.star.drawing.Text") Then ; Determine if the Shape is an actual Frame or not.
					If ($oShapes.getByIndex($i).Text.ImplementationName() = "SwXTextFrame") And Not _
							$oShapes.getByIndex($i).getPropertySetInfo().hasPropertyByName("ActualSize") Then
						$asShapes[$iCount] = $oShapes.getByIndex($i).Name()
						$iCount += 1
					EndIf
				EndIf
			Next

			ReDim $asShapes[$iCount]

			$iEndofArray = UBound($asFrameNames)
			ReDim $asFrameNames[UBound($asFrameNames) + $iCount]

			For $i = 0 To UBound($asShapes) - 1
				$asFrameNames[$iEndofArray] = $asShapes[$i]
				$iEndofArray += 1
			Next
		EndIf
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, UBound($asFrameNames), $asFrameNames)
EndFunc   ;==>_LOWriter_FramesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameShadow
; Description ...: Set or Retrieve the shadow settings for a Frame.
; Syntax ........: _LOWriter_FrameShadow(ByRef $oFrame[, $iWidth = Null[, $iColor = Null[, $iLocation = Null]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Width of the Frame Shadow set in Hundredths of a Millimeter (HMM).
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Color of the Frame shadow, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3..
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Frame Shadow, must be one of the Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3..
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3..
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
;                  LibreOffice may change the shadow width +/- a Hundredth of a Millimeter.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, _LOWriter_FrameGetObjByCursor, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameShadow(ByRef $oFrame, $iWidth = Null, $iColor = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tShdwFrmt
	Local $iError = 0
	Local $avShadow[3]

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tShdwFrmt = $oFrame.ShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.Location())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Or ($iWidth < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

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

	$oFrame.ShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oFrame.ShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($tShdwFrmt.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iColor)) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iLocation)) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 4)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleAreaColor
; Description ...: Set or Retrieve background color settings for a Frame style.
; Syntax ........: _LOWriter_FrameStyleAreaColor(ByRef $oFrameStyle[, $iBackColor = Null])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Background color.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleAreaColor(ByRef $oFrameStyle, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency
	Local $iColor

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iBackColor) Then
		$iColor = __LOWriter_ColorRemoveAlpha($oFrameStyle.BackColor())
		If Not IsInt($iColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iOldTransparency = $oFrameStyle.FillTransparence()
	If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oFrameStyle.BackColor = $iBackColor
	$iError = ($oFrameStyle.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))

	$oFrameStyle.FillTransparence = $iOldTransparency

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOWriter_FrameStyleAreaFillStyle(ByRef $oFrameStyle)
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
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
Func _LOWriter_FrameStyleAreaFillStyle(ByRef $oFrameStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oFrameStyle.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOWriter_FrameStyleAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleAreaGradient
; Description ...: Modify or retrieve the settings for Frame Style Background color Gradient.
; Syntax ........: _LOWriter_FrameStyleAreaGradient(ByRef $oDoc, ByRef $oFrameStyle[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See remarks. See constants, $LOW_GRAD_NAME_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient that you want to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iIncrement          - [optional] an integer value. Default is Null. Specifies the number of steps of change color. Allowed values are: 0, 3 to 256. 0 = Automatic.
;                  $iXCenter            - [optional] an integer value. Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage, Min. 0%, Max 100%. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value. Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage, Min. 0%, Max 100%. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value. Default is Null. The rotation angle for the gradient. Set in degrees, min 0, max 359 degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value. Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage. Minimum is 0, Maximum is 100%.
;                  $iFromColor          - [optional] an integer value (0-16777215). Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] an integer value (0-16777215). Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] an integer value. Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color. Min. 0%, Max 100%
;                  $iToIntense          - [optional] an integer value. Default is Null . Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color. Min. 0%, Max 100%
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 4 Return 0 = $sGradientName Not a String.
;                  @Error 1 @Extended 5 Return 0 = $iType Not an Integer, less than -1 or greater than 5. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3..
;                  @Error 1 @Extended 6 Return 0 = $iIncrement Not an Integer, less than 3 but not 0, or greater than 256.
;                  @Error 1 @Extended 7 Return 0 = $iXCenter Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iYCenter Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iAngle Not an Integer, less than 0 or greater than 359.
;                  @Error 1 @Extended 10 Return 0 = $iTransitionStart Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 11 Return 0 = $iFromColor Not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 12 Return 0 = $iToColor Not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 13 Return 0 = $iFromIntense Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 14 Return 0 = $iToIntense Not an Integer, less than 0 or greater than 100.
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
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleAreaGradient(ByRef $oDoc, ByRef $oFrameStyle, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName
	Local $atColorStop[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tStyleGradient = $oFrameStyle.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oFrameStyle.FillGradientName(), $tStyleGradient.Style(), _
				$oFrameStyle.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oFrameStyle.FillStyle() <> $LOW_AREA_FILL_STYLE_GRADIENT) Then $oFrameStyle.FillStyle = $LOW_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		__LOWriter_GradientPresets($oDoc, $oFrameStyle, $tStyleGradient, $sGradientName)
		$iError = ($oFrameStyle.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oFrameStyle.FillStyle = $LOW_AREA_FILL_STYLE_OFF
			$oFrameStyle.FillGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrameStyle.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oFrameStyle.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
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

	If ($oFrameStyle.FillGradientName() = "") Or __LOWriter_GradientIsModified($tStyleGradient, $oFrameStyle.FillGradientName()) Then
		$sGradName = __LOWriter_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oFrameStyle.FillGradientName = $sGradName
		If ($oFrameStyle.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oFrameStyle.FillGradient = $tStyleGradient

	; Error checking
	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oFrameStyle.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oFrameStyle.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oFrameStyle.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oFrameStyle.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oFrameStyle.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = (__LO_VarsAreNull($iFromColor)) ? ($iError) : (($oFrameStyle.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = (__LO_VarsAreNull($iToColor)) ? ($iError) : (($oFrameStyle.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = (__LO_VarsAreNull($iFromIntense)) ? ($iError) : (($oFrameStyle.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = (__LO_VarsAreNull($iToIntense)) ? ($iError) : (($oFrameStyle.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleAreaGradientMulticolor
; Description ...: Set or Retrieve a Frame Style's Multicolor Gradient settings. See remarks.
; Syntax ........: _LOWriter_FrameStyleAreaGradientMulticolor(ByRef $oFrameStyle[, $avColorStops = Null])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
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
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, as a RGB Color Integer.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOWriter_GradientMulticolorAdd, _LOWriter_GradientMulticolorDelete, _LOWriter_GradientMulticolorModify, _LOWriter_FrameStyleAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleAreaGradientMulticolor(ByRef $oFrameStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oFrameStyle.FillGradient()
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
	$oFrameStyle.FillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oFrameStyle.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleAreaTransparency
; Description ...: Modify or retrieve Transparency settings for a Frame style.
; Syntax ........: _LOWriter_FrameStyleAreaTransparency(ByRef $oFrameStyle[, $iTransparency = Null])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
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
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleAreaTransparency(ByRef $oFrameStyle, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oFrameStyle.FillTransparence())

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oFrameStyle.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oFrameStyle.FillTransparence = $iTransparency
	$iError = ($oFrameStyle.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleAreaTransparencyGradient
; Description ...: Modify or retrieve the Frame Style transparency gradient settings.
; Syntax ........: _LOWriter_FrameStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oFrameStyle[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Call with $LOW_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 4 Return 0 = $iType not an Integer, less than -1 or greater than 5. See constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error 1 @Extended 8 Return 0 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iStart not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 10 Return 0 = $iEnd not an Integer, less than 0 or greater than 100.
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
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oFrameStyle, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $sTGradName
	Local $aiTransparent[7]
	Local $atColorStop
	Local $fValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tStyleGradient = $oFrameStyle.FillTransparenceGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		__LO_ArrayFill($aiTransparent, $tStyleGradient.Style(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), _
				Int($tStyleGradient.Angle() / 10), $tStyleGradient.Border(), __LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.StartColor()), _
				__LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oFrameStyle.FillTransparenceGradientName = ""

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

			$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; Last StopOffset is the "End" Value.

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

	If ($oFrameStyle.FillTransparenceGradientName() = "") Then
		$sTGradName = __LOWriter_TransparencyGradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oFrameStyle.FillTransparenceGradientName = $sTGradName
		If ($oFrameStyle.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oFrameStyle.FillTransparenceGradient = $tStyleGradient

	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oFrameStyle.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oFrameStyle.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oFrameStyle.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oFrameStyle.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oFrameStyle.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iStart)) ? ($iError) : (($oFrameStyle.FillTransparenceGradient.StartColor() = __LOWriter_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iEnd)) ? ($iError) : (($oFrameStyle.FillTransparenceGradient.EndColor() = __LOWriter_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Frame Style's Multi Transparency Gradient settings. See remarks.
; Syntax ........: _LOWriter_FrameStyleAreaTransparencyGradientMulti(ByRef $oFrameStyle[, $avColorStops = Null])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
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
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple Transparency stops in a Gradient rather than just a beginning and an ending value, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the Transparency percentage value between 0 and 100%.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOWriter_TransparencyGradientMultiModify, _LOWriter_TransparencyGradientMultiDelete, _LOWriter_TransparencyGradientMultiAdd, _LOWriter_FrameStyleAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleAreaTransparencyGradientMulti(ByRef $oFrameStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oFrameStyle.FillTransparenceGradient()
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
	$oFrameStyle.FillTransparenceGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oFrameStyle.FillTransparenceGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleBorderColor
; Description ...: Set or retrieve the Frame Style Border Line Color. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_FrameStyleBorderColor(ByRef $oFrameStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. The Top Border Line Color of the Frame, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. The Bottom Border Line Color of the Frame, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. The Left Border Line Color of the Frame, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. The Right Border Line Color of the Frame, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
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
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_FrameStyleBorderWidth, _LOWriter_FrameStyleBorderStyle, _LOWriter_FrameStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleBorderColor(ByRef $oFrameStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oFrameStyle, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FrameStyleBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleBorderPadding
; Description ...: Set or retrieve the Frame Style Border Padding settings.
; Syntax ........: _LOWriter_FrameStyleBorderPadding(ByRef $oFrameStyle[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The Top Distance between the Border and Frame contents in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Distance between the Border and Frame contents in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Distance between the Border and Frame contents in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The Right Distance between the Border and Frame contents in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
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
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_UnitConvert, _LOWriter_FrameStyleBorderWidth, _LOWriter_FrameStyleBorderStyle, _LOWriter_FrameStyleBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleBorderPadding(ByRef $oFrameStyle, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oFrameStyle.BorderDistance(), $oFrameStyle.TopBorderDistance(), _
				$oFrameStyle.BottomBorderDistance(), $oFrameStyle.LeftBorderDistance(), $oFrameStyle.RightBorderDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LO_IntIsBetween($iAll, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrameStyle.BorderDistance = $iAll
		$iError = (__LO_IntIsBetween($oFrameStyle.BorderDistance(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrameStyle.TopBorderDistance = $iTop
		$iError = (__LO_IntIsBetween($oFrameStyle.TopBorderDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrameStyle.BottomBorderDistance = $iBottom
		$iError = (__LO_IntIsBetween($oFrameStyle.BottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrameStyle.LeftBorderDistance = $iLeft
		$iError = (__LO_IntIsBetween($oFrameStyle.LeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrameStyle.RightBorderDistance = $iRight
		$iError = (__LO_IntIsBetween($oFrameStyle.RightBorderDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleBorderStyle
; Description ...: Set or Retrieve the Frame Style Border Line style. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_FrameStyleBorderStyle(ByRef $oFrameStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iTop                - [optional] an integer value (0x7FFF-17). Default is Null. The Top Border Line Style of the Frame. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF-17). Default is Null. The Bottom Border Line Style of the Frame. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF-17). Default is Null. The Left Border Line Style of the Frame. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF-17). Default is Null. The Right Border Line Style of the Frame. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 4 Return 0 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 5 Return 0 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
;                  @Error 1 @Extended 6 Return 0 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF.
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
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LOWriter_FrameStyleBorderWidth, _LOWriter_FrameStyleBorderColor, _LOWriter_FrameStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleBorderStyle(ByRef $oFrameStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oFrameStyle, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FrameStyleBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleBorderWidth
; Description ...: Set or Retrieve the Frame Style Border Line Width. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_FrameStyleBorderWidth(ByRef $oFrameStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iTop                - [optional] an integer value. Default is Null. The Top Border Line width of the Frame in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Border Line Width of the Frame in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Border Line width of the Frame in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. The Right Border Line Width of the Frame in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
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
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set Width to 0
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_UnitConvert, _LOWriter_FrameStyleBorderStyle, _LOWriter_FrameStyleBorderColor, _LOWriter_FrameStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleBorderWidth(ByRef $oFrameStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oFrameStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FrameStyleBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleColumnSeparator
; Description ...: Modify or retrieve Frame Style Column Separator line settings.
; Syntax ........: _LOWriter_FrameStyleColumnSeparator(ByRef $oFrameStyle[, $bSeparatorOn = Null[, $iStyle = Null[, $iWidth = Null[, $iColor = Null[, $iHeight = Null[, $iPosition = Null]]]]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $bSeparatorOn        - [optional] a boolean value. Default is Null. If True, add a separator line between two or more columns.
;                  $iStyle              - [optional] an integer value (0-3). Default is Null. The formatting style for the column separator line. See Constants, $LOW_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iWidth              - [optional] an integer value (5-180). Default is Null. The width of the separator line. Set in Hundredths of a Millimeter (HMM).
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The separator line color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iHeight             - [optional] an integer value (0-100). Default is Null. The length of the separator line as a percentage of the height of the column area.
;                  $iPosition           - [optional] an integer value (0-2). Default is Null. Select the vertical alignment of the separator line. This option is only available if Height value of the line is less than 100%. See Constants, $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $bSeparatorOn not a Boolean value.
;                  @Error 1 @Extended 4 Return 0 = $iStyle not an Integer, less than 0 or greater than 3. See constants, $LOW_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iWidth not an Integer, less than 5 or greater than 180.
;                  @Error 1 @Extended 6 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 7 Return 0 = $iHeight not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iPosition not an Integer, less than 0 or greater than 2. See constants, $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleColumnSeparator(ByRef $oFrameStyle, $bSeparatorOn = Null, $iStyle = Null, $iWidth = Null, $iColor = Null, $iHeight = Null, $iPosition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns
	Local $iError = 0
	Local $avColumnLine[6]

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextColumns = $oFrameStyle.TextColumns()
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

	$oFrameStyle.TextColumns = $oTextColumns

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleColumnSeparator

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleColumnSettings
; Description ...: Set or retrieve Frame style Column count.
; Syntax ........: _LOWriter_FrameStyleColumnSettings(ByRef $oFrameStyle[, $iColumns = Null])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iColumns            - [optional] an integer value. Default is Null. The number of columns that you want in the Frame. Min. 1.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iColumns not an Integer, or less than 1.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Text Columns Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iColumns
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current column count.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleColumnSettings(ByRef $oFrameStyle, $iColumns = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns
	Local $iError = 0

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextColumns = $oFrameStyle.TextColumns()
	If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iColumns) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oTextColumns.ColumnCount())

	If Not __LO_IntIsBetween($iColumns, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextColumns.ColumnCount = $iColumns
	$oFrameStyle.TextColumns = $oTextColumns

	$iError = ($oFrameStyle.TextColumns.ColumnCount() = $iColumns) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleColumnSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleColumnSize
; Description ...: Set or retrieve Column sizing settings.
; Syntax ........: _LOWriter_FrameStyleColumnSize(ByRef $oFrameStyle, $iColumn[, $bAutoWidth = Null[, $iGlobalSpacing = Null[, $iSpacing = Null[, $iWidth = Null]]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iColumn             - an integer value. The column to modify the settings on. See Remarks.
;                  $bAutoWidth          - [optional] a boolean value. Default is Null. If True, Column Width is automatically adjusted.
;                  $iGlobalSpacing      - [optional] an integer value. Default is Null. Set a spacing value for between all columns. Set in Hundredths of a Millimeter (HMM). See remarks.
;                  $iSpacing            - [optional] an integer value. Default is Null. The Space between two columns, in Hundredths of a Millimeter (HMM). Cannot be set for the last column.
;                  $iWidth              - [optional] an integer value. Default is Null. If $iGlobalSpacing is set to other than 0, enter the width of the column. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iColumn not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iColumn greater than number of columns in the document or less than 1.
;                  @Error 1 @Extended 5 Return 0 = $bAutoWidth not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iGlobalSpacing not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iSpacing not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iWidth not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Text Columns Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Frame Style Column Object Array.
;                  @Error 3 @Extended 3 Return 0 = No columns present for requested Frame Style.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve array of Columns.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoWidth
;                  |                               2 = Error setting $iGlobalSpacing
;                  |                               4 = Error setting $iSpacing
;                  |                               8 = Error setting $iWidth
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work fine for setting AutoWidth, and Spacing values, however Width will not work the best, Spacing etc is set in plain Hundredths of a Millimeter (HMM) values, however width is set in a relative value, and I am unable to find a way to be able to convert a specific value, such as 1" (2540 HMM) etc, to the appropriate relative value, especially when spacing is set.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  To set $bAutoWidth or $iGlobalSpacing you may enter any number in $iColumn as long as you are not setting width or spacing, as AutoWidth is not column specific. If you set a value for $iGlobalSpacing with $bAutoWidth set to False, the value is applied to all the columns still.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleColumnSize(ByRef $oFrameStyle, $iColumn, $bAutoWidth = Null, $iGlobalSpacing = Null, $iSpacing = Null, $iWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns
	Local $atColumns
	Local $iError = 0, $iRightMargin, $iLeftMargin
	Local $avColumnSize[4]

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextColumns = $oFrameStyle.TextColumns()
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

		If ($bAutoWidth <> $oTextColumns.IsAutomatic()) Then
			If ($bAutoWidth = True) Then
				; retrieve both outside column inner margin settings to add together for determining AutoWidth value.
				$iGlobalSpacing = ($iGlobalSpacing = Null) ? ($atColumns[0].RightMargin() + $atColumns[UBound($atColumns) - 1].LeftMargin()) : ($iGlobalSpacing)
				; If $iGlobalSpacing is not called with a value, set my own, else use the called value.
				$oTextColumns.ColumnCount = $oTextColumns.ColumnCount()
				$oFrameStyle.TextColumns = $oTextColumns
				; Setting the number of columns activates the AutoWidth option, so set it to the same number of columns.

			Else ; If False
				; If GlobalSpacing isn't set, then set it myself to the current automatic distance.
				$iGlobalSpacing = ($iGlobalSpacing = Null) ? ($oTextColumns.AutomaticDistance()) : ($iGlobalSpacing)
				$oTextColumns.setColumns($atColumns) ; Inserting the Column Array(Sequence) again, even without changes, deactivates AutoWidth.
			EndIf

			$oFrameStyle.TextColumns = $oTextColumns
			$iError = ($oFrameStyle.TextColumns.IsAutomatic() = $bAutoWidth) ? ($iError) : (BitOR($iError, 1))
		EndIf

		If ($iGlobalSpacing <> Null) Then
			If Not IsInt($iGlobalSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oTextColumns.AutomaticDistance = $iGlobalSpacing
			$oFrameStyle.TextColumns = $oTextColumns

			If ($oFrameStyle.TextColumns.IsAutomatic() = True) Then ; If AutoWidth is on (True) Then error test, else don't, because I use $iGlobalSpacing
				; for setting the width internally also.
				$iError = (__LO_IntIsBetween($oFrameStyle.TextColumns.AutomaticDistance(), $iGlobalSpacing - 2, $iGlobalSpacing + 2)) ? ($iError) : (BitOR($iError, 2))
			EndIf
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
			$oFrameStyle.TextColumns = $oTextColumns

			; Retrieve Array of columns again for testing.
			$atColumns = $oTextColumns.Columns()
			If Not IsArray($atColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

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
		$oFrameStyle.TextColumns = $oTextColumns

		; Retrieve Array of columns again for testing.
		$atColumns = $oFrameStyle.TextColumns.Columns()
		If Not IsArray($atColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($atColumns[$iColumn].Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 8)))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleColumnSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleCreate
; Description ...: Create a new Frame Style in a Document.
; Syntax ........: _LOWriter_FrameStyleCreate(ByRef $oDoc, $sFrameStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFrameStyle         - a string value. The Name of the New Frame Style to Create.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFrameStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = $sFrameStyle name already exists in document.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating "com.sun.star.style.FrameStyle" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Retrieving "FrameStyles" Object.
;                  @Error 3 @Extended 2 Return 0 = Error creating new Frame Style by Name.
;                  @Error 3 @Extended 3 Return 0 = Error Retrieving New Frame Style Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. New Frame Style successfully created. Returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FrameStyleDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleCreate(ByRef $oDoc, $sFrameStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFrameStyles, $oStyle, $oFrameStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oFrameStyles = $oDoc.StyleFamilies().getByName("FrameStyles")
	If Not IsObj($oFrameStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If _LOWriter_FrameStyleExists($oDoc, $sFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oStyle = $oDoc.createInstance("com.sun.star.style.FrameStyle")
	If Not IsObj($oStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFrameStyles.insertByName($sFrameStyle, $oStyle)

	If Not $oFrameStyles.hasByName($sFrameStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oFrameStyle = $oFrameStyles.getByName($sFrameStyle)
	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFrameStyle)
EndFunc   ;==>_LOWriter_FrameStyleCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleCurrent
; Description ...: Set or Retrieve the current Frame Style used for a Frame.
; Syntax ........: _LOWriter_FrameStyleCurrent(ByRef $oDoc, ByRef $oFrameObj[, $sFrameStyle = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrameObj           - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $sFrameStyle         - [optional] a string value. Default is Null. The Frame Style name to set the frame to.
; Return values .: Success: 1 or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oFrameObj does not support Base Frame Service, not a Frame Object.
;                  @Error 1 @Extended 4 Return 0 = $sFrameStyle not a String.
;                  @Error 1 @Extended 5 Return 0 = Frame Style called in $sFrameStyle doesn't exist in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Frame Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFrameStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Frame Style successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning current Frame Style set for the Frame.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByCursor, _LOWriter_FrameGetObjByName, _LOWriter_FrameStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleCurrent(ByRef $oDoc, ByRef $oFrameObj, $sFrameStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sCurrStyle
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrameObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oFrameObj.supportsService("com.sun.star.text.BaseFrame") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($sFrameStyle) Then
		$sCurrStyle = $oFrameObj.FrameStyleName()
		If Not IsString($sCurrStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrStyle)
	EndIf

	If Not IsString($sFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not _LOWriter_FrameStyleExists($oDoc, $sFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oFrameObj.FrameStyleName = $sFrameStyle
	$iError = ($oFrameObj.FrameStyleName() = $sFrameStyle) ? ($iError) : (BitOR($iError, 1))

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOWriter_FrameStyleCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleDelete
; Description ...: Delete a User-Created Frame Style from a Document.
; Syntax ........: _LOWriter_FrameStyleDelete(ByRef $oDoc, $oFrameStyle[, $bForceDelete = False[, $sReplacementStyle = "Frame"]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $bForceDelete        - [optional] a boolean value. Default is False. If True, Frame style will be deleted regardless of whether it is in use or not.
;                  $sReplacementStyle   - [optional] a string value. Default is "Frame". The Frame style to use instead of the one being deleted if the Frame style being deleted was already applied to a Frame in the document.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 4 Return 0 = $bForceDelete not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $sReplacementStyle not a String.
;                  @Error 1 @Extended 6 Return 0 = Frame Style called in $sReplacementStyle doesn't exist in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving "FrameStyles" Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Frame Style Name.
;                  @Error 3 @Extended 3 Return 0 = $sFrameStyle is not a User-Created Frame Style and cannot be deleted.
;                  @Error 3 @Extended 4 Return 0 = $sFrameStyle is in use and $bForceDelete is False.
;                  @Error 3 @Extended 5 Return 0 = $sFrameStyle still exists after deletion attempt.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Frame Style called in $sFrameStyle was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleDelete(ByRef $oDoc, $oFrameStyle, $bForceDelete = False, $sReplacementStyle = "Frame")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFrameStyles
	Local $sFrameStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bForceDelete) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplacementStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($sReplacementStyle <> "") And Not _LOWriter_FrameStyleExists($oDoc, $sReplacementStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oFrameStyles = $oDoc.StyleFamilies().getByName("FrameStyles")
	If Not IsObj($oFrameStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sFrameStyle = $oFrameStyle.Name()
	If Not IsString($sFrameStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oFrameStyle.isUserDefined() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If $oFrameStyle.isInUse() And Not ($bForceDelete) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; If Style is in use return an error unless force delete is true.

	If ($oFrameStyle.getParentStyle() = Null) Or ($sReplacementStyle <> "Frame") Then $oFrameStyle.setParentStyle($sReplacementStyle)
	; If Parent style is blank set it to "Frame" style, Or if not but User has called a specific style set it to that.

	$oFrameStyles.removeByName($sFrameStyle)

	Return ($oFrameStyles.hasByName($sFrameStyle)) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleExists
; Description ...: Check whether a Document contains a specific Frame Style by name.
; Syntax ........: _LOWriter_FrameStyleExists(ByRef $oDoc, $sFrameStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFrameStyle         - a string value. The Frame Style Name to search for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFrameStyle not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the Document contains a Frame style matching the input name, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleExists(ByRef $oDoc, $sFrameStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oDoc.StyleFamilies.getByName("FrameStyles").hasByName($sFrameStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOWriter_FrameStyleExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleGetObj
; Description ...: Retrieve a Frame Style Object for use with other Frame Style functions.
; Syntax ........: _LOWriter_FrameStyleGetObj(ByRef $oDoc, $sFrameStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFrameStyle         - a string value. The Frame Style name to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFrameStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = Frame Style called in $sFrameStyle not found in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Frame Style Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Frame Style successfully retrieved, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FrameStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleGetObj(ByRef $oDoc, $sFrameStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFrameStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_FrameStyleExists($oDoc, $sFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oFrameStyle = $oDoc.StyleFamilies().getByName("FrameStyles").getByName($sFrameStyle)
	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFrameStyle)
EndFunc   ;==>_LOWriter_FrameStyleGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleOptions
; Description ...: Set or Retrieve Frame Style Options.
; Syntax ........: _LOWriter_FrameStyleOptions(ByRef $oFrameStyle[, $bProtectContent = Null[, $bProtectPos = Null[, $bProtectSize = Null[, $iVertAlign = Null[, $bEditInRead = Null[, $bPrint = Null[, $iTxtDirection = Null]]]]]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $bProtectContent     - [optional] a boolean value. Default is Null. If True, Prevents changes to the contents of the frame.
;                  $bProtectPos         - [optional] a boolean value. Default is Null. If True, Locks the position of the frame in the current document.
;                  $bProtectSize        - [optional] a boolean value. Default is Null. If True, Locks the size of the frame.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. Specifies the vertical alignment of the frame's content. See Constants, $LOW_TXT_ADJ_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEditInRead         - [optional] a boolean value. Default is Null. If True, Allows you to edit the contents of a frame in a document that is read-only.
;                  $bPrint              - [optional] a boolean value. Default is Null. If True, Includes the Frame when you print the document.
;                  $iTxtDirection       - [optional] an integer value (0-5). Default is Null. Specifies the preferred text flow direction in a frame. See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $bProtectContent not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bProtectPos not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bProtectSize not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants, $LOW_TXT_ADJ_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bEditInRead not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bPrint not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iTxtDirection not an Integer, less than 0 or greater than 5. See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bProtectContent
;                  |                               2 = Error setting $bProtectPos
;                  |                               4 = Error setting $bProtectSize
;                  |                               8 = Error setting $iVertAlign
;                  |                               16 = Error setting $bEditInRead
;                  |                               32 = Error setting $bPrint
;                  |                               64 = Error setting $iTxtDirection
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleOptions(ByRef $oFrameStyle, $bProtectContent = Null, $bProtectPos = Null, $bProtectSize = Null, $iVertAlign = Null, $bEditInRead = Null, $bPrint = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOptions[7]

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bProtectContent, $bProtectPos, $bProtectSize, $iVertAlign, $bEditInRead, $bPrint, $iTxtDirection) Then
		__LO_ArrayFill($avOptions, $oFrameStyle.ContentProtected(), $oFrameStyle.PositionProtected(), $oFrameStyle.SizeProtected(), _
				$oFrameStyle.TextVerticalAdjust(), $oFrameStyle.EditInReadOnly(), $oFrameStyle.Print(), $oFrameStyle.WritingMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOptions)
	EndIf

	If ($bProtectContent <> Null) Then
		If Not IsBool($bProtectContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrameStyle.ContentProtected = $bProtectContent
		$iError = ($oFrameStyle.ContentProtected() = $bProtectContent) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bProtectPos <> Null) Then
		If Not IsBool($bProtectPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrameStyle.PositionProtected = $bProtectPos
		$iError = ($oFrameStyle.PositionProtected() = $bProtectPos) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bProtectSize <> Null) Then
		If Not IsBool($bProtectSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrameStyle.SizeProtected = $bProtectSize
		$iError = ($oFrameStyle.SizeProtected() = $bProtectSize) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_TXT_ADJ_VERT_TOP, $LOW_TXT_ADJ_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrameStyle.TextVerticalAdjust = $iVertAlign
		$iError = ($oFrameStyle.TextVerticalAdjust() = $iVertAlign) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEditInRead <> Null) Then
		If Not IsBool($bEditInRead) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrameStyle.EditInReadOnly = $bEditInRead
		$iError = ($oFrameStyle.EditInReadOnly() = $bEditInRead) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bPrint <> Null) Then
		If Not IsBool($bPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFrameStyle.Print = $bPrint
		$iError = ($oFrameStyle.Print() = $bPrint) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iTxtDirection <> Null) Then
		If Not __LO_IntIsBetween($iTxtDirection, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oFrameStyle.WritingMode = $iTxtDirection
		$iError = ($oFrameStyle.WritingMode() = $iTxtDirection) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleOptions

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleOrganizer
; Description ...: Set or retrieve the Organizer settings of a Frame Style.
; Syntax ........: _LOWriter_FrameStyleOrganizer(ByRef $oDoc, $oFrameStyle[, $sNewFrameStyleName = Null[, $sParentStyle = Null[, $bAutoUpdate = Null[, $bHidden = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $sNewFrameStyleName  - [optional] a string value. Default is Null. The new name to set $sFrameStyle Frame style to.
;                  $sParentStyle        - [optional] a string value. Default is Null. Set an existing Frame style (or an Empty String ("") = - None -) to apply its settings to the current style. Use the other settings to modify the inherited style settings.
;                  $bAutoUpdate         - [optional] a boolean value. Default is Null. If True, Updates the style when you apply direct formatting to a Frame using this style in your document. The formatting of all Frames using this style is automatically updated.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, hide the style in the UI. (Libre 4.0 and up only.)
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 4 Return 0 = $sNewFrameStyleName not a String.
;                  @Error 1 @Extended 5 Return 0 = A Frame style already exists in document by the name called in $sNewFrameStyleName .
;                  @Error 1 @Extended 6 Return 0 = $sParentStyle not a String.
;                  @Error 1 @Extended 7 Return 0 = Frame Style called in $sParentStyle doesn't exist in this Document.
;                  @Error 1 @Extended 8 Return 0 = $bAutoUpdate not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bHidden not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sNewFrameStyleName
;                  |                               2 = Error setting $sParentStyle
;                  |                               4 = Error setting $bAutoUpdate
;                  |                               8 = Error setting $bHidden
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 or 4 Element Array with values in order of function parameters. If the Libre Office version is below 4.0, the Array will contain 3 elements because $bHidden is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LOWriter_FrameStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleOrganizer(ByRef $oDoc, $oFrameStyle, $sNewFrameStyleName = Null, $sParentStyle = Null, $bAutoUpdate = Null, $bHidden = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOrganizer[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($sNewFrameStyleName, $sParentStyle, $bAutoUpdate, $bHidden) Then
		If __LO_VersionCheck(4.0) Then
			__LO_ArrayFill($avOrganizer, $oFrameStyle.Name(), $oFrameStyle.ParentStyle(), $oFrameStyle.IsAutoUpdate(), $oFrameStyle.Hidden())

		Else
			__LO_ArrayFill($avOrganizer, $oFrameStyle.Name(), $oFrameStyle.ParentStyle(), $oFrameStyle.IsAutoUpdate())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOrganizer)
	EndIf

	If ($sNewFrameStyleName <> Null) Then
		If Not IsString($sNewFrameStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If _LOWriter_FrameStyleExists($oDoc, $sNewFrameStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrameStyle.Name = $sNewFrameStyleName
		$iError = ($oFrameStyle.Name() = $sNewFrameStyleName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sParentStyle <> Null) Then
		If Not IsString($sParentStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		If ($sParentStyle <> "") Then
			If Not _LOWriter_FrameStyleExists($oDoc, $sParentStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		EndIf
		$oFrameStyle.ParentStyle = $sParentStyle
		$iError = ($oFrameStyle.ParentStyle() = $sParentStyle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bAutoUpdate <> Null) Then
		If Not IsBool($bAutoUpdate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFrameStyle.IsAutoUpdate = $bAutoUpdate
		$iError = ($oFrameStyle.IsAutoUpdate() = $bAutoUpdate) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not __LO_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oFrameStyle.Hidden = $bHidden
		$iError = ($oFrameStyle.Hidden() = $bHidden) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleOrganizer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStylesGetNames
; Description ...: Retrieve an array of all Frame Style names available for a document.
; Syntax ........: _LOWriter_FrameStylesGetNames(ByRef $oDoc[, $bUserOnly = False[, $bAppliedOnly = False[, $bDisplayName = False]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True only User-Created Frame Styles are returned.
;                  $bAppliedOnly        - [optional] a boolean value. Default is False. If True only Applied Frame Styles are returned.
;                  $bDisplayName        - [optional] a boolean value. Default is False. If True, the style name displayed in the UI (Display Name), instead of the programmatic style name, is returned. See remarks.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bUserOnly not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bAppliedOnly not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bDisplayName not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Frame Style names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. An Array containing all Frame Styles matching the called parameters. @Extended contains the count of results returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If only a Document object is called, all available Frame styles will be returned.
;                  If Both $bUserOnly and $bAppliedOnly are called with True, only User-Created styles that are applied are returned.
;                  Calling $bDisplayName with True will return a list of Style names, as the user sees them in the UI, in the same order as they are returned if $bDisplayName is False. It is best not to use these when setting Styling.
; Related .......: _LOWriter_FrameStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStylesGetNames(ByRef $oDoc, $bUserOnly = False, $bAppliedOnly = False, $bDisplayName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asStyles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDisplayName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$asStyles = __LO_StylesGetNames($oDoc, "FrameStyles", $bUserOnly, $bAppliedOnly, $bDisplayName)
	If Not IsArray($asStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asStyles), $asStyles)
EndFunc   ;==>_LOWriter_FrameStylesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleShadow
; Description ...: Set or Retrieve the shadow settings for a Frame Style.
; Syntax ........: _LOWriter_FrameStyleShadow(ByRef $oFrameStyle[, $iWidth = Null[, $iColor = Null[, $iLocation = Null]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Width of the Frame Shadow set in Hundredths of a Millimeter (HMM).
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Color of the Frame shadow, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Frame Shadow, must be one of the Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer, or less than 0.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
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
;                  LibreOffice may change the shadow width +/- a Hundredth of a Millimeter.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleShadow(ByRef $oFrameStyle, $iWidth = Null, $iColor = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tShdwFrmt
	Local $iError = 0
	Local $avShadow[3]

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tShdwFrmt = $oFrameStyle.ShadowFormat()
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
		If Not __LO_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LOW_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oFrameStyle.ShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oFrameStyle.ShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($tShdwFrmt.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iColor)) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iLocation)) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 4)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleTypePosition
; Description ...: Set or Retrieve Frame Style Position Settings.
; Syntax ........: _LOWriter_FrameStyleTypePosition(ByRef $oFrameStyle[, $iHorAlign = Null[, $iHorPos = Null[, $iHorRelation = Null[, $bMirror = Null[, $iVertAlign = Null[, $iVertPos = Null[, $iVertRelation = Null[, $bKeepInside = Null[, $iAnchorPos = Null]]]]]]]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iHorAlign           - [optional] an integer value (0-3). Default is Null. The horizontal orientation of the Frame. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3. Can't be set if Anchor position is set to "As Character".
;                  $iHorPos             - [optional] an integer value. Default is Null. The horizontal position of the Frame. set in Hundredths of a Millimeter (HMM). Only valid if $iHorAlign is set to $LOW_ORIENT_HORI_NONE().
;                  $iHorRelation        - [optional] an integer value (0-8). Default is Null. The reference point for the selected horizontal alignment option. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3., and Remarks for acceptable values.
;                  $bMirror             - [optional] a boolean value. Default is Null. If True, Reverses the current horizontal alignment settings on even pages.
;                  $iVertAlign          - [optional] an integer value (0-9). Default is Null. The vertical orientation of the Frame. See Constants, $LOW_ORIENT_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertPos            - [optional] an integer value. Default is Null. The vertical position of the Frame. set in Hundredths of a Millimeter (HMM). Only valid if $iVertAlign is set to $LOW_ORIENT_VERT_NONE().
;                  $iVertRelation       - [optional] an integer value (-1-9). Default is Null. The reference point for the selected vertical alignment option. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3., and Remarks for acceptable values.
;                  $bKeepInside         - [optional] a boolean value. Default is Null. If True, Keeps the frame within the layout boundaries of the text that the frame is anchored to.
;                  $iAnchorPos          - [optional] an integer value (0-2,4). Default is Null. Specify the anchoring options for the frame style. See Constants, $LOW_ANCHOR_AT_* as defined in LibreOfficeWriter_Constants.au3..
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iHorAlign Not an Integer, less than 0 or greater than 3. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iHorPos not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iHorRelation not an Integer, less than 0 or greater than 8. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3..
;                  @Error 1 @Extended 6 Return 0 = $bMirror not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 9. See Constants, $LOW_ORIENT_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iVertPos not an Integer.
;                  @Error 1 @Extended 9 Return 0 = $iVertRelation Not an Integer, Less than -1 or greater than 9. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3..
;                  @Error 1 @Extended 10 Return 0 = $bKeepInside not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $iAnchorPos not an Integer, less than 0 or greater than 4, or equal to 3. See Constants, $LOW_ANCHOR_AT_* as defined in LibreOfficeWriter_Constants.au3..
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iHorAlign
;                  |                               2 = Error setting $iHorPos
;                  |                               4 = Error setting $iHorRelation
;                  |                               8 = Error setting $bMirror
;                  |                               16 = Error setting $iVertAlign
;                  |                               32 = Error setting $iVertPos
;                  |                               64 = Error setting $iVertRelation
;                  |                               128 = Error setting $bKeepInside
;                  |                               256 = Error setting $iAnchorPos
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 9 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $iHorRelation has varying acceptable values, depending on the current Anchor position and also the current $iHorAlign setting.
;                  The Following is a list of acceptable values per anchor position.
;                  # $LOW_ANCHOR_AT_PARAGRAPH(0) Accepts the following $iHorRelation Values:
;                  - $LOW_RELATIVE_PARAGRAPH (0),
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT (1),
;                  - $LOW_RELATIVE_PAGE_LEFT (3)[Same as Left Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PAGE_RIGHT (4)[Same as Right Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PARAGRAPH_LEFT (5),
;                  - $LOW_RELATIVE_PARAGRAPH_RIGHT (6),
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  # $LOW_ANCHOR_AS_CHARACTER(1) Accepts No $iHorRelation Values.
;                  # $LOW_ANCHOR_AT_PAGE(2) Accepts the following $iHorRelation Values:
;                  - $LOW_RELATIVE_PAGE_LEFT (3)[Same as Left Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PAGE_RIGHT (4)[Same as Right Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  # $LOW_ANCHOR_AT_CHARACTER(4) Accepts the following $iHorRelation Values:
;                  - $LOW_RELATIVE_PARAGRAPH (0),
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT (1),
;                  - $LOW_RELATIVE_CHARACTER (2),
;                  - $LOW_RELATIVE_PAGE_LEFT (3)[Same as Left Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PAGE_RIGHT (4)[Same as Right Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PARAGRAPH_LEFT (5),
;                  - $LOW_RELATIVE_PARAGRAPH_RIGHT (6),
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  $iVertRelation has varying acceptable values, depending on the current Anchor position. The Following is a list of acceptable values per anchor position.
;                  # $LOW_ANCHOR_AT_PARAGRAPH(0) Accepts the following $iVertRelation Values:
;                  - $LOW_RELATIVE_PARAGRAPH (0)[The Same as "Margin" in L.O. UI],
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  # $LOW_ANCHOR_AS_CHARACTER(1) Accepts the following $iVertRelation Values:
;                  - $LOW_RELATIVE_ROW(-1),
;                  - $LOW_RELATIVE_PARAGRAPH (0)[The Same as "Baseline" in L.O. UI],
;                  - $LOW_RELATIVE_CHARACTER (2),
;                  # $LOW_ANCHOR_AT_PAGE(2) Accepts the following $iVertRelation Values:
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  # $LOW_ANCHOR_AT_CHARACTER(4) Accepts the following $iVertRelation Values:
;                  - $LOW_RELATIVE_PARAGRAPH (0)[The same as "Margin" in L.O. UI],
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT (1),
;                  - $LOW_RELATIVE_CHARACTER (2),
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  - $LOW_RELATIVE_TEXT_LINE (9)[The same as "Line of Text" in L.O. UI]
;                  The behavior of each Relation constant is described below.
;                  - $LOW_RELATIVE_ROW(-1), This option will position the frame considering the height of the row where the anchor is placed.
;                  - $LOW_RELATIVE_PARAGRAPH (0), [For Horizontal Relation:] the frame is positioned considering the whole width available for the paragraph, including indent spaces.
;                  - $LOW_RELATIVE_PARAGRAPH [For Vertical Relation:] {$LOW_RELATIVE_PARAGRAPH is Also called "Margin" or "Baseline" in L.O. UI], Depending on the anchoring type, the frame is positioned considering the space between the top margin and the character ("To character" anchoring) or bottom edge of the paragraph ("To paragraph" anchoring) where the anchor is placed. Or will position the frame considering the text baseline over which all characters are placed. ("As Character" anchoring.)
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT (1), [For Horizontal Relation:] the frame is positioned considering the whole width available for text in the paragraph, excluding indent spaces.
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT [For Vertical relation:] the frame is positioned considering the height of the paragraph where the anchor is placed.
;                  - $LOW_RELATIVE_CHARACTER (2), [For Horizontal Relation:] the frame is positioned considering the horizontal space used by the character.
;                  - $LOW_RELATIVE_CHARACTER [For Vertical relation:] the frame is positioned considering the vertical space used by the character.
;                  - $LOW_RELATIVE_PAGE_LEFT (3),[For Horizontal Relation:], the frame is positioned considering the space available between the left page border and the left paragraph border. [Same as Left Page Border in L.O. UI]
;                  - $LOW_RELATIVE_PAGE_RIGHT (4),[For Horizontal Relation:], the frame is positioned considering the space available between the Right page border and the right paragraph border. [Same as Right Page Border in L.O. UI]
;                  - $LOW_RELATIVE_PARAGRAPH_LEFT (5),[For Horizontal Relation:] the frame is positioned considering the width of the indent space available to the left of the paragraph.
;                  - $LOW_RELATIVE_PARAGRAPH_RIGHT (6),[For Horizontal Relation:], the frame is positioned considering the width of the indent space available to the right of the paragraph.
;                  - $LOW_RELATIVE_PAGE (7),[For Horizontal Relation:], the frame is positioned considering the whole width of the page, from the left to the right page borders.
;                  - $LOW_RELATIVE_PAGE [For Vertical relation:], the frame is positioned considering the full page height, from top to bottom page borders.
;                  - $LOW_RELATIVE_PAGE_PRINT (8),[For Horizontal Relation:], [Same as Page Text Area in L.O. UI] the frame is positioned considering the whole width available for text in the page, from the left to the right page margins.
;                  - $LOW_RELATIVE_PAGE_PRINT [For Vertical relation:], the frame is positioned considering the full height available for text, from top to bottom margins.
;                  - $LOW_RELATIVE_TEXT_LINE (9),[For Vertical relation:], the frame is positioned considering the height of the line of text where the anchor is placed.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleTypePosition(ByRef $oFrameStyle, $iHorAlign = Null, $iHorPos = Null, $iHorRelation = Null, $bMirror = Null, $iVertAlign = Null, $iVertPos = Null, $iVertRelation = Null, $bKeepInside = Null, $iAnchorPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurrentAnchor
	Local $avPosition[9]

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iHorAlign, $iHorPos, $iHorRelation, $bMirror, $iVertAlign, $iVertPos, $iVertRelation, $bKeepInside, $iAnchorPos) Then
		__LO_ArrayFill($avPosition, $oFrameStyle.HoriOrient(), $oFrameStyle.HoriOrientPosition(), $oFrameStyle.HoriOrientRelation(), _
				$oFrameStyle.PageToggle(), $oFrameStyle.VertOrient(), $oFrameStyle.VertOrientPosition(), $oFrameStyle.VertOrientRelation(), _
				$oFrameStyle.IsFollowingTextFlow(), $oFrameStyle.AnchorType())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosition)
	EndIf
	; Accepts HoriOrient Left, Right, Center, and "None" = "From Left"
	If ($iHorAlign <> Null) Then ; Cant be set if Anchor is set to "As Char"
		If Not __LO_IntIsBetween($iHorAlign, $LOW_ORIENT_HORI_NONE, $LOW_ORIENT_HORI_LEFT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrameStyle.HoriOrient = $iHorAlign
		$iError = ($oFrameStyle.HoriOrient() = $iHorAlign) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iHorPos <> Null) Then
		If Not IsInt($iHorPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrameStyle.HoriOrientPosition = $iHorPos
		$iError = (__LO_IntIsBetween($oFrameStyle.HoriOrientPosition(), $iHorPos - 1, $iHorPos + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iHorRelation <> Null) Then
		If Not __LO_IntIsBetween($iHorRelation, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PAGE_PRINT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrameStyle.HoriOrientRelation = $iHorRelation
		$iError = ($oFrameStyle.HoriOrientRelation() = $iHorRelation) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bMirror <> Null) Then
		If Not IsBool($bMirror) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrameStyle.PageToggle = $bMirror
		$iError = ($oFrameStyle.PageToggle() = $bMirror) ? ($iError) : (BitOR($iError, 8))
	EndIf

	; Accepts Orient Top, Bottom, Center, and "None" = "From Top"/From Bottom, plus Row and Char.
	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ORIENT_VERT_NONE, $LOW_ORIENT_VERT_LINE_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrameStyle.VertOrient = $iVertAlign
		$iError = ($oFrameStyle.VertOrient() = $iVertAlign) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iVertPos <> Null) Then
		If Not IsInt($iVertPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFrameStyle.VertOrientPosition = $iVertPos
		$iError = (__LO_IntIsBetween($oFrameStyle.VertOrientPosition(), $iVertPos - 1, $iVertPos + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iVertRelation <> Null) Then
		If Not __LO_IntIsBetween($iVertRelation, $LOW_RELATIVE_ROW, $LOW_RELATIVE_TEXT_LINE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$iCurrentAnchor = (($iAnchorPos <> Null) ? $iAnchorPos : $oFrameStyle.AnchorType())

		; Libre Office is a bit complex in this anchor setting; When set to "As Character", there aren't specific setting values
		; for "Baseline, "Character" and "Row", But For Baseline the VertOrientRelation value is 0, or "$LOW_RELATIVE_PARAGRAPH",
		; For "Character", The VertOrientRelation value is still 0, and the "VertOrient" value (In the L.O. UI the furthest left
		; drop down box) is modified, which can be either $LOW_ORIENT_VERT_CHAR_TOP(1), $LOW_ORIENT_VERT_CHAR_CENTER(2),
		; $LOW_ORIENT_VERT_CHAR_BOTTOM(3), depending on the current value of Top, Bottom and Center, or "From Bottom"/
		; "From Top", of "VertOrient". The same is true For "Row", which means when the anchor is set to "As Character", I need
		; to first determine the desired user setting, $LOW_RELATIVE_ROW(-1), $LOW_RELATIVE_PARAGRAPH(0), or
		; $LOW_RELATIVE_CHARACTER(2), and then determine the current "VertOrient" setting, and then manually set the value to the
		; correct setting. Such as Line_Top, Line_Bottom etc.

		If ($iCurrentAnchor = $LOW_ANCHOR_AS_CHARACTER) Then
			If ($iVertRelation = $LOW_RELATIVE_ROW) Then
				Switch $oFrameStyle.VertOrient()
					Case $LOW_ORIENT_VERT_NONE ; None = "From Bottom or From Top in L.O. UI
						$iError = BitOR($iError, 64) ; -- Row not accepted with this VertOrient Setting.

					Case $LOW_ORIENT_VERT_TOP, $LOW_ORIENT_VERT_CHAR_TOP, $LOW_ORIENT_VERT_LINE_TOP
						$oFrameStyle.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrameStyle.VertOrient = $LOW_ORIENT_VERT_LINE_TOP
						$iError = (($oFrameStyle.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrameStyle.VertOrient() = $LOW_ORIENT_VERT_LINE_TOP)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_CENTER, $LOW_ORIENT_VERT_CHAR_CENTER, $LOW_ORIENT_VERT_LINE_CENTER
						$oFrameStyle.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrameStyle.VertOrient = $LOW_ORIENT_VERT_LINE_CENTER
						$iError = (($oFrameStyle.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrameStyle.VertOrient() = $LOW_ORIENT_VERT_LINE_CENTER)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_BOTTOM, $LOW_ORIENT_VERT_CHAR_BOTTOM, $LOW_ORIENT_VERT_LINE_BOTTOM
						$oFrameStyle.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrameStyle.VertOrient = $LOW_ORIENT_VERT_LINE_BOTTOM
						$iError = (($oFrameStyle.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrameStyle.VertOrient() = $LOW_ORIENT_VERT_LINE_BOTTOM)) ? ($iError) : (BitOR($iError, 64))
				EndSwitch

			ElseIf ($iVertRelation = $LOW_RELATIVE_PARAGRAPH) Then ; Paragraph = Baseline setting in L.O. UI
				$oFrameStyle.VertOrientRelation = $iVertRelation ; Paragraph = Baseline in this case
				$iError = (($oFrameStyle.VertOrientRelation() = $iVertRelation)) ? ($iError) : (BitOR($iError, 64))

			ElseIf ($iVertRelation = $LOW_RELATIVE_CHARACTER) Then
				Switch $oFrameStyle.VertOrient()
					Case $LOW_ORIENT_VERT_NONE ; None = "From Bottom or From Top in L.O. UI
						$iError = BitOR($iError, 64) ; -- Character not accepted with this VertOrient Setting.

					Case $LOW_ORIENT_VERT_TOP, $LOW_ORIENT_VERT_CHAR_TOP, $LOW_ORIENT_VERT_LINE_TOP
						$oFrameStyle.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrameStyle.VertOrient = $LOW_ORIENT_VERT_CHAR_TOP
						$iError = (($oFrameStyle.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrameStyle.VertOrient() = $LOW_ORIENT_VERT_CHAR_TOP)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_CENTER, $LOW_ORIENT_VERT_CHAR_CENTER, $LOW_ORIENT_VERT_LINE_CENTER
						$oFrameStyle.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrameStyle.VertOrient = $LOW_ORIENT_VERT_CHAR_CENTER
						$iError = (($oFrameStyle.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrameStyle.VertOrient() = $LOW_ORIENT_VERT_CHAR_CENTER)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_BOTTOM, $LOW_ORIENT_VERT_CHAR_BOTTOM, $LOW_ORIENT_VERT_LINE_BOTTOM
						$oFrameStyle.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrameStyle.VertOrient = $LOW_ORIENT_VERT_CHAR_BOTTOM
						$iError = (($oFrameStyle.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrameStyle.VertOrient() = $LOW_ORIENT_VERT_CHAR_BOTTOM)) ? ($iError) : (BitOR($iError, 64))
				EndSwitch
			EndIf

		Else
			$oFrameStyle.VertOrientRelation = $iVertRelation
			$iError = ($oFrameStyle.VertOrientRelation() = $iVertRelation) ? ($iError) : (BitOR($iError, 64))
		EndIf
	EndIf

	If ($bKeepInside <> Null) Then
		If Not IsBool($bKeepInside) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oFrameStyle.IsFollowingTextFlow = $bKeepInside
		$iError = ($oFrameStyle.IsFollowingTextFlow() = $bKeepInside) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iAnchorPos <> Null) Then
		If Not __LO_IntIsBetween($iAnchorPos, $LOW_ANCHOR_AT_PARAGRAPH, $LOW_ANCHOR_AT_CHARACTER, $LOW_ANCHOR_AT_FRAME) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oFrameStyle.AnchorType = $iAnchorPos
		$iError = ($oFrameStyle.AnchorType() = $iAnchorPos) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleTypePosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleTypeSize
; Description ...: Set or Retrieve Frame Style Size related settings.
; Syntax ........: _LOWriter_FrameStyleTypeSize(ByRef $oDoc, ByRef $oFrameStyle[, $iWidth = Null[, $iRelativeWidth = Null[, $iWidthRelativeTo = Null[, $bAutoWidth = Null[, $iHeight = Null[, $iRelativeHeight = Null[, $iHeightRelativeTo = Null[, $bAutoHeight = Null[, $bKeepRatio = Null]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Frame, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iRelativeWidth      - [optional] an integer value (0-254). Default is Null. Calculates the width of the frame as a percentage of the width of the page text area. 0 = (off).
;                  $iWidthRelativeTo    - [optional] an integer value (0,7). Default is Null. Determines what 100% width means: either text area (excluding margins) or the entire page (including margins). See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3. Libre Office 4.3 and Up.
;                  $bAutoWidth          - [optional] a boolean value. Default is Null. Automatically adjusts the width of a frame to match the contents of the frame. $iWidth becomes the minimum width the frame must be.
;                  $iHeight             - [optional] an integer value. Default is Null. The height that you want for the Frame, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iRelativeHeight     - [optional] an integer value (0-254). Default is Null. Calculates the Height of the frame as a percentage of the Height of the page text area. 0 = (off).
;                  $iHeightRelativeTo   - [optional] an integer value (0,7). Default is Null. Determines what 100% Height means: either text area (excluding margins) or the entire page (including margins). See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3. Libre Office 4.3 and Up.
;                  $bAutoHeight         - [optional] a boolean value. Default is Null. Automatically adjusts the height of a frame to match the contents of the frame. $iHeight becomes the minimum height the frame must be.
;                  $bKeepRatio          - [optional] a boolean value. Default is Null. Maintains the height and width ratio when you change the width or the height setting.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 4 Return 0 = $iWidth Not an Integer, or less than 51.
;                  @Error 1 @Extended 5 Return 0 = $iRelativeWidth not an Integer, less than 0 or greater than 254.
;                  @Error 1 @Extended 6 Return 0 = $iWidthRelativeTo not an Integer, not equal to 0 and not equal to 7. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bAutoWidth not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iHeight Not an Integer, or less than 51.
;                  @Error 1 @Extended 9 Return 0 = $iRelativeHeight not an Integer, less than 0 or greater than 254.
;                  @Error 1 @Extended 10 Return 0 = $iHeightRelativeTo not an Integer, not equal to 0 and not equal to 7. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $bAutoHeight not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $bKeepRatio not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iRelativeWidth
;                  |                               4 = Error setting $iWidthRelativeTo
;                  |                               8 = Error setting $bAutoWidth
;                  |                               16 = Error setting $iHeight
;                  |                               32 = Error setting $iRelativeHeight
;                  |                               64 = Error setting $iHeightRelativeTo
;                  |                               128 = Error setting $bAutoHeight
;                  |                               256 = Error setting $bKeepRatio
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.3.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 or 9 Element Array depending on current Libre Office Version, If the current Libre Office version is greater or equal to than 4.3, then a 9 element Array is returned, else 7 element array with both $iWidthRelativeTo and $iHeightRelativeTo skipped. Array Element values will be in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function can successfully set "Keep Ratio" however when the user changes this setting in the UI, for some reason the applicable setting values are not updated, so this function may return incorrect values for "Keep Ratio".
;                  When Keep Ratio is set to True, setting Width/Height values via this function will not be kept in ratio.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleTypeSize(ByRef $oDoc, ByRef $oFrameStyle, $iWidth = Null, $iRelativeWidth = Null, $iWidthRelativeTo = Null, $bAutoWidth = Null, $iHeight = Null, $iRelativeHeight = Null, $iHeightRelativeTo = Null, $bAutoHeight = Null, $bKeepRatio = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSize[7]
	Local Const $iCONST_AutoHW_OFF = 1, $iCONST_AutoHW_ON = 2

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iWidth, $iRelativeWidth, $iWidthRelativeTo, $bAutoWidth, $iHeight, $iRelativeHeight, $iHeightRelativeTo, $bAutoHeight, $bKeepRatio) Then
		If __LO_VersionCheck(4.3) Then
			__LO_ArrayFill($avSize, $oFrameStyle.Width(), $oFrameStyle.RelativeWidth(), $oFrameStyle.RelativeWidthRelation(), _
					($oFrameStyle.WidthType() = $iCONST_AutoHW_ON) ? (True) : (False), $oFrameStyle.Height(), $oFrameStyle.RelativeHeight(), _
					$oFrameStyle.RelativeHeightRelation(), ($oFrameStyle.SizeType() = $iCONST_AutoHW_ON) ? (True) : (False), _
					(($oFrameStyle.IsSyncHeightToWidth() And $oFrameStyle.IsSyncWidthToHeight()) ? (True) : (False)))

		Else
			__LO_ArrayFill($avSize, $oFrameStyle.Width(), $oFrameStyle.RelativeWidth(), _
					($oFrameStyle.WidthType() = $iCONST_AutoHW_ON) ? (True) : (False), $oFrameStyle.Height(), _
					$oFrameStyle.RelativeHeight(), ($oFrameStyle.SizeType() = $iCONST_AutoHW_ON) ? (True) : (False), _
					(($oFrameStyle.IsSyncHeightToWidth() And $oFrameStyle.IsSyncWidthToHeight()) ? (True) : (False)))
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSize)
	EndIf

	If ($iWidth <> Null) Then ; Min 51
		If Not __LO_IntIsBetween($iWidth, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrameStyle.Width = $iWidth
		$iError = (__LO_IntIsBetween($oFrameStyle.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRelativeWidth <> Null) Then
		If Not __LO_IntIsBetween($iRelativeWidth, 0, 254) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrameStyle.RelativeWidth = $iRelativeWidth
		$iError = ($oFrameStyle.RelativeWidth() = $iRelativeWidth) ? ($iError) : (BitOR($iError, 2))

		If ($iRelativeWidth <> 0) And ($bAutoWidth <> True) Then ; If AutoWidth is not On, and Relative Width isn't being turned off, then set Width Value.
			If ($oFrameStyle.WidthType() = $iCONST_AutoHW_OFF) Or ($bAutoWidth = False) Then __LOWriter_ObjRelativeSize($oDoc, $oFrameStyle, True)
		EndIf
	EndIf

	If ($iWidthRelativeTo <> Null) Then
		If Not __LO_IntIsBetween($iWidthRelativeTo, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PARAGRAPH, "", $LOW_RELATIVE_PAGE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not __LO_VersionCheck(4.3) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oFrameStyle.RelativeWidthRelation = $iWidthRelativeTo
		$iError = ($oFrameStyle.RelativeWidthRelation() = $iWidthRelativeTo) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bAutoWidth <> Null) Then
		If Not IsBool($bAutoWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrameStyle.WidthType = ($bAutoWidth) ? ($iCONST_AutoHW_ON) : ($iCONST_AutoHW_OFF)
		$iError = ($oFrameStyle.WidthType() = (($bAutoWidth) ? $iCONST_AutoHW_ON : $iCONST_AutoHW_OFF)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iHeight <> Null) Then
		If Not __LO_IntIsBetween($iHeight, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFrameStyle.Height = $iHeight
		$iError = ($oFrameStyle.Height() = $iHeight) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iRelativeHeight <> Null) Then
		If Not __LO_IntIsBetween($iRelativeHeight, 0, 254) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oFrameStyle.RelativeHeight = $iRelativeHeight
		$iError = ($oFrameStyle.RelativeHeight() = $iRelativeHeight) ? ($iError) : (BitOR($iError, 32))

		If ($iRelativeHeight <> 0) And ($bAutoHeight <> True) Then ; If AutoHeight is not On, and Relative Height isn't being turned off, then set Height Value.
			If ($oFrameStyle.SizeType() = $iCONST_AutoHW_OFF) Or ($bAutoHeight = False) Then __LOWriter_ObjRelativeSize($oDoc, $oFrameStyle, False, True)
		EndIf
	EndIf

	If ($iHeightRelativeTo <> Null) Then
		If Not __LO_IntIsBetween($iHeightRelativeTo, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PARAGRAPH, "", $LOW_RELATIVE_PAGE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		If Not __LO_VersionCheck(4.3) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oFrameStyle.RelativeHeightRelation = $iHeightRelativeTo
		$iError = ($oFrameStyle.RelativeHeightRelation() = $iHeightRelativeTo) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bAutoHeight <> Null) Then
		If Not IsBool($bAutoHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oFrameStyle.SizeType = ($bAutoHeight) ? $iCONST_AutoHW_ON : $iCONST_AutoHW_OFF
		$iError = ($oFrameStyle.SizeType = (($bAutoHeight) ? $iCONST_AutoHW_ON : $iCONST_AutoHW_OFF)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bKeepRatio <> Null) Then
		If Not IsBool($bKeepRatio) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oFrameStyle.IsSyncHeightToWidth = $bKeepRatio
		$oFrameStyle.IsSyncWidthToHeight = $bKeepRatio
		$iError = (($oFrameStyle.IsSyncHeightToWidth() = $bKeepRatio) And ($oFrameStyle.IsSyncWidthToHeight() = $bKeepRatio)) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleTypeSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleWrap
; Description ...: Set or Retrieve Frame Style Wrap and Spacing settings.
; Syntax ........: _LOWriter_FrameStyleWrap(ByRef $oFrameStyle[, $iWrapType = Null[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $iWrapType           - [optional] an integer value (0-5). Default is Null. The way you want text to wrap around the frame. See Constants, $LOW_WRAP_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The amount of space that you want between the Left edge of the frame and the text. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The amount of space that you want between the Right edge of the frame and the text. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The amount of space that you want between the Top edge of the frame and the text. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The amount of space that you want between the Bottom edge of the frame and the text. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $iWrapType not an Integer, less than 0 or greater than 5. See Constants, $LOW_WRAP_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iLeft not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iRight not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iBottom not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Property Set Info Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWrapType
;                  |                               2 = Error setting $iLeft
;                  |                               4 = Error setting $iRight
;                  |                               8 = Error setting $iTop
;                  |                               16 = Error setting $iBottom
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleWrap(ByRef $oFrameStyle, $iWrapType = Null, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPropInfo
	Local $iError = 0
	Local $avWrap[5]

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oPropInfo = $oFrameStyle.getPropertySetInfo()
	If Not IsObj($oPropInfo) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWrapType, $iLeft, $iRight, $iTop, $iBottom) Then
		If $oPropInfo.hasPropertyByName("Surround") Then ; Surround is marked as deprecated, but there is no indication of what version of L.O. this occurred. So Test for its existence.
			__LO_ArrayFill($avWrap, $oFrameStyle.Surround(), $oFrameStyle.LeftMargin(), $oFrameStyle.RightMargin(), $oFrameStyle.TopMargin(), _
					$oFrameStyle.BottomMargin())

		Else
			__LO_ArrayFill($avWrap, $oFrameStyle.TextWrap(), $oFrameStyle.LeftMargin(), $oFrameStyle.RightMargin(), $oFrameStyle.TopMargin(), _
					$oFrameStyle.BottomMargin())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avWrap)
	EndIf

	If ($iWrapType <> Null) Then
		If Not __LO_IntIsBetween($iWrapType, $LOW_WRAP_MODE_NONE, $LOW_WRAP_MODE_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If $oPropInfo.hasPropertyByName("Surround") Then $oFrameStyle.Surround = $iWrapType
		If $oPropInfo.hasPropertyByName("TextWrap") Then $oFrameStyle.TextWrap = $iWrapType
		If $oPropInfo.hasPropertyByName("Surround") Then
			$iError = ($oFrameStyle.Surround() = $iWrapType) ? ($iError) : (BitOR($iError, 1))

		Else
			$iError = ($oFrameStyle.TextWrap() = $iWrapType) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($iLeft <> Null) Then
		If Not IsInt($iLeft) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrameStyle.LeftMargin = $iLeft
		$iError = (__LO_IntIsBetween($oFrameStyle.LeftMargin(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iRight <> Null) Then
		If Not IsInt($iRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrameStyle.RightMargin = $iRight
		$iError = (__LO_IntIsBetween($oFrameStyle.RightMargin(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iTop <> Null) Then
		If Not IsInt($iTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrameStyle.TopMargin = $iTop
		$iError = (__LO_IntIsBetween($oFrameStyle.TopMargin(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iBottom <> Null) Then
		If Not IsInt($iBottom) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrameStyle.BottomMargin = $iBottom
		$iError = (__LO_IntIsBetween($oFrameStyle.BottomMargin(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleWrap

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameStyleWrapOptions
; Description ...: Set or Retrieve Frame Style Wrap Options.
; Syntax ........: _LOWriter_FrameStyleWrapOptions(ByRef $oFrameStyle[, $bFirstPar = Null[, $bInBackground = Null[, $bAllowOverlap = Null]]])
; Parameters ....: $oFrameStyle         - [in/out] an object. A Frame Style object returned by a previous _LOWriter_FrameStyleCreate, or _LOWriter_FrameStyleGetObj function.
;                  $bFirstPar           - [optional] a boolean value. Default is Null. If True, starts a new paragraph below the Frame.
;                  $bInBackground       - [optional] a boolean value. Default is Null. If True, moves the selected Frame to the background. This option is only available with the "Through" wrap type.
;                  $bAllowOverlap       - [optional] a boolean value. Default is Null. If True, the Frame is allowed to overlap another Frame. This option has no effect on wrap "Through" Frames, which can always overlap.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrameStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrameStyle not a Frame Style Object.
;                  @Error 1 @Extended 3 Return 0 = $bFirstPar not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bInBackground not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bAllowOverlap not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bFirstPar
;                  |                               2 = Error setting $bInBackground
;                  |                               4 = Error setting $bAllowOverlap
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function may indicate the settings were set successfully when they haven't been if the appropriate wrap type, anchor type etc. hasn't been set before hand.
; Related .......: _LOWriter_FrameStyleCreate, _LOWriter_FrameStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameStyleWrapOptions(ByRef $oFrameStyle, $bFirstPar = Null, $bInBackground = Null, $bAllowOverlap = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abWrapOptions[3]

	If Not IsObj($oFrameStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFrameStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bFirstPar, $bInBackground, $bAllowOverlap) Then
		__LO_ArrayFill($abWrapOptions, $oFrameStyle.SurroundAnchorOnly(), (($oFrameStyle.Opaque()) ? (False) : (True)), _
				$oFrameStyle.AllowOverlap()) ; Opaque/Background is False when InBackground is checked, so switch Boolean values around.

		Return SetError($__LO_STATUS_SUCCESS, 1, $abWrapOptions)
	EndIf

	If ($bFirstPar <> Null) Then
		If Not IsBool($bFirstPar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrameStyle.SurroundAnchorOnly = $bFirstPar
		$iError = ($oFrameStyle.SurroundAnchorOnly() = $bFirstPar) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInBackground <> Null) Then
		If Not IsBool($bInBackground) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrameStyle.Opaque = (($bInBackground) ? False : True)
		$iError = ($oFrameStyle.Opaque() = (($bInBackground) ? False : True)) ? ($iError) : (BitOR($iError, 2)) ; Opaque/Background is False when InBackground is checked, so switch Boolean values around.
	EndIf

	If ($bAllowOverlap <> Null) Then
		If Not IsBool($bAllowOverlap) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrameStyle.AllowOverlap = $bAllowOverlap
		$iError = ($oFrameStyle.AllowOverlap() = $bAllowOverlap) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameStyleWrapOptions

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameTypePosition
; Description ...: Set or Retrieve Frame Position Settings.
; Syntax ........: _LOWriter_FrameTypePosition(ByRef $oFrame[, $iHorAlign = Null[, $iHorPos = Null[, $iHorRelation = Null[, $bMirror = Null[, $iVertAlign = Null[, $iVertPos = Null[, $iVertRelation = Null[, $bKeepInside = Null[, $iAnchorPos = Null]]]]]]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iHorAlign           - [optional] an integer value (0-3). Default is Null. The horizontal orientation of the Frame. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3. Can't be set if Anchor position is set to "As Character".
;                  $iHorPos             - [optional] an integer value. Default is Null. The horizontal position of the Frame. set in Hundredths of a Millimeter (HMM). Only valid if $iHorAlign is set to $LOW_ORIENT_HORI_NONE().
;                  $iHorRelation        - [optional] an integer value (0-8). Default is Null. The reference point for the selected horizontal alignment option. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3, and Remarks for acceptable values.
;                  $bMirror             - [optional] a boolean value. Default is Null. If True, Reverses the current horizontal alignment settings on even pages.
;                  $iVertAlign          - [optional] an integer value (0-9). Default is Null. The vertical orientation of the Frame. See Constants, $LOW_ORIENT_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertPos            - [optional] an integer value. Default is Null. The vertical position of the Frame. set in Hundredths of a Millimeter (HMM). Only valid if $iVertAlign is set to $LOW_ORIENT_VERT_NONE().
;                  $iVertRelation       - [optional] an integer value (-1-9). Default is Null. The reference point for the selected vertical alignment option. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3, and Remarks for acceptable values.
;                  $bKeepInside         - [optional] a boolean value. Default is Null. If True, Keeps the frame within the layout boundaries of the text that the frame is anchored to.
;                  $iAnchorPos          - [optional] an integer value(0-2,4). Default is Null. Specify the anchoring options for the frame. See Constants, $LOW_ANCHOR_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iHorAlign Not an Integer, less than 0 or greater than 3. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iHorPos not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iHorRelation not an Integer, less than 0 or greater than 8. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bMirror not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 9. See Constants, $LOW_ORIENT_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iVertPos not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iVertRelation Not an Integer, Less than -1 or greater than 9. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $bKeepInside not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $iAnchorPos not an Integer, less than 0 or greater than 4, or equal to 3. See Constants, $LOW_ANCHOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iHorAlign
;                  |                               2 = Error setting $iHorPos
;                  |                               4 = Error setting $iHorRelation
;                  |                               8 = Error setting $bMirror
;                  |                               16 = Error setting $iVertAlign
;                  |                               32 = Error setting $iVertPos
;                  |                               64 = Error setting $iVertRelation
;                  |                               128 = Error setting $bKeepInside
;                  |                               256 = Error setting $iAnchorPos
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 9 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $iHorRelation has varying acceptable values, depending on the current Anchor position and also the current $iHorAlign setting.
;                  The Following is a list of acceptable values per anchor position.
;                  # $LOW_ANCHOR_AT_PARAGRAPH(0) Accepts the following $iHorRelation Values:
;                  - $LOW_RELATIVE_PARAGRAPH (0),
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT (1),
;                  - $LOW_RELATIVE_PAGE_LEFT (3)[Same as Left Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PAGE_RIGHT (4)[Same as Right Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PARAGRAPH_LEFT (5),
;                  - $LOW_RELATIVE_PARAGRAPH_RIGHT (6),
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  # $LOW_ANCHOR_AS_CHARACTER(1) Accepts No $iHorRelation Values.
;                  # $LOW_ANCHOR_AT_PAGE(2) Accepts the following $iHorRelation Values:
;                  - $LOW_RELATIVE_PAGE_LEFT (3)[Same as Left Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PAGE_RIGHT (4)[Same as Right Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  # $LOW_ANCHOR_AT_CHARACTER(4) Accepts the following $iHorRelation Values:
;                  - $LOW_RELATIVE_PARAGRAPH (0),
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT (1),
;                  - $LOW_RELATIVE_CHARACTER (2),
;                  - $LOW_RELATIVE_PAGE_LEFT (3)[Same as Left Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PAGE_RIGHT (4)[Same as Right Page Border in L.O. UI],
;                  - $LOW_RELATIVE_PARAGRAPH_LEFT (5),
;                  - $LOW_RELATIVE_PARAGRAPH_RIGHT (6),
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  $iVertRelation has varying acceptable values, depending on the current Anchor position. The Following is a list of acceptable values per anchor position.
;                  # $LOW_ANCHOR_AT_PARAGRAPH(0) Accepts the following $iVertRelation Values:
;                  - $LOW_RELATIVE_PARAGRAPH (0)[The Same as "Margin" in L.O. UI],
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  # $LOW_ANCHOR_AS_CHARACTER(1) Accepts the following $iVertRelation Values:
;                  - $LOW_RELATIVE_ROW(-1),
;                  - $LOW_RELATIVE_PARAGRAPH (0)[The Same as "Baseline" in L.O. UI],
;                  - $LOW_RELATIVE_CHARACTER (2),
;                  # $LOW_ANCHOR_AT_PAGE(2) Accepts the following $iVertRelation Values:
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  # $LOW_ANCHOR_AT_CHARACTER(4) Accepts the following $iVertRelation Values:
;                  - $LOW_RELATIVE_PARAGRAPH (0)[The same as "Margin" in L.O. UI],
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT (1),
;                  - $LOW_RELATIVE_CHARACTER (2),
;                  - $LOW_RELATIVE_PAGE (7),
;                  - $LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;                  - $LOW_RELATIVE_TEXT_LINE (9)[The same as "Line of Text" in L.O. UI]
;                  The behavior of each Relation constant is described below.
;                  - $LOW_RELATIVE_ROW(-1), This option will position the frame considering the height of the row where the anchor is placed.
;                  - $LOW_RELATIVE_PARAGRAPH (0), [For Horizontal Relation:] the frame is positioned considering the whole width available for the paragraph, including indent spaces.
;                  - $LOW_RELATIVE_PARAGRAPH [For Vertical Relation:] {$LOW_RELATIVE_PARAGRAPH is Also called "Margin" or "Baseline" in L.O. UI], Depending on the anchoring type, the frame is positioned considering the space between the top margin and the character ("To character" anchoring) or bottom edge of the paragraph ("To paragraph" anchoring) where the anchor is placed. Or will position the frame considering the text baseline over which all characters are placed. ("As Character" anchoring.)
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT (1), [For Horizontal Relation:] the frame is positioned considering the whole width available for text in the paragraph, excluding indent spaces.
;                  - $LOW_RELATIVE_PARAGRAPH_TEXT [For Vertical relation:] the frame is positioned considering the height of the paragraph where the anchor is placed.
;                  - $LOW_RELATIVE_CHARACTER (2), [For Horizontal Relation:] the frame is positioned considering the horizontal space used by the character.
;                  - $LOW_RELATIVE_CHARACTER [For Vertical relation:] the frame is positioned considering the vertical space used by the character.
;                  - $LOW_RELATIVE_PAGE_LEFT (3),[For Horizontal Relation:], the frame is positioned considering the space available between the left page border and the left paragraph border. [Same as Left Page Border in L.O. UI]
;                  - $LOW_RELATIVE_PAGE_RIGHT (4),[For Horizontal Relation:], the frame is positioned considering the space available between the Right page border and the right paragraph border. [Same as Right Page Border in L.O. UI]
;                  - $LOW_RELATIVE_PARAGRAPH_LEFT (5),[For Horizontal Relation:] the frame is positioned considering the width of the indent space available to the left of the paragraph.
;                  - $LOW_RELATIVE_PARAGRAPH_RIGHT (6),[For Horizontal Relation:], the frame is positioned considering the width of the indent space available to the right of the paragraph.
;                  - $LOW_RELATIVE_PAGE (7),[For Horizontal Relation:], the frame is positioned considering the whole width of the page, from the left to the right page borders.
;                  - $LOW_RELATIVE_PAGE [For Vertical relation:], the frame is positioned considering the full page height, from top to bottom page borders.
;                  - $LOW_RELATIVE_PAGE_PRINT (8),[For Horizontal Relation:], [Same as Page Text Area in L.O. UI] the frame is positioned considering the whole width available for text in the page, from the left to the right page margins.
;                  - $LOW_RELATIVE_PAGE_PRINT [For Vertical relation:], the frame is positioned considering the full height available for text, from top to bottom margins.
;                  - $LOW_RELATIVE_TEXT_LINE (9),[For Vertical relation:], the frame is positioned considering the height of the line of text where the anchor is placed.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByCursor, _LOWriter_FrameGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameTypePosition(ByRef $oFrame, $iHorAlign = Null, $iHorPos = Null, $iHorRelation = Null, $bMirror = Null, $iVertAlign = Null, $iVertPos = Null, $iVertRelation = Null, $bKeepInside = Null, $iAnchorPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurrentAnchor
	Local $avPosition[9]

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iHorAlign, $iHorPos, $iHorRelation, $bMirror, $iVertAlign, $iVertPos, $iVertRelation, $bKeepInside, $iAnchorPos) Then
		__LO_ArrayFill($avPosition, $oFrame.HoriOrient(), $oFrame.HoriOrientPosition(), $oFrame.HoriOrientRelation(), _
				$oFrame.PageToggle(), $oFrame.VertOrient(), $oFrame.VertOrientPosition(), $oFrame.VertOrientRelation(), _
				$oFrame.IsFollowingTextFlow(), $oFrame.AnchorType())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosition)
	EndIf
	; Accepts HoriOrient Left,Right, Center, and "None" = "From Left"
	If ($iHorAlign <> Null) Then ; Cant be set if Anchor is set to "As Char"
		If Not __LO_IntIsBetween($iHorAlign, $LOW_ORIENT_HORI_NONE, $LOW_ORIENT_HORI_LEFT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oFrame.HoriOrient = $iHorAlign
		$iError = ($oFrame.HoriOrient() = $iHorAlign) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iHorPos <> Null) Then
		If Not IsInt($iHorPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrame.HoriOrientPosition = $iHorPos
		$iError = (__LO_IntIsBetween($oFrame.HoriOrientPosition(), $iHorPos - 1, $iHorPos + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iHorRelation <> Null) Then
		If Not __LO_IntIsBetween($iHorRelation, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PAGE_PRINT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrame.HoriOrientRelation = $iHorRelation
		$iError = ($oFrame.HoriOrientRelation() = $iHorRelation) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bMirror <> Null) Then
		If Not IsBool($bMirror) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrame.PageToggle = $bMirror
		$iError = ($oFrame.PageToggle() = $bMirror) ? ($iError) : (BitOR($iError, 8))
	EndIf

	; Accepts Orient Top,Bottom, Center, and "None" = "From Top"/From Bottom, plus Row and Char.
	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ORIENT_VERT_NONE, $LOW_ORIENT_VERT_LINE_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrame.VertOrient = $iVertAlign
		$iError = ($oFrame.VertOrient() = $iVertAlign) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iVertPos <> Null) Then
		If Not IsInt($iVertPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrame.VertOrientPosition = $iVertPos
		$iError = (__LO_IntIsBetween($oFrame.VertOrientPosition(), $iVertPos - 1, $iVertPos + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iVertRelation <> Null) Then
		If Not __LO_IntIsBetween($iVertRelation, $LOW_RELATIVE_ROW, $LOW_RELATIVE_TEXT_LINE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$iCurrentAnchor = (($iAnchorPos <> Null) ? $iAnchorPos : $oFrame.AnchorType())

		; Libre Office is a bit complex in this anchor setting; When set to "As Character", there aren't specific setting
		;		values for "Baseline, "Character" and "Row", But For Baseline the VertOrientRelation value is 0, or
		; "$LOW_RELATIVE_PARAGRAPH", For "Character", The VertOrientRelation value is still 0, and the "VertOrient" value (In the
		; L.O. UI the furthest left drop down box)  is modified, which can be either $LOW_ORIENT_VERT_CHAR_TOP(1),
		; $LOW_ORIENT_VERT_CHAR_CENTER(2), $LOW_ORIENT_VERT_CHAR_BOTTOM(3), depending on the current value of Top, Bottom and
		; Center, or "From Bottom"/ "From Top", of "VertOrient". The same is true For "Row", which means when the anchor is set
		; to "As Character", I need to first determine the the desired user setting, $LOW_RELATIVE_ROW(-1),
		; $LOW_RELATIVE_PARAGRAPH(0), or $LOW_RELATIVE_CHARACTER(2), and then determine the current "VertOrient" setting, and
		; then manually set the value to the correct setting. Such as Line_Top, Line_Bottom etc.

		If ($iCurrentAnchor = $LOW_ANCHOR_AS_CHARACTER) Then
			If ($iVertRelation = $LOW_RELATIVE_ROW) Then
				Switch $oFrame.VertOrient()
					Case $LOW_ORIENT_VERT_NONE ; None = "From Bottom or From Top in L.O. UI
						$iError = BitOR($iError, 64) ; -- Row not accepted with this VertOrient Setting.

					Case $LOW_ORIENT_VERT_TOP, $LOW_ORIENT_VERT_CHAR_TOP, $LOW_ORIENT_VERT_LINE_TOP
						$oFrame.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrame.VertOrient = $LOW_ORIENT_VERT_LINE_TOP
						$iError = (($oFrame.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrame.VertOrient() = $LOW_ORIENT_VERT_LINE_TOP)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_CENTER, $LOW_ORIENT_VERT_CHAR_CENTER, $LOW_ORIENT_VERT_LINE_CENTER
						$oFrame.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrame.VertOrient = $LOW_ORIENT_VERT_LINE_CENTER
						$iError = (($oFrame.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrame.VertOrient() = $LOW_ORIENT_VERT_LINE_CENTER)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_BOTTOM, $LOW_ORIENT_VERT_CHAR_BOTTOM, $LOW_ORIENT_VERT_LINE_BOTTOM
						$oFrame.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrame.VertOrient = $LOW_ORIENT_VERT_LINE_BOTTOM
						$iError = (($oFrame.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrame.VertOrient() = $LOW_ORIENT_VERT_LINE_BOTTOM)) ? ($iError) : (BitOR($iError, 64))
				EndSwitch

			ElseIf ($iVertRelation = $LOW_RELATIVE_PARAGRAPH) Then ; Paragraph = Baseline setting in L.O. UI
				$oFrame.VertOrientRelation = $iVertRelation ; Paragraph = Baseline in this case
				$iError = (($oFrame.VertOrientRelation() = $iVertRelation)) ? ($iError) : (BitOR($iError, 64))

			ElseIf ($iVertRelation = $LOW_RELATIVE_CHARACTER) Then
				Switch $oFrame.VertOrient()
					Case $LOW_ORIENT_VERT_NONE ; None = "From Bottom or From Top in L.O. UI
						$iError = BitOR($iError, 64) ; -- Character not accepted with this VertOrient Setting.

					Case $LOW_ORIENT_VERT_TOP, $LOW_ORIENT_VERT_CHAR_TOP, $LOW_ORIENT_VERT_LINE_TOP
						$oFrame.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrame.VertOrient = $LOW_ORIENT_VERT_CHAR_TOP
						$iError = (($oFrame.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrame.VertOrient() = $LOW_ORIENT_VERT_CHAR_TOP)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_CENTER, $LOW_ORIENT_VERT_CHAR_CENTER, $LOW_ORIENT_VERT_LINE_CENTER
						$oFrame.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrame.VertOrient = $LOW_ORIENT_VERT_CHAR_CENTER
						$iError = (($oFrame.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrame.VertOrient() = $LOW_ORIENT_VERT_CHAR_CENTER)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_BOTTOM, $LOW_ORIENT_VERT_CHAR_BOTTOM, $LOW_ORIENT_VERT_LINE_BOTTOM
						$oFrame.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oFrame.VertOrient = $LOW_ORIENT_VERT_CHAR_BOTTOM
						$iError = (($oFrame.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oFrame.VertOrient() = $LOW_ORIENT_VERT_CHAR_BOTTOM)) ? ($iError) : (BitOR($iError, 64))
				EndSwitch
			EndIf

		Else
			$oFrame.VertOrientRelation = $iVertRelation
			$iError = ($oFrame.VertOrientRelation() = $iVertRelation) ? ($iError) : (BitOR($iError, 64))
		EndIf
	EndIf

	If ($bKeepInside <> Null) Then
		If Not IsBool($bKeepInside) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oFrame.IsFollowingTextFlow = $bKeepInside
		$iError = ($oFrame.IsFollowingTextFlow() = $bKeepInside) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iAnchorPos <> Null) Then
		If Not __LO_IntIsBetween($iAnchorPos, $LOW_ANCHOR_AT_PARAGRAPH, $LOW_ANCHOR_AT_CHARACTER, $LOW_ANCHOR_AT_FRAME) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oFrame.AnchorType = $iAnchorPos
		$iError = ($oFrame.AnchorType() = $iAnchorPos) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameTypePosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameTypeSize
; Description ...: Set or Retrieve Frame Size related settings.
; Syntax ........: _LOWriter_FrameTypeSize(ByRef $oDoc, ByRef $oFrame[, $iWidth = Null[, $iRelativeWidth = Null[, $iWidthRelativeTo = Null[, $bAutoWidth = Null[, $iHeight = Null[, $iRelativeHeight = Null[, $iHeightRelativeTo = Null[, $bAutoHeight = Null[, $bKeepRatio = Null]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Frame, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iRelativeWidth      - [optional] an integer value (0-254). Default is Null. Calculates the width of the frame as a percentage of the width of the page text area. Min. 0 (off). Max 254.
;                  $iWidthRelativeTo    - [optional] an integer value (0,7). Default is Null. determines what 100% width means: either text area (excluding margins) or the entire page (including margins). See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3. Libre Office 4.3 and Up.
;                  $bAutoWidth          - [optional] a boolean value. Default is Null. Automatically adjusts the width of a frame to match the contents of the frame. $iWidth becomes the minimum width the frame must be.
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the Frame, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iRelativeHeight     - [optional] an integer value (0-254). Default is Null. Calculates the Height of the frame as a percentage of the Height of the page text area. Min. 0 (off). Max 254.
;                  $iHeightRelativeTo   - [optional] an integer value (0,7). Default is Null. determines what 100% Height means: either text area (excluding margins) or the entire page (including margins). See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3. Libre Office 4.3 and Up.
;                  $bAutoHeight         - [optional] a boolean value. Default is Null. Automatically adjusts the height of a frame to match the contents of the frame. $iHeight becomes the minimum height the frame must be.
;                  $bKeepRatio          - [optional] a boolean value. Default is Null. Maintains the height and width ratio when you change the width or the height setting.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth Not an Integer, or less than 51.
;                  @Error 1 @Extended 4 Return 0 = $iRelativeWidth not an Integer, less than 0 or greater than 254.
;                  @Error 1 @Extended 5 Return 0 = $iWidthRelativeTo not an Integer, not equal to 0 and not equal to 7. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bAutoWidth not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iHeight Not an Integer, or less than 51.
;                  @Error 1 @Extended 8 Return 0 = $iRelativeHeight not an Integer, less than 0 or greater than 254.
;                  @Error 1 @Extended 9 Return 0 = $iHeightRelativeTo not an Integer, not equal to 0 and not equal to 7. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $bAutoHeight not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bKeepRatio not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iRelativeWidth
;                  |                               4 = Error setting $iWidthRelativeTo
;                  |                               8 = Error setting $bAutoWidth
;                  |                               16 = Error setting $iHeight
;                  |                               32 = Error setting $iRelativeHeight
;                  |                               64 = Error setting $iHeightRelativeTo
;                  |                               128 = Error setting $bAutoHeight
;                  |                               256 = Error setting $bKeepRatio
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.3.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 or 9 Element Array depending on current Libre Office Version, If the current Libre Office version is greater or equal to than 4.3, then a 9 element Array is returned, else 7 element array with both $iWidthRelativeTo and $iHeightRelativeTo skipped. Array Element values will be in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function can successfully set "Keep Ratio" however when the user changes this setting in the UI, for some reason the applicable setting values are not updated, so this function may return incorrect values for "Keep Ratio".
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByCursor, _LOWriter_FrameGetObjByName, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameTypeSize(ByRef $oDoc, ByRef $oFrame, $iWidth = Null, $iRelativeWidth = Null, $iWidthRelativeTo = Null, $bAutoWidth = Null, $iHeight = Null, $iRelativeHeight = Null, $iHeightRelativeTo = Null, $bAutoHeight = Null, $bKeepRatio = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSize[7]
	Local Const $iCONST_AutoHW_OFF = 1, $iCONST_AutoHW_ON = 2

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iWidth, $iRelativeWidth, $iWidthRelativeTo, $bAutoWidth, $iHeight, $iRelativeHeight, $iHeightRelativeTo, $bAutoHeight, $bKeepRatio) Then
		If __LO_VersionCheck(4.3) Then
			__LO_ArrayFill($avSize, $oFrame.Width(), $oFrame.RelativeWidth(), $oFrame.RelativeWidthRelation(), _
					($oFrame.WidthType() = $iCONST_AutoHW_ON) ? (True) : (False), $oFrame.Height(), $oFrame.RelativeHeight(), _
					$oFrame.RelativeHeightRelation(), ($oFrame.SizeType() = $iCONST_AutoHW_ON) ? (True) : (False), _
					(($oFrame.IsSyncHeightToWidth() And $oFrame.IsSyncWidthToHeight()) ? (True) : (False)))

		Else
			__LO_ArrayFill($avSize, $oFrame.Width(), $oFrame.RelativeWidth(), _
					($oFrame.WidthType() = $iCONST_AutoHW_ON) ? (True) : (False), $oFrame.Height(), _
					$oFrame.RelativeHeight(), ($oFrame.SizeType() = $iCONST_AutoHW_ON) ? (True) : (False), _
					(($oFrame.IsSyncHeightToWidth() And $oFrame.IsSyncWidthToHeight()) ? (True) : (False)))
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSize)
	EndIf

	If ($iWidth <> Null) Then ; Min 51
		If Not __LO_IntIsBetween($iWidth, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrame.Width = $iWidth
		$iError = (__LO_IntIsBetween($oFrame.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRelativeWidth <> Null) Then
		If Not __LO_IntIsBetween($iRelativeWidth, 0, 254) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrame.RelativeWidth = $iRelativeWidth
		$iError = ($oFrame.RelativeWidth() = $iRelativeWidth) ? ($iError) : (BitOR($iError, 2))

		If ($iRelativeWidth <> 0) And ($bAutoWidth <> True) Then ; If AutoWidth is not On, and Relative Width isn't being turned off, then set Width Value.
			If ($oFrame.WidthType() = $iCONST_AutoHW_OFF) Or ($bAutoWidth = False) Then __LOWriter_ObjRelativeSize($oDoc, $oFrame, True)
		EndIf
	EndIf

	If ($iWidthRelativeTo <> Null) Then
		If Not __LO_IntIsBetween($iWidthRelativeTo, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PARAGRAPH, "", $LOW_RELATIVE_PAGE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not __LO_VersionCheck(4.3) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oFrame.RelativeWidthRelation = $iWidthRelativeTo
		$iError = ($oFrame.RelativeWidthRelation() = $iWidthRelativeTo) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bAutoWidth <> Null) Then
		If Not IsBool($bAutoWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrame.WidthType = ($bAutoWidth) ? $iCONST_AutoHW_ON : $iCONST_AutoHW_OFF
		$iError = ($oFrame.WidthType() = (($bAutoWidth) ? $iCONST_AutoHW_ON : $iCONST_AutoHW_OFF)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iHeight <> Null) Then
		If Not __LO_IntIsBetween($iHeight, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFrame.Height = $iHeight
		$iError = ($oFrame.Height() = $iHeight) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iRelativeHeight <> Null) Then
		If Not __LO_IntIsBetween($iRelativeHeight, 0, 254) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFrame.RelativeHeight = $iRelativeHeight
		$iError = ($oFrame.RelativeHeight() = $iRelativeHeight) ? ($iError) : (BitOR($iError, 32))

		If ($iRelativeHeight <> 0) And ($bAutoHeight <> True) Then ; If AutoHeight is not On, and Relative Height isn't being turned off, then set Height Value.
			If ($oFrame.SizeType() = $iCONST_AutoHW_OFF) Or ($bAutoHeight = False) Then __LOWriter_ObjRelativeSize($oDoc, $oFrame, False, True)
		EndIf
	EndIf

	If ($iHeightRelativeTo <> Null) Then
		If Not __LO_IntIsBetween($iHeightRelativeTo, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PARAGRAPH, "", $LOW_RELATIVE_PAGE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not __LO_VersionCheck(4.3) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oFrame.RelativeHeightRelation = $iHeightRelativeTo
		$iError = ($oFrame.RelativeHeightRelation() = $iHeightRelativeTo) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bAutoHeight <> Null) Then
		If Not IsBool($bAutoHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oFrame.SizeType = ($bAutoHeight) ? $iCONST_AutoHW_ON : $iCONST_AutoHW_OFF
		$iError = ($oFrame.SizeType = (($bAutoHeight) ? $iCONST_AutoHW_ON : $iCONST_AutoHW_OFF)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bKeepRatio <> Null) Then
		If Not IsBool($bKeepRatio) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oFrame.IsSyncHeightToWidth = $bKeepRatio
		$oFrame.IsSyncWidthToHeight = $bKeepRatio
		$iError = (($oFrame.IsSyncHeightToWidth() = $bKeepRatio) And ($oFrame.IsSyncWidthToHeight() = $bKeepRatio)) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameTypeSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameWrap
; Description ...: Set or Retrieve Frame Wrap and Spacing settings.
; Syntax ........: _LOWriter_FrameWrap(ByRef $oFrame[, $iWrapType = Null[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $iWrapType           - [optional] an integer value (0-5). Default is Null. The way you want text to wrap around the frame. See Constants, $LOW_WRAP_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The amount of space that you want between the Left edge of the frame and the text. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The amount of space that you want between the Right edge of the frame and the text. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The amount of space that you want between the Top edge of the frame and the text. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The amount of space that you want between the Bottom edge of the frame and the text. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWrapType not an Integer, less than 0 or greater than 5. See Constants, $LOW_WRAP_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iLeft not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iRight not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Property Set Info Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWrapType
;                  |                               2 = Error setting $iLeft
;                  |                               4 = Error setting $iRight
;                  |                               8 = Error setting $iTop
;                  |                               16 = Error setting $iBottom
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByCursor, _LOWriter_FrameGetObjByName, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameWrap(ByRef $oFrame, $iWrapType = Null, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPropInfo
	Local $iError = 0
	Local $avWrap[5]

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPropInfo = $oFrame.getPropertySetInfo()
	If Not IsObj($oPropInfo) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWrapType, $iLeft, $iRight, $iTop, $iBottom) Then
		If $oPropInfo.hasPropertyByName("Surround") Then ; Surround is marked as deprecated, but there is no indication of what version of L.O. this occurred. So Test for its existence.
			__LO_ArrayFill($avWrap, $oFrame.Surround(), $oFrame.LeftMargin(), $oFrame.RightMargin(), $oFrame.TopMargin(), _
					$oFrame.BottomMargin())

		Else
			__LO_ArrayFill($avWrap, $oFrame.TextWrap(), $oFrame.LeftMargin(), $oFrame.RightMargin(), $oFrame.TopMargin(), _
					$oFrame.BottomMargin())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avWrap)
	EndIf

	If ($iWrapType <> Null) Then
		If Not __LO_IntIsBetween($iWrapType, $LOW_WRAP_MODE_NONE, $LOW_WRAP_MODE_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If $oPropInfo.hasPropertyByName("Surround") Then $oFrame.Surround = $iWrapType
		If $oPropInfo.hasPropertyByName("TextWrap") Then $oFrame.TextWrap = $iWrapType

		If $oPropInfo.hasPropertyByName("Surround") Then
			$iError = ($oFrame.Surround() = $iWrapType) ? ($iError) : (BitOR($iError, 1))

		Else
			$iError = ($oFrame.TextWrap() = $iWrapType) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($iLeft <> Null) Then
		If Not IsInt($iLeft) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrame.LeftMargin = $iLeft
		$iError = (__LO_IntIsBetween($oFrame.LeftMargin(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iRight <> Null) Then
		If Not IsInt($iRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrame.RightMargin = $iRight
		$iError = (__LO_IntIsBetween($oFrame.RightMargin(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iTop <> Null) Then
		If Not IsInt($iTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFrame.TopMargin = $iTop
		$iError = (__LO_IntIsBetween($oFrame.TopMargin(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iBottom <> Null) Then
		If Not IsInt($iBottom) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFrame.BottomMargin = $iBottom
		$iError = (__LO_IntIsBetween($oFrame.BottomMargin(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameWrap

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FrameWrapOptions
; Description ...: Set or Retrieve Frame Wrap Options.
; Syntax ........: _LOWriter_FrameWrapOptions(ByRef $oFrame[, $bFirstPar = Null[, $bInBackground = Null[, $bAllowOverlap = Null]]])
; Parameters ....: $oFrame              - [in/out] an object. A Frame object returned by a previous _LOWriter_FrameCreate, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function.
;                  $bFirstPar           - [optional] a boolean value. Default is Null. If True, starts a new paragraph below the object.
;                  $bInBackground       - [optional] a boolean value. Default is Null. If True, moves the selected object to the background. This option is only available with the "Through" wrap type.
;                  $bAllowOverlap       - [optional] a boolean value. Default is Null. If True, the object is allowed to overlap another object. This option has no effect on wrap "Through" objects, which can always overlap.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFrame not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFirstPar not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bInBackground not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bAllowOverlap not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bFirstPar
;                  |                               2 = Error setting $bInBackground
;                  |                               4 = Error setting $bAllowOverlap
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Contour and Outside only, though shown on the L.O. UI, are not available for frames, as stated in the L.O. Offline help file.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function may indicate the settings were set successfully when they haven't been if the appropriate wrap type, anchor type etc. hasn't been set before hand.
; Related .......: _LOWriter_FrameCreate, _LOWriter_FrameGetObjByCursor, _LOWriter_FrameGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FrameWrapOptions(ByRef $oFrame, $bFirstPar = Null, $bInBackground = Null, $bAllowOverlap = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abWrapOptions[3]

	If Not IsObj($oFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bFirstPar, $bInBackground, $bAllowOverlap) Then
		__LO_ArrayFill($abWrapOptions, $oFrame.SurroundAnchorOnly(), (($oFrame.Opaque()) ? (False) : (True)), $oFrame.AllowOverlap())
		; Opaque/Background is False when InBackground is checked, so switch Boolean values around.

		Return SetError($__LO_STATUS_SUCCESS, 1, $abWrapOptions)
	EndIf

	If ($bFirstPar <> Null) Then
		If Not IsBool($bFirstPar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oFrame.SurroundAnchorOnly = $bFirstPar
		$iError = ($oFrame.SurroundAnchorOnly() = $bFirstPar) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInBackground <> Null) Then
		If Not IsBool($bInBackground) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFrame.Opaque = (($bInBackground) ? False : True)
		$iError = ($oFrame.Opaque() = (($bInBackground) ? False : True)) ? ($iError) : (BitOR($iError, 2)) ; Opaque/Background is False when InBackground is checked, so switch Boolean values around.
	EndIf

	If ($bAllowOverlap <> Null) Then
		If Not IsBool($bAllowOverlap) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFrame.AllowOverlap = $bAllowOverlap
		$iError = ($oFrame.AllowOverlap() = $bAllowOverlap) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FrameWrapOptions
