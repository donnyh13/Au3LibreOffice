#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Deleting, etc. Impress Drawing Shapes.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_DrawShapeAreaColor
; _LOImpress_DrawShapeAreaFillStyle
; _LOImpress_DrawShapeAreaGradient
; _LOImpress_DrawShapeAreaGradientMulticolor
; _LOImpress_DrawShapeAreaTransparency
; _LOImpress_DrawShapeAreaTransparencyGradient
; _LOImpress_DrawShapeAreaTransparencyGradientMulti
; _LOImpress_DrawShapeDelete
; _LOImpress_DrawShapeExists
; _LOImpress_DrawShapeGetType
; _LOImpress_DrawShapeInsert
; _LOImpress_DrawShapeLineArrowStyles
; _LOImpress_DrawShapeLineProperties
; _LOImpress_DrawShapeName
; _LOImpress_DrawShapePointsAdd
; _LOImpress_DrawShapePointsGetCount
; _LOImpress_DrawShapePointsModify
; _LOImpress_DrawShapePointsRemove
; _LOImpress_DrawShapePosition
; _LOImpress_DrawShapeRotateSlant
; _LOImpress_DrawShapeTextBox
; _LOImpress_DrawShapeTextboxCreateTextCursor
; _LOImpress_DrawShapeTypePosition
; _LOImpress_DrawShapeTypeSize
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeAreaColor
; Description ...: Set or Retrieve the Fill color settings for a Shape.
; Syntax ........: _LOImpress_DrawShapeAreaColor(ByRef $oShape[, $iColor = Null])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is Null. The Fill color. Set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColor not an integer, less than -1, or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current Fill color as an integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOImpress_DrawShapeInsert, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeAreaColor(ByRef $oShape, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; If $iColor is Null, and Fill Style is set to solid, then return current color value, else return LO_COLOR_OFF.
	If ($iColor = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, ($oShape.FillStyle() = $LOI_AREA_FILL_STYLE_SOLID) ? (__LOImpress_ColorRemoveAlpha($oShape.FillColor())) : ($LO_COLOR_OFF))

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($iColor = $LO_COLOR_OFF) Then
		$oShape.FillStyle = $LOI_AREA_FILL_STYLE_OFF

	Else
		$iOldTransparency = $oShape.FillTransparence()
		If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oShape.FillStyle = $LOI_AREA_FILL_STYLE_SOLID
		$oShape.FillColor = $iColor
		$iError = ($oShape.FillColor() = $iColor) ? ($iError) : (BitOR($iError, 1))

		$oShape.FillTransparence = $iOldTransparency
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOImpress_DrawShapeAreaFillStyle(ByRef $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Fill Style.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning current background fill style. Return will be one of the constants $LOI_AREA_FILL_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeAreaFillStyle(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oShape.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOImpress_DrawShapeAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeAreaGradient
; Description ...: Modify or retrieve the settings for Shape Background color Gradient.
; Syntax ........: _LOImpress_DrawShapeAreaGradient(ByRef $oShape[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See remarks. See constants, $LOI_GRAD_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient type to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iIncrement          - [optional] an integer value (0, 3-256). Default is Null. The number of steps of color change. 0 = Automatic.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] an integer value (0-16777215). Default is Null. A color for the beginning point of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] an integer value (0-16777215). Default is Null. A color for the endpoint of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sGradientName not a String.
;                  @Error 1 @Extended 3 Return 0 = $iType not an Integer, less than -1, or greater than 5. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error 1 @Extended 5 Return 0 = $iXCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iYCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iAngle not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 8 Return 0 = $iTransitionStart not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iFromColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 10 Return 0 = $iToColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 11 Return 0 = $iFromIntense not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 12 Return 0 = $iToIntense not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving "FillGradient" Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Parent Document Object.
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
; Related .......: _LOImpress_DrawShapeInsert, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeAreaGradient(ByRef $oShape, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc
	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName
	Local $atColorStop[0]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tStyleGradient = $oShape.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oShape.FillGradientName(), $tStyleGradient.Style(), _
				$oShape.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	$oDoc = $oShape.Parent.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If ($oShape.FillStyle() <> $LOI_AREA_FILL_STYLE_GRADIENT) Then $oShape.FillStyle = $LOI_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		__LOImpress_GradientPresets($oDoc, $oShape, $tStyleGradient, $sGradientName)
		$iError = ($oShape.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oShape.FillStyle = $LOI_AREA_FILL_STYLE_OFF
			$oShape.FillGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOI_GRAD_TYPE_LINEAR, $LOI_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oShape.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
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

	If ($iFromColor <> Null) Then
		If Not __LO_IntIsBetween($iFromColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

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
		If Not __LO_IntIsBetween($iToColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

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
		If Not __LO_IntIsBetween($iFromIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tStyleGradient.StartIntensity = $iFromIntense
	EndIf

	If ($iToIntense <> Null) Then
		If Not __LO_IntIsBetween($iToIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$tStyleGradient.EndIntensity = $iToIntense
	EndIf

	If ($oShape.FillGradientName() = "") Then
		$sGradName = __LOImpress_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oShape.FillGradientName = $sGradName
		If ($oShape.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	$oShape.FillGradient = $tStyleGradient

	; Error checking
	$iError = ($iType = Null) ? ($iError) : (($oShape.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oShape.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oShape.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oShape.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oShape.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = ($iFromColor = Null) ? ($iError) : (($oShape.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = ($iToColor = Null) ? ($iError) : (($oShape.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = ($iFromIntense = Null) ? ($iError) : (($oShape.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = ($iToIntense = Null) ? ($iError) : (($oShape.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeAreaGradientMulticolor
; Description ...: Set or Retrieve a Shape's Multicolor Gradient settings. See remarks.
; Syntax ........: _LOImpress_DrawShapeAreaGradientMulticolor(ByRef $oShape[, $avColorStops = Null])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
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
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were set to Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, in Long integer format.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOImpress_GradientMulticolorAdd, _LOImpress_GradientMulticolorDelete, _LOImpress_GradientMulticolorModify, _LOImpress_DrawShapeAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeAreaGradientMulticolor(ByRef $oShape, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oShape.FillGradient()
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
	$oShape.FillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oShape.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeAreaTransparency
; Description ...: Set or retrieve Transparency settings for a Shape.
; Syntax ........: _LOImpress_DrawShapeAreaTransparency(ByRef $oShape[, $iTransparency = Null])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
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
; Related .......: _LOImpress_DrawShapeInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeAreaTransparency(ByRef $oShape, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oShape.FillTransparence())

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oShape.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oShape.FillTransparence = $iTransparency
	$iError = ($oShape.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeAreaTransparencyGradient
; Description ...: Set or retrieve the Shape transparency gradient settings.
; Syntax ........: _LOImpress_DrawShapeAreaTransparencyGradient(ByRef $oShape[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient that you want to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3. Set to $LOI_GRAD_TYPE_OFF to turn Transparency Gradient off.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iStart              - [optional] an integer value (0-100). Default is Null. The transparency value for the beginning point of the gradient, where 0% is fully opaque and 100% is fully transparent.
;                  $iEnd                - [optional] an integer value (0-100). Default is Null. The transparency value for the endpoint of the gradient, where 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than -1, or greater than 5, see constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iXCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 4 Return 0 = $iYCenter not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 5 Return 0 = $iAngle not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 6 Return 0 = $iTransitionStart not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iStart not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iEnd not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving "FillTransparenceGradient" Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Parent Document Object.
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
; Related .......: _LOImpress_DrawShapeInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeAreaTransparencyGradient(ByRef $oShape, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc
	Local $tGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop[0]
	Local $fValue

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tGradient = $oShape.FillTransparenceGradient()
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		__LO_ArrayFill($aiTransparent, $tGradient.Style(), $tGradient.XOffset(), $tGradient.YOffset(), _
				Int($tGradient.Angle() / 10), $tGradient.Border(), __LOImpress_TransparencyGradientConvert(Null, $tGradient.StartColor()), _
				__LOImpress_TransparencyGradientConvert(Null, $tGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	$oDoc = $oShape.Parent.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If ($iType <> Null) Then
		If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oShape.FillTransparenceGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOI_GRAD_TYPE_LINEAR, $LOI_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tGradient.Style = $iType
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tGradient.Border = $iTransitionStart
	EndIf

	If ($iStart <> Null) Then
		If Not __LO_IntIsBetween($iStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tGradient.StartColor = __LOImpress_TransparencyGradientConvert($iStart)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

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
		If Not __LO_IntIsBetween($iEnd, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tGradient.EndColor = __LOImpress_TransparencyGradientConvert($iEnd)

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

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

	If ($oShape.FillTransparenceGradientName() = "") Then
		$sTGradName = __LOImpress_TransparencyGradientNameInsert($oDoc, $tGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oShape.FillTransparenceGradientName = $sTGradName
		If ($oShape.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	$oShape.FillTransparenceGradient = $tGradient

	$iError = ($iType = Null) ? ($iError) : (($oShape.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oShape.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oShape.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oShape.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oShape.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iStart = Null) ? ($iError) : (($oShape.FillTransparenceGradient.StartColor() = __LOImpress_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iEnd = Null) ? ($iError) : (($oShape.FillTransparenceGradient.EndColor() = __LOImpress_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Shape's Multi Transparency Gradient settings. See remarks.
; Syntax ........: _LOImpress_DrawShapeAreaTransparencyGradientMulti(ByRef $oShape[, $avColorStops = Null])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
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
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were set to Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple Transparency stops in a Gradient rather than just a beginning and an ending value, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the Transparency percentage value between 0 and 100%.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOImpress_TransparencyGradientMultiModify, _LOImpress_TransparencyGradientMultiDelete, _LOImpress_TransparencyGradientMultiAdd, _LOImpress_DrawShapeAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeAreaTransparencyGradientMulti(ByRef $oShape, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oShape.FillTransparenceGradient()
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
	$oShape.FillTransparenceGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oShape.FillTransparenceGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeDelete
; Description ...: Delete a Shape.
; Syntax ........: _LOImpress_DrawShapeDelete(ByRef $oShape)
; Parameters ....: $oShape                - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape's containing Slide.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of shapes.
;                  @Error 3 @Extended 3 Return 0 = Same number of shapes still present. Failed to delete the Shape.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Shape was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeDelete(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iShapes
	Local $oDrawPage

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDrawPage = $oShape.Parent()
	If Not IsObj($oDrawPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iShapes = $oDrawPage.getCount()
	If Not IsInt($iShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDrawPage.remove($oShape)

	Return ($oDrawPage.getCount() = $iShapes) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeExists
; Description ...: Check if a Slide contains a Shape with the specified name.
; Syntax ........: _LOImpress_DrawShapeExists(ByRef $oSlide, $sShapeName)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $sShapeName          - a string value. The Shape name to search for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sShapeName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Shape name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If a Shape was found matching $sShapeName, True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: On newly created shapes, the name value is blank for some reason, even though the shape in the UI has a name.
;                  The Shape name must be unique, however due to the above issue, it is possible to have two shapes with the same name in the UI (But not internally).
;                  I have modified the check for a Shape name already existing by assuming a Shape's name when encountering one without a name, e.g. assuming the name would be "Shape 1" etc.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeExists(ByRef $oSlide, $sShapeName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oSlide.hasElements() Then
		For $i = 0 To $oSlide.getCount() - 1
			$sName = $oSlide.getByIndex($i).Name()
			If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			; Impress doesn't set the Shape name on new shapes. It has names in the UI that would correspond to the order of the shapes inserted, i.e. Shape 1, Shape 2. Etc.
			If (($sName = "") And (("Shape " & ($i + 1)) = $sShapeName)) Or ($sName = $sShapeName) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; No matches
EndFunc   ;==>_LOImpress_DrawShapeExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeGetType
; Description ...: Return the Drawing Shape's Type corresponding to the constants $LOI_DRAWSHAPE_TYPE_*
; Syntax ........: _LOImpress_DrawShapeGetType(ByRef $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve CustomShapeGeometry Array.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve CustomShapeGeometry "Type" value.
;                  @Error 3 @Extended 3 Return 0 = Failed to determine CustomShape's type.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify what type of "com.sun.star.drawing.EllipseShape" called shape is.
;                  @Error 3 @Extended 5 Return 0 = Called Shape is a unknown shape type.
;                  --Success--
;                  @Error 0 @Extended 1 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_CONNECTOR_* Type Shape. Returning $LOI_DRAWSHAPE_TYPE_CONNECTOR Constant Value. See Remarks #4.
;                  @Error 0 @Extended 2 Return Integer = Success. Shape is a Custom Shape Type. Returning appropriate Constant for shape type if successfully identified, else -1 if identification failed. See Remarks #1. See Constants, $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 0 @Extended 3 Return Integer = Success. Shape is a*_BASIC_CIRCLE_SEGMENT or *_BASIC_ARC Type Shape. Returning appropriate Constant, See Constants, $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 0 @Extended 4 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_CURVE Shape.
;                  @Error 0 @Extended 5 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED Shape.
;                  @Error 0 @Extended 6 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE Shape.
;                  @Error 0 @Extended 7 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED Shape.
;                  @Error 0 @Extended 8 Return Integer = Success. Shape is a *_LINE_LINE, *_LINE_LINE_45 or $LOI_DRAWSHAPE_TYPE_LINE_ARROW_* Type Shape. Returning $LOI_DRAWSHAPE_TYPE_LINE_LINE Constant Value. See Remarks #4.
;                  @Error 0 @Extended 9 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION Shape.
;                  @Error 0 @Extended 10 Return Integer = Success. Shape is a *_LINE_POLYGON, or *_LINE_POLYGON_45 Type Shape. Returning $LOI_DRAWSHAPE_TYPE_LINE_POLYGON Constant Value. See Remarks #2.
;                  @Error 0 @Extended 11 Return Integer = Success. Shape is a *_LINE_POLYGON_FILLED, or *_LINE_POLYGON_45_FILLED Type Shape. Returning $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED Constant Value. See Remarks #2.
;                  @Error 0 @Extended 11 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_3D_* Type Shape. Returning $LOI_DRAWSHAPE_TYPE_3D_CONE Constant Value. See Remarks #5.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: #1 Some shapes are not implemented, or not fully implemented into LibreOffice for automation, consequently they do not have appropriate type names as of yet. Many have simply ambiguous names, such as "non-primitive".
;                  Because of this the following Custom shape types cannot be identified, and this function will return -1:
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
;                  - $LOI_DRAWSHAPE_TYPE_LINE_POLYGON, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45 (The Value of $LOI_DRAWSHAPE_TYPE_LINE_POLYGON is returned for either of these.)
;                  - $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED (The Value of $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED is returned for either of these.)
;                  #3 The following Shapes have strange names that may change in the future, but currently are able to be identified:
;                  - $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE, known as, "mso-spt21", should be "doorplate"
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, known as, "col-502ad400", should be ??
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, known as, "col-60da8460", should be ??
;                  #4 The following Shapes are customizable one to another, and are consequently indistinguishable:
;                  - $LOI_DRAWSHAPE_TYPE_FONTWORK_* (The Value of $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL is returned for any of these.)
;                  - $LOI_DRAWSHAPE_TYPE_LINE_LINE* or $LOI_DRAWSHAPE_TYPE_LINE_LINE_45 (The Value of $LOI_DRAWSHAPE_TYPE_LINE_LINE is returned for any of these.)
;                  #5 The following Shapes are have nothing unique that I have found yet to identify each, and are consequently indistinguishable:
;                  - $LOI_DRAWSHAPE_TYPE_3D_* (The Value of $LOI_DRAWSHAPE_TYPE_3D_CONE is returned for any of these.)
;                  #6 The following shapes are customizable one to another, they may be identified, or may return a general shape type:
;                  When the arrowhead type "Arrow" is set in the LO UI, or upon creation of a line with arrows, the internal name of the arrowhead is set to an incrementing name of "Arrowheads x", where x is an integer value. Since I have no way to determine if the head is a custom arrowhead or supposed to be the "Arrow" type, I cannot necessarily identify the right connector or Line.
;                  When setting an Arrowhead to be $LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW, the head is set correctly, but the LibreOffice UI will show "None". The return for Arrowhead type will be the correct name however, and will allow me to identify the shape.
;                  - If I fail to identify, or it is customized differently, $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_* Connector, $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT will be returned.
;                  - If I fail to identify, or it is customized differently, $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_* Connector, $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE will be returned.
;                  - If I fail to identify, or it is customized differently, $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_* Connector, $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED will be returned.
;                  - If I fail to identify, or it is customized differently, $LOI_DRAWSHAPE_TYPE_CONNECTOR_* Connector, or if I fail to identify the sub-type of a connector, $LOI_DRAWSHAPE_TYPE_CONNECTOR will be returned.
;                  - If I fail to identify $LOI_DRAWSHAPE_TYPE_LINE_ARROW_*, the Value of $LOI_DRAWSHAPE_TYPE_LINE_LINE is returned for any of these.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeGetType(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atCusShapeGeo[0]
	Local Const $iCircleKind_CUT = 2 ; a circle with a cut connected by a line.
	Local Const $iCircleKind_ARC = 3 ; a circle with an open cut.
	Local $sType
	Local $iReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $oShape.ShapeType()
		Case "com.sun.star.drawing.ConnectorShape"
			Switch $oShape.EdgeKind()
				Case $LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD
					Select
						Case ($oShape.LineEndName = "Arrow") And ($oShape.LineStartName = "")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_ENDS_ARROW)

						Case ($oShape.LineEndName = "Arrow") And ($oShape.LineStartName = "Arrow")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_ARROWS)

						Case ($oShape.LineEndName = "") And ($oShape.LineStartName = "")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR)

						Case Else

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR)
					EndSelect

				Case $LOI_DRAWSHAPE_CONNECTOR_TYPE_CURVE
					Select
						Case ($oShape.LineEndName = "Arrow") And ($oShape.LineStartName = "")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ENDS_ARROW)

						Case ($oShape.LineEndName = "Arrow") And ($oShape.LineStartName = "Arrow")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ARROWS)

						Case ($oShape.LineEndName = "") And ($oShape.LineStartName = "")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED)

						Case Else

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED)
					EndSelect

				Case $LOI_DRAWSHAPE_CONNECTOR_TYPE_STRAIGHT
					Select
						Case ($oShape.LineEndName = "Arrow") And ($oShape.LineStartName = "")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW)

						Case ($oShape.LineEndName = "Arrow") And ($oShape.LineStartName = "Arrow")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ARROWS)

						Case ($oShape.LineEndName = "") And ($oShape.LineStartName = "")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT)

						Case Else

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT)
					EndSelect

				Case $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE
					Select
						Case ($oShape.LineEndName = "Arrow") And ($oShape.LineStartName = "")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ENDS_ARROW)

						Case ($oShape.LineEndName = "Arrow") And ($oShape.LineStartName = "Arrow")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ARROWS)

						Case ($oShape.LineEndName = "") And ($oShape.LineStartName = "")

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE)

						Case Else

							Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE)
					EndSelect

				Case Else ; on error fall back.

					Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR)
			EndSwitch

		Case "com.sun.star.drawing.CustomShape"
			$atCusShapeGeo = $oShape.CustomShapeGeometry()
			If Not IsArray($atCusShapeGeo) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			For $i = 0 To UBound($atCusShapeGeo) - 1
				If ($atCusShapeGeo[$i].Name() = "Type") Then
					$sType = $atCusShapeGeo[$i].Value()
					If Not IsString($sType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

					ExitLoop
				EndIf

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$iReturn = __LOImpress_DrawShape_GetCustomType($sType)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			Return SetError($__LO_STATUS_SUCCESS, 2, $iReturn)

		Case "com.sun.star.drawing.EllipseShape"
			If ($oShape.CircleKind() = $iCircleKind_CUT) Then ; Circle Segment = CircleKind_CUT(2), Arc = CircleKind_ARC(3)

				Return SetError($__LO_STATUS_SUCCESS, 3, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT)

			ElseIf ($oShape.CircleKind() = $iCircleKind_ARC) Then

				Return SetError($__LO_STATUS_SUCCESS, 3, $LOI_DRAWSHAPE_TYPE_BASIC_ARC)

			Else

				Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
			EndIf

		Case "com.sun.star.drawing.OpenBezierShape"

			Return SetError($__LO_STATUS_SUCCESS, 4, $LOI_DRAWSHAPE_TYPE_LINE_CURVE)

		Case "com.sun.star.drawing.ClosedBezierShape"

			Return SetError($__LO_STATUS_SUCCESS, 5, $LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED)

		Case "com.sun.star.drawing.OpenFreeHandShape"

			Return SetError($__LO_STATUS_SUCCESS, 6, $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE)

		Case "com.sun.star.drawing.ClosedFreeHandShape"

			Return SetError($__LO_STATUS_SUCCESS, 7, $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED)

		Case "com.sun.star.drawing.LineShape" ; No way to differentiate between these?? (Lines + Arrows)
			Select
				Case ($oShape.LineStartName = "") And ($oShape.LineEndName = "")

					Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_LINE)

				Case ($oShape.LineStartName = "Arrow") And ($oShape.LineEndName = "")

					Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_STARTS_ARROW)

				Case ($oShape.LineStartName = "Square") And ($oShape.LineEndName = "Arrow")

					Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_SQUARE_ARROW)

				Case ($oShape.LineStartName = "") And ($oShape.LineEndName = "Arrow")

					Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ENDS_ARROW)

				Case ($oShape.LineStartName = "Circle") And ($oShape.LineEndName = "Arrow")

					Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_CIRCLE_ARROW)

				Case ($oShape.LineStartName = "Arrow") And ($oShape.LineEndName = "Square")

					Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_SQUARE)

				Case ($oShape.LineStartName = "Arrow") And ($oShape.LineEndName = "Circle")

					Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_CIRCLE)

				Case ($oShape.LineStartName = "Arrow") And ($oShape.LineEndName = "Arrow")

					Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS)

				Case Else

					Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_LINE) ; On error fall back.
			EndSelect

		Case "com.sun.star.drawing.MeasureShape"

			Return SetError($__LO_STATUS_SUCCESS, 9, $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION)

		Case "com.sun.star.drawing.PolyLineShape"
;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON ; No way to differentiate between these??
;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45

			Return SetError($__LO_STATUS_SUCCESS, 10, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON)

		Case "com.sun.star.drawing.PolyPolygonShape"

			Return SetError($__LO_STATUS_SUCCESS, 11, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED)
;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED ; No way to differentiate between these??
;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED

		Case "com.sun.star.drawing.Shape3DSceneObject" ; No way to differentiate between these??

			Return SetError($__LO_STATUS_SUCCESS, 12, $LOI_DRAWSHAPE_TYPE_3D_CONE)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Unknown shape type.
	EndSwitch
EndFunc   ;==>_LOImpress_DrawShapeGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeInsert
; Description ...: Insert a shape into a slide.
; Syntax ........: _LOImpress_DrawShapeInsert(ByRef $oSlide, $iShapeType, $iWidth, $iHeight)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iShapeType          - an integer value (0-187). The Type of shape to create. See remarks. See $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $iWidth              - an integer value. The Shape's Width in Micrometers. Note, for Lines, Width is the length of the line
;                  $iHeight             - an integer value. The Shape's Height in Micrometers. Note, for Lines, Height is the amount the line goes below the point of insertion.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iShapeType not an Integer, less than 0, or greater than 187. See $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iHeight not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create requested Shape.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Can't insert 3D shapes.
;                  @Error 3 @Extended 2 Return 0 = Can't insert Fontwork shapes.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. The Shape was successfully inserted. Returning the Shape's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Line Shapes, such as Curves etc., may not be smoothly curved. This is due to my lack of understanding of setting Point type settings. You will need to manually select the individual points and set the Point type in L.O. UI.
;                  Polygon and Polygon 45 degree are the same shape internally, one only allows you to draw the lines at 45 degree angles in L.O. UI.
;                  The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
;                  - $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, $LOI_DRAWSHAPE_TYPE_BASIC_FRAME
;                  - $LOI_DRAWSHAPE_TYPE_STARS_6_POINT, $LOI_DRAWSHAPE_TYPE_STARS_12_POINT, $LOI_DRAWSHAPE_TYPE_STARS_SIGNET, $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD, $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER, $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
;                  Inserting any of the above shapes will still show successful, but the shape will be invisible, and could cause the document to crash.
;                  The following shape is visually different from the manually inserted one in L.O. 7.3.4.2:
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING
;                  I presently don't know how to insert 3D shapes or Fontwork, consequently all $LOI_DRAWSHAPE_TYPE_3D_* and $LOI_DRAWSHAPE_TYPE_FONTWORK_* will return a processing error.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeInsert(ByRef $oSlide, $iShapeType, $iWidth, $iHeight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iShapeType, $LOI_DRAWSHAPE_TYPE_3D_CONE, $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_3D_CONE To $LOI_DRAWSHAPE_TYPE_3D_TORUS ; Can't create a 3D shape.

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY To $LOI_DRAWSHAPE_TYPE_ARROWS_PENTAGON ; Create an Arrow Shape.
			$oShape = __LOImpress_DrawShape_CreateArrow($oSlide, $iWidth, $iHeight, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_BASIC_ARC To $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT ; Create a Basic Shape.
			$oShape = __LOImpress_DrawShape_CreateBasic($oSlide, $iWidth, $iHeight, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_CLOUD To $LOI_DRAWSHAPE_TYPE_CALLOUT_ROUND ; Create a Callout Shape.
			$oShape = __LOImpress_DrawShape_CreateCallout($oSlide, $iWidth, $iHeight, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_CONNECTOR To $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW ; Create a Connector.
			$oShape = __LOImpress_DrawShape_CreateLine($oSlide, $iWidth, $iHeight, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_CARD To $LOI_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR ; Create a Flowchart Shape.
			$oShape = __LOImpress_DrawShape_CreateFlowchart($oSlide, $iWidth, $iHeight, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL To $LOI_DRAWSHAPE_TYPE_FONTWORK_TRICOLORE ; Can't create Fontwork.

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS To $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED ; Create a Line Shape.
			$oShape = __LOImpress_DrawShape_CreateLine($oSlide, $iWidth, $iHeight, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_STARS_4_POINT To $LOI_DRAWSHAPE_TYPE_STARS_SIGNET ; Create a Star or Banner Shape.
			$oShape = __LOImpress_DrawShape_CreateStars($oSlide, $iWidth, $iHeight, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND To $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE ; Create a Symbol Shape.
			$oShape = __LOImpress_DrawShape_CreateSymbol($oSlide, $iWidth, $iHeight, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>_LOImpress_DrawShapeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeLineArrowStyles
; Description ...: Set or Retrieve Shape Line Start and End Arrow Style settings.
; Syntax ........: _LOImpress_DrawShapeLineArrowStyles(ByRef $oShape[, $vStartStyle = Null[, $iStartWidth = Null[, $bStartCenter = Null[, $bSync = Null[, $vEndStyle = Null[, $iEndWidth = Null[, $bEndCenter = Null]]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $vStartStyle         - [optional] a variant value (0-32, or String). Default is Null. The Arrow head to apply to the start of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_DRAWSHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iStartWidth         - [optional] an integer value (0-5004). Default is Null. The Width of the Starting Arrowhead, in Micrometers.
;                  $bStartCenter        - [optional] a boolean value. Default is Null. If True, Places the center of the Start arrowhead on the endpoint of the line.
;                  $bSync               - [optional] a boolean value. Default is Null. If True, Synchronizes the Start Arrowhead settings with the end Arrowhead settings. See remarks.
;                  $vEndStyle           - [optional] a variant value (0-32, or String). Default is Null. The Arrow head to apply to the end of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_DRAWSHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iEndWidth           - [optional] an integer value (0-5004). Default is Null. The Width of the Ending Arrowhead, in Micrometers.
;                  $bEndCenter          - [optional] a boolean value. Default is Null. If True, Places the center of the End arrowhead on the endpoint of the line.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $vStartStyle not a String, and not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $vStartStyle is an Integer, but less than 0, or greater than 32. See constants $LOI_DRAWSHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iStartWidth not an Integer, less than 0, or greater than 5004.
;                  @Error 1 @Extended 5 Return 0 = $bStartCenter not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bSync not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $vEndStyle not a String, and not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $vSEndStyle is an Integer, but less than 0, or greater than 32. See constants $LOI_DRAWSHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $iEndWidth not an Integer, less than 0, or greater than 5004.
;                  @Error 1 @Extended 10 Return 0 = $bEndCenter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to convert Constant to Arrowhead name.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function works for connector shapes also.
;                  When the arrowhead type "Arrow" is set in the LO UI, or upon creation of a line with arrows, the internal name of the arrowhead is set to an incrementing name of "Arrowheads x", where x is an integer value. Since I have no way to determine if the head is a custom arrowhead or supposed to be the "Arrow" type, the return when this is present will be the name "Arrowheads x", and not $LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW.
;                  When setting an Arrowhead to be $LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW, the head is set correctly, but the LibreOffice UI will show "None". The return for Arrowhead type will be correct, $LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW.
;                  Libre Office has no setting for $bSync, so I have made a manual version of it in this function. It only accepts True, and must be called with True each time you want it to synchronize.
;                  When retrieving the current settings, $bSync will be a Boolean value of whether the Start Arrowhead settings are currently equal to the End Arrowhead setting values.
;                  Both $vStartStyle and $vEndStyle accept a String or an Integer because there is the possibility of a custom Arrowhead being available the user may want to use.
;                  When retrieving the current settings, both $vStartStyle and $vEndStyle could be either an integer or a String. It will be a String if the current Arrowhead is a custom Arrowhead, else an Integer, corresponding to one of the constants, $LOI_DRAWSHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeLineArrowStyles(ByRef $oShape, $vStartStyle = Null, $iStartWidth = Null, $bStartCenter = Null, $bSync = Null, $vEndStyle = Null, $iEndWidth = Null, $bEndCenter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avArrow[7]
	Local $sStartStyle, $sEndStyle

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($vStartStyle, $iStartWidth, $bStartCenter, $bSync, $vEndStyle, $iEndWidth, $bEndCenter) Then
		__LO_ArrayFill($avArrow, __LOImpress_DrawShapeArrowStyleName(Null, $oShape.LineStartName()), $oShape.LineStartWidth(), $oShape.LineStartCenter(), _
				((($oShape.LineStartName() = $oShape.LineEndName()) And ($oShape.LineStartWidth() = $oShape.LineEndWidth()) And ($oShape.LineStartCenter() = $oShape.LineEndCenter())) ? (True) : (False)), _ ; See if Start and End are the same.
				__LOImpress_DrawShapeArrowStyleName(Null, $oShape.LineEndName()), $oShape.LineEndWidth(), $oShape.LineEndCenter())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avArrow)
	EndIf

	If ($vStartStyle <> Null) Then
		If Not IsString($vStartStyle) And Not IsInt($vStartStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If IsInt($vStartStyle) Then
			If Not __LO_IntIsBetween($vStartStyle, $LOI_DRAWSHAPE_LINE_ARROW_TYPE_NONE, $LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$sStartStyle = __LOImpress_DrawShapeArrowStyleName($vStartStyle)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Else
			$sStartStyle = $vStartStyle
		EndIf

		$oShape.LineStartName = $sStartStyle
		$iError = ($oShape.LineStartName() = $sStartStyle) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iStartWidth <> Null) Then
		If Not __LO_IntIsBetween($iStartWidth, 0, 5004) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.LineStartWidth = $iStartWidth
		$iError = (__LO_IntIsBetween($oShape.LineStartWidth(), $iStartWidth - 1, $iStartWidth + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bStartCenter <> Null) Then
		If Not IsBool($bStartCenter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oShape.LineStartCenter = $bStartCenter
		$iError = ($oShape.LineStartCenter() = $bStartCenter) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bSync <> Null) Then
		If Not IsBool($bSync) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		If ($bSync = True) Then
			$oShape.LineEndName = $oShape.LineStartName()
			$oShape.LineEndWidth = $oShape.LineStartWidth()
			$oShape.LineEndCenter = $oShape.LineStartCenter()
			$iError = (($oShape.LineStartName() = $oShape.LineEndName()) And _
					($oShape.LineStartWidth() = $oShape.LineEndWidth()) And _
					($oShape.LineStartCenter() = $oShape.LineEndCenter())) ? ($iError) : (BitOR($iError, 8))
		EndIf
	EndIf

	If ($vEndStyle <> Null) Then
		If Not IsString($vEndStyle) And Not IsInt($vEndStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		If IsInt($vEndStyle) Then
			If Not __LO_IntIsBetween($vEndStyle, $LOI_DRAWSHAPE_LINE_ARROW_TYPE_NONE, $LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$sEndStyle = __LOImpress_DrawShapeArrowStyleName($vEndStyle)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Else
			$sEndStyle = $vEndStyle
		EndIf

		$oShape.LineEndName = $sEndStyle
		$iError = ($oShape.LineEndName() = $sEndStyle) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iEndWidth <> Null) Then
		If Not __LO_IntIsBetween($iEndWidth, 0, 5004) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oShape.LineEndWidth = $iEndWidth
		$iError = (__LO_IntIsBetween($oShape.LineEndWidth(), $iEndWidth - 1, $iEndWidth + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bEndCenter <> Null) Then
		If Not IsBool($bEndCenter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oShape.LineEndCenter = $bEndCenter
		$iError = ($oShape.LineEndCenter() = $bEndCenter) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeLineArrowStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeLineProperties
; Description ...: Set or Retrieve Shape Line settings.
; Syntax ........: _LOImpress_DrawShapeLineProperties(ByRef $oShape[, $vStyle = Null[, $iColor = Null[, $iWidth = Null[, $iTransparency = Null[, $iCornerStyle = Null[, $iCapStyle = Null]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $vStyle              - [optional] a variant value (0-31, or String). Default is Null. The Line Style to use. Can be a Custom Line Style name, or one of the constants, $LOI_DRAWSHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Line color, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iWidth              - [optional] an integer value (0-5004). Default is Null. The line Width, set in Micrometers.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The Line transparency percentage. 100% = fully transparent.
;                  $iCornerStyle        - [optional] an integer value (0,2-4). Default is Null. The Line Corner Style. See Constants $LOI_DRAWSHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  $iCapStyle           - [optional] an integer value (0-2). Default is Null. The Line Cap Style. See Constants $LOI_DRAWSHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $vStyle not a String, and not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $vStyle is an Integer, but less than 0, or greater than 31. See constants $LOI_DRAWSHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iWidth not an Integer, less than 0, or greater than 5004.
;                  @Error 1 @Extended 6 Return 0 = $iTransparency not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iCornerStyle not an Integer, not equal to 0, equal to 1, not equal to 2 or greater than 4. See Constants $LOI_DRAWSHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 8 Return 0 = $iCapStyle is an Integer, but less than 0, or greater than 2. See constants $LOI_DRAWSHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to convert Constant to Line Style name.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $vStyle accepts a String or an Integer because there is the possibility of a custom Line Style being available that the user may want to use.
;                  When retrieving the current settings, $vStyle could be either an integer or a String. It will be a String if the current Line Style is a custom Line Style, else an Integer, corresponding to one of the constants, $LOI_DRAWSHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOImpress_DrawShapeInsert, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeLineProperties(ByRef $oShape, $vStyle = Null, $iColor = Null, $iWidth = Null, $iTransparency = Null, $iCornerStyle = Null, $iCapStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local Const $__LOI_SHAPE_LINE_STYLE_NONE = 0, $__LOI_SHAPE_LINE_STYLE_SOLID = 1, $__LOI_SHAPE_LINE_STYLE_DASH = 2
	Local $avLine[6]
	Local $sStyle
	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($vStyle, $iColor, $iWidth, $iTransparency, $iCornerStyle, $iCapStyle) Then
		Switch $oShape.LineStyle()
			Case $__LOI_SHAPE_LINE_STYLE_NONE
				$vReturn = $LOI_DRAWSHAPE_LINE_STYLE_NONE

			Case $__LOI_SHAPE_LINE_STYLE_SOLID
				$vReturn = $LOI_DRAWSHAPE_LINE_STYLE_CONTINUOUS

			Case $__LOI_SHAPE_LINE_STYLE_DASH
				$vReturn = __LOImpress_DrawShapeLineStyleName(Null, $oShape.LineDashName())
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		EndSwitch

		__LO_ArrayFill($avLine, $vReturn, $oShape.LineColor(), $oShape.LineWidth(), $oShape.LineTransparence(), $oShape.LineJoint(), $oShape.LineCap())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avLine)
	EndIf

	If ($vStyle <> Null) Then
		If Not IsString($vStyle) And Not IsInt($vStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If IsInt($vStyle) Then
			If Not __LO_IntIsBetween($vStyle, $LOI_DRAWSHAPE_LINE_STYLE_NONE, $LOI_DRAWSHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			Switch $vStyle
				Case $LOI_DRAWSHAPE_LINE_STYLE_NONE
					$oShape.LineStyle = $__LOI_SHAPE_LINE_STYLE_NONE
					$iError = ($oShape.LineStyle() = $__LOI_SHAPE_LINE_STYLE_NONE) ? ($iError) : (BitOR($iError, 1))

				Case $LOI_DRAWSHAPE_LINE_STYLE_CONTINUOUS
					$oShape.LineStyle = $__LOI_SHAPE_LINE_STYLE_SOLID
					$iError = ($oShape.LineStyle() = $__LOI_SHAPE_LINE_STYLE_SOLID) ? ($iError) : (BitOR($iError, 1))

				Case Else
					$sStyle = __LOImpress_DrawShapeLineStyleName($vStyle)
					If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

					$oShape.LineStyle = $__LOI_SHAPE_LINE_STYLE_DASH
					$oShape.LineDashName = $sStyle
					$iError = ($oShape.LineDashName() = $sStyle) ? ($iError) : (BitOR($iError, 1))
			EndSwitch

		Else
			$sStyle = $vStyle
			$oShape.LineDashName = $sStyle
			$iError = ($oShape.LineDashName() = $sStyle) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.LineColor = $iColor
		$iError = ($oShape.LineColor() = $iColor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 0, 5004) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oShape.LineWidth = $iWidth
		$iError = (__LO_IntIsBetween($oShape.LineWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oShape.LineTransparence = $iTransparency
		$iError = ($oShape.LineTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iCornerStyle <> Null) Then
		If Not __LO_IntIsBetween($iCornerStyle, $LOI_DRAWSHAPE_LINE_JOINT_NONE, $LOI_DRAWSHAPE_LINE_JOINT_ROUND, $LOI_DRAWSHAPE_LINE_JOINT_MIDDLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oShape.LineJoint = $iCornerStyle
		$iError = ($oShape.LineJoint() = $iCornerStyle) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iCapStyle <> Null) Then
		If Not __LO_IntIsBetween($iCapStyle, $LOI_DRAWSHAPE_LINE_CAP_FLAT, $LOI_DRAWSHAPE_LINE_CAP_SQUARE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oShape.LineCap = $iCapStyle
		$iError = ($oShape.LineCap() = $iCapStyle) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeLineProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeName
; Description ...: Set or Retrieve a Shape's Name.
; Syntax ........: _LOImpress_DrawShapeName(ByRef $oShape[, $sName = Null])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $sName               - [optional] a string value. Default is Null. The new, unique Name for the Shape.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Document already contains a Shape with the same name as called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Slide Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sName
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Shape's name was successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Shape's current name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  On newly created shapes, the name value is blank for some reason, even though the shape in the UI has a name.
;                  The Shape name must be unique, however due to the above issue, it is possible to have two shapes with the same name in the UI (But not internally).
;                  I have modified the check for a Shape name already existing by assuming a Shape's name when encountering one without a name, e.g. assuming the name would be "Shape 1" etc.
; Related .......: _LOImpress_DrawShapeInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeName(ByRef $oShape, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oSlide

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($sName = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oShape.Name())

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSlide = $oShape.Parent()
	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If _LOImpress_DrawShapeExists($oSlide, $sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oShape.Name = $sName
	$iError = ($oShape.Name() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapePointsAdd
; Description ...: Add a Position Point to Shape.
; Syntax ........: _LOImpress_DrawShapePointsAdd(ByRef $oShape, $iPoint, $iX, $iY[, $iPointType = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL[, $bIsCurve = False]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function. See remarks.
;                  $iPoint              - an integer value. The Point to insert the new point AFTER. 0 means insert at the beginning.
;                  $iX                  - an integer value. The X coordinate value, set in Micrometers.
;                  $iY                  - an integer value. The Y coordinate value, set in Micrometers.
;                  $iPointType          - [optional] an integer value (0,1,3). Default is $LOI_DRAWSHAPE_POINT_TYPE_NORMAL. The Type of Point this new Point is. See Remarks. See constants $LOI_DRAWSHAPE_POINT_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $bIsCurve            - [optional] a boolean value. Default is False. If True, the Normal Point is a Curve. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oShape does not have property "PolyPolygonBezier", and consequently does not have Position Points that can be modified.
;                  @Error 1 @Extended 3 Return 0 = $iPoint not an Integer, less than 0 or greater than number of points in the shape.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer
;                  @Error 1 @Extended 6 Return 0 = $iPointType not an Integer, less than 0 or greater than 3, or equal to 2. See constants $LOI_DRAWSHAPE_POINT_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 7 Return 0 = $bIsCurve not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = First or Last Points in a shape can only be a "Normal" type point.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a new Position Point Structure.
;                  @Error 2 @Extended 2 Return 0 = Failed to Create a new Position Point Structure for the First Control Point.
;                  @Error 2 @Extended 3 Return 0 = Failed to Create a new Position Point Structure for the Second Control Point.
;                  @Error 2 @Extended 4 Return 0 = Failed to Create a new Position Point Structure for the Third Control Point.
;                  @Error 2 @Extended 5 Return 0 = Failed to Create a new Position Point Structure for the Fourth Control Point.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Retrieve Array of Point Type Flags.
;                  @Error 3 @Extended 2 Return 0 = Failed to Retrieve Array of Points.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify the requested Array element.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify the next normal Point in the Array of Points.
;                  @Error 3 @Extended 5 Return 0 = Failed to Retrieve PolyPolygonBezier Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. New Position Point was successfully added to the Shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only $LOI_DRAWSHAPE_TYPE_LINE_* type shapes have Points that can be added to, removed, or modified.
;                  This is a homemade function as LibreOffice doesn't offer an easy way for adding points to a shape. Consequently this will not produce similar results as when working with Libre office manually, and may wreck your shape's shape. Use with caution.
;                  For an unknown reason, I am unable to insert "SMOOTH" Points, and consequently, any smooth Points are sometimes reverted back to "Normal" points, but still having their Smooth control points upon insertion that were already present in the shape. If you call a new point with "SMOOTH" type, it will be, for now, replaced with "Symmetrical".
;                  The first and last points in a shape can only be a "Normal" Point Type. The last point cannot be Curved, but the first can be.
;                  Calling any Smooth or Symmetrical point types with $bIsCurve = True, will be ignored, as with the last point in a shape, as they are already a curve, or not supported in the case of the last point.
; Related .......: _LOImpress_DrawShapePointsModify, _LOImpress_DrawShapePointsRemove, _LOImpress_DrawShapePointsGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapePointsAdd(ByRef $oShape, $iPoint, $iX, $iY, $iPointType = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $bIsCurve = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tPoint, $tPolyCoords, $tControlPoint1, $tControlPoint2, $tControlPoint3, $tControlPoint4
	Local $iCount = 0, $iArrayElement, $iNextArrayElement, $iOffset = 0, $iForOffset = 0, $iReDimCount, $iSymmetricalPointXValue, $iSymmetricalPointYValue
	Local $aiFlags[0]
	Local $atPoints[0]
	Local $avArray[0], $avArray2[0]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oShape.getPropertySetInfo().hasPropertyByName("PolyPolygonBezier")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iPoint, 0, _LOImpress_DrawShapePointsGetCount($oShape)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Error if point called is not between 0 or number of points.
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not __LO_IntIsBetween($iPointType, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bIsCurve) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	; Temporary -- Smooth cannot be set, change it to Symmetrical -- Need to find a way to make work.
	If ($iPointType = $LOI_DRAWSHAPE_POINT_TYPE_SMOOTH) Then $iPointType = $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC

	$tPoint = __LOImpress_CreatePoint($iX, $iY)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$aiFlags = $oShape.PolyPolygonBezier.Flags()[0]
	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$atPoints = $oShape.PolyPolygonBezier.Coordinates()[0]
	If Not IsArray($atPoints) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If ($iPoint = 0) Then
		$iArrayElement = -1

	Else
		; Identify the Array element to add the point after.
		For $i = 0 To UBound($aiFlags) - 1
			If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iCount += 1 ; Skip any points that are Control Points, as they aren't actual points used for drawing the shape.

			If ($iCount = $iPoint) Then
				$iArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next
	EndIf

	If Not IsInt($iArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If ($iArrayElement = -1) Then ; Insertion will be at the beginning of the Points.

		If ($iPointType <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0) ; First Point in a shape can only be a "Normal" type Point.

		If ($bIsCurve = True) Then ; Point is a curve.

			ReDim $avArray[UBound($atPoints) + 3]
			ReDim $avArray2[UBound($aiFlags) + 3]
			; Make the control Point's Coordinates the new Point's Coordinates, plus half the difference between this new point and the next point, which will be the first element in the Points array.
			$tControlPoint1 = __LOImpress_CreatePoint(Int(($iX + (($atPoints[0]).X() - $iX) * .5)), Int(($iY + (($atPoints[0]).Y() - $iY) * .5)))
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			; Make the next control Point's Coordinates the Next Point's Coordinates, minus half the difference between this new point and the next point, which will be the first element in the Points array.
			$tControlPoint2 = __LOImpress_CreatePoint(Int($atPoints[0].X() - (($atPoints[0].X() - $iX) * .5)), Int($atPoints[0].Y() - (($atPoints[0].Y() - $iY) * .5)))
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

			$avArray[0] = $tPoint ; Place the new point at the beginning of the array.
			$avArray2[0] = $iPointType ; Place the new point's Type at the beginning of the array.
			$avArray[1] = $tControlPoint1 ; Place the two new Control Points next in the Array.
			$avArray2[1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL ; Place the two new Control Point's types next in the Array. Both are "Control" points.
			$avArray[2] = $tControlPoint2
			$avArray2[2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

			For $i = 3 To UBound($avArray) - 1
				$avArray[$i] = $atPoints[$i - 3] ; Add the rest of the points to the array.
				$avArray2[$i] = $aiFlags[$i - 3] ; Add the rest of the point's types to the array.

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$atPoints = $avArray
			$aiFlags = $avArray2

		Else ; Point is a regular point.
			ReDim $avArray[UBound($atPoints) + 1]
			ReDim $avArray2[UBound($aiFlags) + 1]

			$avArray[0] = $tPoint ; Place the new point at the beginning of the array.
			$avArray2[0] = $iPointType ; Place the new point's Type at the beginning of the array.

			For $i = 1 To UBound($avArray) - 1
				$avArray[$i] = $atPoints[$i - 1] ; Add the rest of the points to the array.
				$avArray2[$i] = $aiFlags[$i - 1] ; Add the rest of the point's types to the array.

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$atPoints = $avArray
			$aiFlags = $avArray2
		EndIf

	ElseIf ($iArrayElement = (UBound($aiFlags) - 1)) Then ; Insertion will be at the end of the Points.
		If ($iPointType <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0) ; Last Point in a shape can only be a "Normal" type Point.

		ReDim $avArray[UBound($atPoints) + 1]
		ReDim $avArray2[UBound($aiFlags) + 1]

		For $i = 0 To UBound($atPoints) - 1
			$avArray[$i] = $atPoints[$i] ; Add the rest of the points to the array.
			$avArray2[$i] = $aiFlags[$i] ; Add the rest of the point's types to the array.

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		$avArray[$i] = $tPoint ; Place the new point at the end of the array.
		$avArray2[$i] = $iPointType ; Place the new point's Type at the end of the array.

		$atPoints = $avArray
		$aiFlags = $avArray2

	Else ; Insertion is in the middle.
		For $i = ($iArrayElement + 1) To UBound($aiFlags) - 1 ; Locate the next non-Control Point in the Array for later use.
			If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$iNextArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		If Not IsInt($iNextArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If ($iPointType <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; Point Type is a curve of some form. Create four control points.

			; Check if the point I am placing the new point after is a normal point or not. If the Point Type after this point is a Control point,
			; the point I am inserting after is not a regular point.
			If ($aiFlags[$iArrayElement + 1] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$tControlPoint1 = __LOImpress_CreatePoint($atPoints[$iArrayElement].X(), $atPoints[$iArrayElement].Y()) ; If the point I am inserting after is normal, the control point has the same coordinates.
				If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			Else
				$tControlPoint1 = $atPoints[$iArrayElement + 1] ; Copy the existing Control point.
			EndIf

			; Pick the lowest X value difference between previous point and New point and Next point and New Point.
			$iSymmetricalPointXValue = ((($iX - $atPoints[$iArrayElement].X()) * .5) < (($atPoints[$iNextArrayElement].X() - $iX) * .5)) ? Int((($iX - $atPoints[$iArrayElement].X()) * .5)) : Int(($atPoints[$iNextArrayElement].X() - $iX) * .5)
			$iSymmetricalPointYValue = (((($iY - $atPoints[$iArrayElement].Y()) * .5)) < (($atPoints[$iNextArrayElement].Y() - $iY) * .5)) ? Int((($iY - $atPoints[$iArrayElement].Y()) * .5)) : Int((($atPoints[$iNextArrayElement].Y() - $iY) * .5))

			; Make the Second control Point's Coordinates the New Point's Coordinates, minus $iSymmetricalPointXValue and $iSymmetricalPointYValue
			$tControlPoint2 = __LOImpress_CreatePoint(($iX - $iSymmetricalPointXValue), ($iY - $iSymmetricalPointYValue))
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

			; Make the Third control Point's Coordinates the New Point's Coordinates, plus $iSymmetricalPointXValue and $iSymmetricalPointYValue
			$tControlPoint3 = __LOImpress_CreatePoint(($iX + $iSymmetricalPointXValue), ($iY + ($iSymmetricalPointYValue)))
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

			; Check if the second point after the point I am placing the new point after is a normal point or not. If the second Point Type after this point is a Control point, copy it.
			If ($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$tControlPoint4 = $atPoints[$iArrayElement + 2]

			Else
				; Make the Fourth control Point's Coordinates the Next Point's Coordinates, minus $iSymmetricalPointXValue and $iSymmetricalPointYValue
				$tControlPoint4 = __LOImpress_CreatePoint(($atPoints[$iNextArrayElement].X() - $iSymmetricalPointXValue), ($atPoints[$iNextArrayElement].Y() - $iSymmetricalPointYValue))
				If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)
			EndIf

			$iOffset = 0
			$iForOffset = 0
			$iReDimCount = 3
			; If point after the point I am inserting the new point after is a control point, don't add one to my Redim Count variable, as the element will be replaced. Else I need to add a
			; new element to the array for it.
			$iReDimCount += ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (0) : (1)
			$iReDimCount += (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (0) : (1)

			ReDim $avArray[UBound($atPoints) + $iReDimCount]
			ReDim $avArray2[UBound($aiFlags) + $iReDimCount]
			$iReDimCount = 0

			For $i = 0 To UBound($atPoints) - 1
				If ($iOffset = 0) Then
					$avArray[$i + $iForOffset] = $atPoints[$i + $iOffset] ; Add the rest of the points to the array.
					$avArray2[$i + $iForOffset] = $aiFlags[$i + $iOffset] ; Add the rest of the point's types to the array.

				Else
					$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
					$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
				EndIf

				If ($i = $iArrayElement) Then ; Insert the new point and its control points.

					$avArray[$i + 1] = $tControlPoint1
					$avArray2[$i + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
					$avArray[$i + 2] = $tControlPoint2
					$avArray2[$i + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
					$avArray[$i + 3] = $tPoint
					$avArray2[$i + 3] = $iPointType
					$avArray[$i + 4] = $tControlPoint3
					$avArray2[$i + 4] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
					$avArray[$i + 5] = $tControlPoint4
					$avArray2[$i + 5] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

					$iOffset += ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am inserting after has a control point after it, I need to skip them in the PointsArray.
					$iOffset += (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0) ; If the point I am inserting after has two control points after it, I need to skip them in the PointsArray.

					$iForOffset += 5 ; Add to $i to skip the elements I manually added.
				EndIf

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$atPoints = $avArray
			$aiFlags = $avArray2

		Else ; New Point is a Normal Point.
			If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Point after the point I am inserting my point at is a control point. I need to determine which point is a curve, and adjust as needed.

				; If the Point I am inserting after is not a normal Point, or if it is a normal point but the coordintes of the Point and the first control point after it are not identical,
				; (Indicating the Normal Point is set to "Create Curve"), modify the control points accordingly.
				If ($aiFlags[$iArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Or _
						(($atPoints[$iArrayElement].X() <> $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() <> $atPoints[$iArrayElement + 1].Y())) Then
					$tControlPoint1 = $atPoints[$iArrayElement + 1] ; Copy the first Control Point.

					; Make the Second control Point's Coordinates the New Point's Coordintes, minus half the difference between this new point and the previous point, which will be in the $iArrayElement of the Points array.
					$tControlPoint2 = __LOImpress_CreatePoint(Int($iX - (($iX - $atPoints[$iArrayElement].X()) * .5)), Int($iY - (($iY - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)
				EndIf

				If ($aiFlags[$iNextArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; If next Point after the one I am inserting is not a normal Point, modify the control points accordingly.

					; Make the Third control Point's Coordinates the New Point's Coordintes
					$tControlPoint3 = __LOImpress_CreatePoint($iX, $iY)
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

					If (($iArrayElement + 2 < $iNextArrayElement) And $atPoints[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; If the second point after the point I am inserting ahead of is a control point, copy it.

						$tControlPoint4 = $atPoints[$iArrayElement + 2] ; Copy the second control point after the point I am inserting after.

					Else ; Create the fourth point.
						; Make the Fourth control Point's Coordinates the Next Point's Coordintes, minus half the difference between this new point and the next point, which will be in the $iNextArrayElement of the Points array.
						$tControlPoint4 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $iX) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $iY) * .5)))
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)
					EndIf
				EndIf

				If ($bIsCurve = True) Then ; If the New Point is a Curved Normal point then either modify the third Control Point or create two new ones.

					; Make the Third control Point's Coordinates the New Point's Coordinates, plus half the difference between this new point and the next point, which will be in the $iNextArrayElement of the Points array.
					$tControlPoint3 = __LOImpress_CreatePoint(Int($iX + (($atPoints[$iNextArrayElement].X() - $iX) * .5)), Int($iY + (($atPoints[$iNextArrayElement].Y() - $iY) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

					If Not IsObj($tControlPoint4) Then ; If I haven't already made Control Point 4, create #4 and add two elements to the main array.

						; Make the Fourth control Point's Coordinates the Next Point's Coordinates, minus half the difference between this new point and the next point, which will be in the $iNextArrayElement of the Points array.
						$tControlPoint4 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $iX) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $iY) * .5)))
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)
					EndIf
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 1 ; Add one element to the array for the new point,

				; If I have created 4 control points add 4 to the Redim Count, else add two if either one or the other set have been created.
				If (IsObj($tControlPoint1) And IsObj($tControlPoint3)) Then
					$iReDimCount += 4

				ElseIf (IsObj($tControlPoint1) And IsObj($tControlPoint2)) Or (IsObj($tControlPoint3) And IsObj($tControlPoint4)) Then
					$iReDimCount += 2
				EndIf

				; If both or either point after the point I am inserting the new point after is a control point, remove one from my Redim Count variable, as the element will be replaced.
				; But only remove one if Redim count is greater than the one I added for my new point.
				$iReDimCount -= (($iReDimCount > 1) And ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)
				$iReDimCount -= (($iReDimCount > 1) And ($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) + $iReDimCount]
				ReDim $avArray2[UBound($atPoints) + $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the point's types to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iArrayElement) Then ; Insert the new point and its control points.

						If IsObj($tControlPoint1) Then ; If ControlPoint1 is an Object, that means both 1 and 2 need inserted.
							$avArray[$i + 1] = $tControlPoint1
							$avArray2[$i + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
							$avArray[$i + 2] = $tControlPoint2
							$avArray2[$i + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
							$avArray[$i + 3] = $tPoint
							$avArray2[$i + 3] = $iPointType
							$iForOffset += 3 ; Add 3 to $i Count.

							$iOffset += ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am inserting after has a control point after it, I need to skip it in the PointsArray.
							$iOffset += (($iArrayElement + 2 < $iNextArrayElement) And $aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am inserting after has two control points after it, I need to skip them in the PointsArray.

						Else
							$avArray[$i + 1] = $tPoint
							$avArray2[$i + 1] = $iPointType
							$iForOffset += 1
						EndIf

						If IsObj($tControlPoint3) Then ; If ControlPoint3 is an Object, that means both 3 and 4 need inserted.
							$avArray[$i + 2 + $iOffset] = $tControlPoint3
							$avArray2[$i + 2 + $iOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
							$avArray[$i + 3 + $iOffset] = $tControlPoint4
							$avArray2[$i + 3 + $iOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
							$iForOffset += 2

							If ($iOffset = 0) Then ; If I haven't already set Offset, check if it needs set.
								$iOffset += (($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0) ; If the point I am inserting after has a control point after it, I need to skip it in the PointsArray.
								$iOffset += (($iArrayElement + 2 < $iNextArrayElement) And $aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am inserting after has two control points after it, I need to skip them in the PointsArray.
							EndIf
						EndIf
					EndIf

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
				Next

				$atPoints = $avArray
				$aiFlags = $avArray2

			Else ; Point after the insertion point is a regular point.
				If ($bIsCurve = True) Then ; If the New Point is a Curved Normal point then create two new control Points.

					; Make the First control Point's Coordinates the New Point's Coordinates, plus half the difference between this new point and the next point, which will be in the $iNextArrayElement of the Points array.
					$tControlPoint1 = __LOImpress_CreatePoint(Int($iX + (($atPoints[$iNextArrayElement].X() - $iX) * .5)), Int($iY + (($atPoints[$iNextArrayElement].Y() - $iY) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

					; Make the Second control Point's Coordinates the Next Point's Coordinates, minus half the difference between this new point and the next point, which will be in the $iNextArrayElement of the Points array.
					$tControlPoint2 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $iX) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $iY) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

					$iReDimCount += 2 ; Add 2 elements in the Array because I had to create two control points.
				EndIf

				$iForOffset = 0
				$iReDimCount += 1 ; Add one element to the array for the new point.

				ReDim $avArray[UBound($atPoints) + $iReDimCount]
				ReDim $avArray2[UBound($atPoints) + $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
					$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the point's types to the array.

					If ($i = $iArrayElement) Then ; Insert the new point and its control points if applicable.

						If IsObj($tControlPoint1) Then ; If ControlPoint1 is an Object, that means both 1 and 2 need inserted.
							$avArray[$i + 1] = $tPoint
							$avArray2[$i + 1] = $iPointType
							$avArray[$i + 2] = $tControlPoint1
							$avArray2[$i + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
							$avArray[$i + 3] = $tControlPoint2
							$avArray2[$i + 3] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

							$iForOffset += 3 ; Add 2 to $i Count.

						Else
							$avArray[$i + 1] = $tPoint
							$avArray2[$i + 1] = $iPointType
							$iForOffset += 1
						EndIf
					EndIf

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
				Next

				$atPoints = $avArray
				$aiFlags = $avArray2
			EndIf
		EndIf
	EndIf

	$tPolyCoords = $oShape.PolyPolygonBezier()
	If Not IsObj($tPolyCoords) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	ReDim $avArray[1]

	; Each Array needs to be nested in an array.
	$avArray[0] = $atPoints
	$tPolyCoords.Coordinates = $avArray

	$avArray[0] = $aiFlags
	$tPolyCoords.Flags = $avArray

	; Set the  new Position Points for the Shape.
	$oShape.PolyPolygonBezier = $tPolyCoords

	; Apply it twice, as after inserting new points, the Point types get lost.
	$oShape.PolyPolygonBezier = $tPolyCoords

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DrawShapePointsAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapePointsGetCount
; Description ...: Retrieve a count of Points present in a Shape.
; Syntax ........: _LOImpress_DrawShapePointsGetCount(ByRef $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oShape does not have property "PolyPolygonBezier", and consequently does not have Position Points that can be modified.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Retrieve Array of Point Type Flags.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returns total number of points present in a shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only $LOI_DRAWSHAPE_TYPE_LINE_* type shapes have Points that can be added to, removed, or modified.
; Related .......: _LOImpress_DrawShapePointsAdd, _LOImpress_DrawShapePointsModify, _LOImpress_DrawShapePointsRemove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapePointsGetCount(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local $aiFlags[0]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oShape.getPropertySetInfo().hasPropertyByName("PolyPolygonBezier")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; Retrieve the Array of Point Type Constants. There is one flag per point, so I can just use these to count them by.
	$aiFlags = $oShape.PolyPolygonBezier.Flags()[0]
	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($aiFlags) - 1
		If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iCount += 1 ; Skip any points that are Control Points, as they aren't actual points used for drawing the shape.

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOImpress_DrawShapePointsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapePointsModify
; Description ...: Modify an existing Position Point or Point Type in a shape.
; Syntax ........: _LOImpress_DrawShapePointsModify(ByRef $oShape, $iPoint[, $iX = Null[, $iY = Null[, $iPointType = Null[, $bIsCurve = Null]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function. See remarks.
;                  $iPoint              - an integer value. The Point to modify, starting at 1.
;                  $iX                  - [optional] an integer value. Default is Null. The X coordinate value, set in Micrometers.
;                  $iY                  - [optional] an integer value. Default is Null. The Y coordinate value, set in Micrometers.
;                  $iPointType          - [optional] an integer value (0,1,3). Default is Null. The Type of Point to change the called point to. See Remarks. See constants $LOI_DRAWSHAPE_POINT_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $bIsCurve            - [optional] a boolean value. Default is Null. If True, the Normal Point is a Curve. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oShape does not have property "PolyPolygonBezier", and consequently does not have Position Points that can be modified.
;                  @Error 1 @Extended 3 Return 0 = $iPoint not an Integer, less than 1 or greater than number of points in the shape.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer
;                  @Error 1 @Extended 6 Return 0 = $PointType not an Integer, less than 0 or greater than 3, or equal to 2.
;                  @Error 1 @Extended 7 Return 0 = $PointType set to other than Normal while $iPoint is referencing first or last point.
;                  @Error 1 @Extended 8 Return 0 = $bIsCurve not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bIsCurve cannot be set for last point in a shape.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Retrieve Array of Point Type Flags.
;                  @Error 3 @Extended 2 Return 0 = Failed to Retrieve Array of Points.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify the requested Array element.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve current settings for requested point.
;                  @Error 3 @Extended 5 Return 0 = Failed to Retrieve PolyPolygonBezier Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Failed to modify the requested point.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings for the Array Element called in $iArrayElement.
;                  Call any optional parameter with Null keyword to skip it.
;                  Only $LOI_DRAWSHAPE_TYPE_LINE_* type shapes have Points that can be added to, removed, or modified.
;                  This is a homemade function as LibreOffice doesn't offer an easy way for modifying points in a shape. Consequently this will not produce similar results as when working with Libre office manually, and may wreck your shape's shape. Use with caution.
;                  For an unknown reason, I am unable to insert "SMOOTH" Points, and consequently, any smooth Points are reverted back to "Normal" points, but still having their Smooth control points upon insertion that were already present in the shape. If you modify a point to "SMOOTH" type, it will be, for now, replaced with "Symmetrical".
;                  The first and last points in a shape can only be a "Normal" Point Type. The last point cannot be Curved, but the first can be.
;                  Calling and Smooth or Symmetrical point types with $bIsCurve = True, will be ignored, as they are already a curve.
; Related .......: _LOImpress_DrawShapePointsAdd, _LOImpress_DrawShapePointsRemove, _LOImpress_DrawShapePointsGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapePointsModify(ByRef $oShape, $iPoint, $iX = Null, $iY = Null, $iPointType = Null, $bIsCurve = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0, $iArrayElement
	Local $tPolyCoords
	Local $aiFlags[0]
	Local $atPoints[0]
	Local $avPosPoint[4], $avArray[1]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oShape.getPropertySetInfo().hasPropertyByName("PolyPolygonBezier")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iPoint, 1, _LOImpress_DrawShapePointsGetCount($oShape)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Error if point called is not between 0 or number of points.

	$aiFlags = $oShape.PolyPolygonBezier.Flags()[0]
	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$atPoints = $oShape.PolyPolygonBezier.Coordinates()[0]
	If Not IsArray($atPoints) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	; Identify the Array element to modify the point.
	For $i = 0 To UBound($aiFlags) - 1
		If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iCount += 1 ; Skip any points that are Control Points, as they aren't actual points used for drawing the shape.

		If ($iCount = $iPoint) Then
			$iArrayElement = $i
			ExitLoop
		EndIf

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
	Next

	If Not IsInt($iArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If __LO_VarsAreNull($iX, $iY, $iPointType, $bIsCurve) Then
		__LOImpress_DrawShapePointGetSettings($avPosPoint, $aiFlags, $atPoints, $iArrayElement)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosPoint)
	EndIf

	If ($iX <> Null) Then
		If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	EndIf

	If ($iY <> Null) Then
		If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	EndIf

	If ($iPointType <> Null) Then
		If Not __LO_IntIsBetween($iPointType, $LOI_DRAWSHAPE_POINT_TYPE_NORMAL, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC, $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If ($iArrayElement = 0) Or ($iArrayElement = (UBound($atPoints) - 1)) And ($iPointType <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0) ; First or last point can't be a curve.

		; ## TEMPORARY
		If ($iPointType = $LOI_DRAWSHAPE_POINT_TYPE_SMOOTH) Then $iPointType = $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC
	EndIf

	If ($bIsCurve <> Null) Then
		If Not IsBool($bIsCurve) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		If ($iArrayElement = (UBound($atPoints) - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0) ; Last point cant be a curve.
	EndIf

	__LOImpress_DrawShapePointModify($aiFlags, $atPoints, $iArrayElement, $iX, $iY, $iPointType, $bIsCurve)
	If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	$tPolyCoords = $oShape.PolyPolygonBezier()
	If Not IsObj($tPolyCoords) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	; Each Array needs to be nested in an array.
	$avArray[0] = $atPoints
	$tPolyCoords.Coordinates = $avArray

	$avArray[0] = $aiFlags
	$tPolyCoords.Flags = $avArray

	; Set the modified Position Points for the Shape.
	$oShape.PolyPolygonBezier = $tPolyCoords

	; Apply it twice, as after modifying points, the Point types get lost.
	$oShape.PolyPolygonBezier = $tPolyCoords

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DrawShapePointsModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapePointsRemove
; Description ...: Remove a position Point from a Shape.
; Syntax ........: _LOImpress_DrawShapePointsRemove(ByRef $oShape, $iPoint)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iPoint              - an integer value. The Point to in the Shape to delete, beginning at 1.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oShape does not have property "PolyPolygonBezier", and consequently does not have Position Points that can be modified.
;                  @Error 1 @Extended 3 Return 0 = $iPoint not an Integer, less than 1 or greater than number of points in the shape.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a new Position Point Structure for the Second Control Point.
;                  @Error 2 @Extended 2 Return 0 = Failed to Create a new Position Point Structure for the Third Control Point.
;                  @Error 2 @Extended 3 Return 0 = Failed to Create a new Position Point Structure for the Fourth Control Point.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Retrieve Array of Point Type Flags.
;                  @Error 3 @Extended 2 Return 0 = Failed to Retrieve Array of Points.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify the requested Array element.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify the next normal Point in the Array of Points.
;                  @Error 3 @Extended 5 Return 0 = Failed to identify the Previous normal Point in the Array of Points.
;                  @Error 3 @Extended 6 Return 0 = Failed to Retrieve PolyPolygonBezier Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Position Point was successfully deleted from the Shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only $LOI_DRAWSHAPE_TYPE_LINE_* type shapes have Points that can be added to, removed, or modified.
;                  This is a homemade function as LibreOffice doesn't offer an easy way for removing points in a shape. Consequently this will not produce similar results as when working with Libre office manually, and may wreck your shape's shape. Use with caution.
;                  For an unknown reason, I am unable to insert "SMOOTH" Points, and consequently, any smooth Points are reverted back to "Normal" points, but still having their Smooth control points upon deletion that were already present in the shape. Some symmetrical points may revert also.
; Related .......: _LOImpress_DrawShapePointsAdd, _LOImpress_DrawShapePointsModify, _LOImpress_DrawShapePointsGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapePointsRemove(ByRef $oShape, $iPoint)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tPolyCoords, $tControlPoint1, $tControlPoint2, $tControlPoint3, $tControlPoint4
	Local $iOffset = 0, $iArrayElement, $iNextArrayElement, $iPreviousArrayElement, $iSkip = 0, $iCount = 0, $iReDimCount
	Local $avArray[0], $avArray2[0]
	Local $aiFlags[0]
	Local $atPoints[0]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oShape.getPropertySetInfo().hasPropertyByName("PolyPolygonBezier")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iPoint, 1, _LOImpress_DrawShapePointsGetCount($oShape)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Error if point called is not between 0 or number of points.

	$aiFlags = $oShape.PolyPolygonBezier.Flags()[0]
	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$atPoints = $oShape.PolyPolygonBezier.Coordinates()[0]
	If Not IsArray($atPoints) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	; Identify the Array element to remove the point.
	For $i = 0 To UBound($aiFlags) - 1
		If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iCount += 1 ; Skip any points that are Control Points, as they aren't actual points used for drawing the shape.

		If ($iCount = $iPoint) Then
			$iArrayElement = $i
			ExitLoop
		EndIf

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
	Next

	If Not IsInt($iArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If ($iArrayElement <> UBound($atPoints) - 1) Then ; If The requested point to be deleted is not at the end of the Array of points, find the next regular point.

		For $i = ($iArrayElement + 1) To UBound($aiFlags) - 1 ; Locate the next non-Control Point in the Array for later use.
			If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$iNextArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		If Not IsInt($iNextArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Else
		$iNextArrayElement = -1
	EndIf

	If ($iPoint > 1) Then ; If Point requested is not the first point, find the previous Point's position.

		For $i = ($iArrayElement - 1) To 0 Step -1 ; Locate the previous non-Control Point in the Array for later use.
			If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$iPreviousArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		If Not IsInt($iPreviousArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	Else
		$iPreviousArrayElement = -1
	EndIf

	If ($iArrayElement = 0) Then ; Point requested to be deleted is the first point.

		; Ensure next Point is a Normal Type Point.
		$aiFlags[$iNextArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

		ReDim $avArray[UBound($atPoints) - $iNextArrayElement]
		ReDim $avArray2[UBound($aiFlags) - $iNextArrayElement]

		For $i = 0 To (UBound($atPoints) - 1)
			If ($i >= $iNextArrayElement) Then
				$avArray[$i - $iSkip] = $atPoints[$i]
				$avArray2[$i - $iSkip] = $aiFlags[$i]

			Else
				$iSkip += 1
			EndIf

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		$atPoints = $avArray
		$aiFlags = $avArray2

	ElseIf ($iArrayElement = UBound($atPoints) - 1) Then ; Point requested to be deleted is the last point in the shape.
		; Ensure the second to last Normal point is a Normal Point.
		$aiFlags[$iPreviousArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

		ReDim $avArray[UBound($atPoints) - (UBound($atPoints) - $iPreviousArrayElement - 1)]
		ReDim $avArray2[UBound($aiFlags) - (UBound($aiFlags) - $iPreviousArrayElement - 1)]

		For $i = 0 To $iPreviousArrayElement + 1
			$avArray[$i] = $atPoints[$i]
			$avArray2[$i] = $aiFlags[$i]

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		$atPoints = $avArray
		$aiFlags = $avArray2

	Else ; Point to be deleted is in the middle.
		If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; If there is a control point after the Previous point.

			If ($aiFlags[$iPreviousArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; If Previous Point is not a normal point.

				$tControlPoint1 = $atPoints[$iPreviousArrayElement + 1] ; Copy the first control point after the previous point.

				If ($aiFlags[$iNextArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And (($iNextArrayElement - 1) > $iArrayElement) Then ; If Point before the next Point is a control point, copy it.
					$tControlPoint2 = $atPoints[$iNextArrayElement - 1]

				Else ; Point before the next Point is not a control point, create a new one.
					; Make the New control Point's Coordinates the Next Point's Coordinates, minus half the difference between the next point and the previous point.
					$tControlPoint2 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iPreviousArrayElement].X()) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iPreviousArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
				EndIf

			Else ; Previous Point is a normal point.
				; If the X and Y Coordinate of the previous point, and the control point after it do not match, the previous point is a "Curve".
				If ($atPoints[$iPreviousArrayElement].X() <> $atPoints[$iPreviousArrayElement + 1].X()) And ($atPoints[$iPreviousArrayElement].Y() <> $atPoints[$iPreviousArrayElement + 1].Y()) Then
					$tControlPoint1 = $atPoints[$iPreviousArrayElement + 1] ; Copy the first control point after the previous point.

					If ($aiFlags[$iNextArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And (($iNextArrayElement - 1) > $iArrayElement) Then ; Point before the next Point is a control point, copy it.
						$tControlPoint2 = $atPoints[$iNextArrayElement - 1]

					Else ; Point before the next Point is not a control point, create a new one.
						; Make the New control Point's Coordinates the Next Point's Coordinates, minus half the difference between the next point and the previous point.
						$tControlPoint2 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iPreviousArrayElement].X()) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iPreviousArrayElement].Y()) * .5)))
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
					EndIf
				EndIf
			EndIf

			$iOffset = 0
			$iSkip = 0
			$iReDimCount = 1 ; Start at one for the point I am deleting.
			$iReDimCount += ($aiFlags[$iNextArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point before the next point is a control point, add 1.
			$iReDimCount += (($aiFlags[$iNextArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And (($iNextArrayElement - 2) > $iArrayElement)) ? (1) : (0) ; If the second point before the next point is a control point, and still after the point to be deleted, add 1.
			$iReDimCount += ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point after the previous point is a control point, add 1.
			$iReDimCount += (($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And (($iPreviousArrayElement + 2) < $iArrayElement)) ? (1) : (0) ; If the second point after the previous point is a control point, and still before the point to be deleted, add 1.
			; If I had to create or retrieve control points to insert, minus 1 per from my Redim count.
			$iReDimCount -= (IsObj($tControlPoint1)) ? (1) : (0)
			$iReDimCount -= (IsObj($tControlPoint2)) ? (1) : (0)

			ReDim $avArray[UBound($atPoints) - $iReDimCount]
			ReDim $avArray2[UBound($aiFlags) - $iReDimCount]

			For $i = 0 To UBound($atPoints) - 1
				If ($i = $iArrayElement) Then
					$iOffset -= 1

				ElseIf ($iSkip = 0) Then
					$avArray[$i + $iOffset] = $atPoints[$i]
					$avArray2[$i + $iOffset] = $aiFlags[$i]

				Else
					$iSkip -= 1
					$iOffset -= 1
				EndIf

				If ($i = $iPreviousArrayElement) Then
					If IsObj($tControlPoint1) Then
						$avArray[$i + 1] = $tControlPoint1
						$avArray2[$i + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$iOffset += 1
					EndIf

					If IsObj($tControlPoint2) Then
						$avArray[$i + 2] = $tControlPoint2
						$avArray2[$i + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$iOffset += 1
					EndIf

					$iSkip += ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0)
					$iSkip += (($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And (($iPreviousArrayElement + 2) < $iArrayElement)) ? (1) : (0)
					$iSkip += ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0)
					$iSkip += (($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And (($iArrayElement + 2) < $iNextArrayElement)) ? (1) : (0)
				EndIf

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$atPoints = $avArray
			$aiFlags = $avArray2

		ElseIf ($aiFlags[$iNextArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; If the next point is not a Normal Point
			If ($aiFlags[$iNextArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$tControlPoint4 = $atPoints[$iNextArrayElement - 1]

			Else
				; Make the New control Point's Coordinates the Next Point's Coordinates, minus half the difference between the next point and the previous point.
				$tControlPoint4 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iPreviousArrayElement].X()) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iPreviousArrayElement].Y()) * .5)))
				If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)
			EndIf

			If ($aiFlags[$iNextArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And (($iNextArrayElement = 2) > $iArrayElement) Then
				$tControlPoint3 = $atPoints[$iNextArrayElement - 2]

			Else
				; Make the New control Point's Coordinates the same as the previous point.
				$tControlPoint3 = __LOImpress_CreatePoint($atPoints[$iPreviousArrayElement].X(), $atPoints[$iPreviousArrayElement].Y())
				If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
			EndIf

			$iOffset = 0
			$iSkip = 0
			$iReDimCount = 1 ; Start at one for the point I am deleting.

			ReDim $avArray[UBound($atPoints) - $iReDimCount]
			ReDim $avArray2[UBound($aiFlags) - $iReDimCount]

			For $i = 0 To UBound($atPoints) - 1
				If ($i = $iArrayElement) Then
					$iOffset -= 1

				ElseIf ($iSkip = 0) Then
					$avArray[$i + $iOffset] = $atPoints[$i]
					$avArray2[$i + $iOffset] = $aiFlags[$i]

				Else
					$iSkip -= 1
					$iOffset -= 1
				EndIf

				If ($i = $iArrayElement) Then
					If IsObj($tControlPoint3) Then
						$avArray[$i] = $tControlPoint3
						$avArray2[$i] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$iOffset += 1
					EndIf

					If IsObj($tControlPoint4) Then
						$avArray[$i + 1] = $tControlPoint4
						$avArray2[$i + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$iOffset += 1
					EndIf

					$iSkip += ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0)
					$iSkip += (($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And (($iArrayElement + 2) < $iNextArrayElement)) ? (1) : (0)
				EndIf

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$atPoints = $avArray
			$aiFlags = $avArray2

		Else ; There are no control points before or after the point to be deleted.
			ReDim $avArray[UBound($atPoints) - 1]
			ReDim $avArray2[UBound($aiFlags) - 1]

			For $i = 0 To UBound($atPoints) - 1
				If ($i = $iArrayElement) Then
					$iOffset -= 1

				Else
					$avArray[$i + $iOffset] = $atPoints[$i]
					$avArray2[$i + $iOffset] = $aiFlags[$i]
				EndIf

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$atPoints = $avArray
			$aiFlags = $avArray2
		EndIf
	EndIf

	$tPolyCoords = $oShape.PolyPolygonBezier()
	If Not IsObj($tPolyCoords) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	ReDim $avArray[1]

	; Each Array needs to be nested in an array.
	$avArray[0] = $atPoints
	$tPolyCoords.Coordinates = $avArray

	$avArray[0] = $aiFlags
	$tPolyCoords.Flags = $avArray

	; Set the  new Position Points for the Shape.
	$oShape.PolyPolygonBezier = $tPolyCoords

	; Apply it twice, as after modifying points, the Point types get lost.
	$oShape.PolyPolygonBezier = $tPolyCoords

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DrawShapePointsRemove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapePosition
; Description ...: Set or Retrieve the Shape's position settings.
; Syntax ........: _LOImpress_DrawShapePosition(ByRef $oShape[, $iX = Null[, $iY = Null[, $bProtectPos = Null]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iX                  - [optional] an integer value. Default is Null. The X position from the insertion point, in Micrometers.
;                  $iY                  - [optional] an integer value. Default is Null. The Y position from the insertion point, in Micrometers.
;                  $bProtectPos         - [optional] a boolean value. Default is Null. If True, the Shape's position is locked.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $bProtectPos not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape's Position Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iX
;                  |                               2 = Error setting $iY
;                  |                               4 = Error setting $bProtectPos
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DrawShapeInsert, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapePosition(ByRef $oShape, $iX = Null, $iY = Null, $bProtectPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPosition[3]
	Local $tPos

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iX, $iY, $bProtectPos) Then
		__LO_ArrayFill($avPosition, $tPos.X(), $tPos.Y(), $oShape.MoveProtect())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosition)
	EndIf

	If ($iX <> Null) Or ($iY <> Null) Then
		If ($iX <> Null) Then
			If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

			$tPos.X = $iX
		EndIf

		If ($iY <> Null) Then
			If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$tPos.Y = $iY
		EndIf

		$oShape.Position = $tPos

		$iError = ($iX = Null) ? ($iError) : ((__LO_IntIsBetween($oShape.Position.X(), $iX - 1, $iX + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = ($iY = Null) ? ($iError) : ((__LO_IntIsBetween($oShape.Position.Y(), $iY - 1, $iY + 1)) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($bProtectPos <> Null) Then
		If Not IsBool($bProtectPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.MoveProtect = $bProtectPos
		$iError = ($oShape.MoveProtect() = $bProtectPos) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapePosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeRotateSlant
; Description ...: Set or retrieve Rotation and Slant settings for a Shape.
; Syntax ........: _LOImpress_DrawShapeRotateSlant(ByRef $oShape[, $nRotate = Null[, $nSlant = Null]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $nRotate             - [optional] a general number value (0-359.99). Default is Null. The Degrees to rotate the shape. See remarks.
;                  $nSlant              - [optional] a general number value (-89-89.00). Default is Null. The Degrees to slant the shape. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $nRotate not a Number, less than 0, or greater than 359.99.
;                  @Error 1 @Extended 3 Return 0 = $nSlant not a Number, less than -89, or greater than 89.00.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $nRotate
;                  |                               2 = Error setting $nSlant
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you attempt to apply rotation to an already slanted Shape, or vice versa, a property setting error will occur, and the values will be very inaccurately applied.
;                  This function uses the deprecated Libre Office methods RotateAngle, and ShearAngle, and may stop working in future Libre Office versions, after 7.3.4.2.
;                  At the present time Control Point settings are not included as they are too complex to manipulate.
;                  At the present time Corner Radius setting is not included, as I was unable to identify a shape that utilized this setting.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOImpress_DrawShapeInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeRotateSlant(ByRef $oShape, $nRotate = Null, $nSlant = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aiShape[2]
	Local $iError = 0

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($nRotate, $nSlant) Then
		__LO_ArrayFill($aiShape, ($oShape.RotateAngle() / 100), ($oShape.ShearAngle() / 100)) ; Divide by 100 to match L.O. values.

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiShape)
	EndIf

	If ($nRotate <> Null) Then
		If Not __LO_NumIsBetween($nRotate, 0, 359.99) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oShape.RotateAngle = ($nRotate * 100) ; * 100 to match L.O. Values.
		$iError = (__LO_NumIsBetween(($oShape.RotateAngle() / 100), $nRotate - .01, $nRotate + .01)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($nSlant <> Null) Then
		If Not __LO_NumIsBetween($nSlant, -89, 89) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oShape.ShearAngle = ($nSlant * 100) ; * 100 to match L.O. Values.
		ConsoleWrite($oShape.ShearAngle() & @CRLF)
		$iError = (__LO_NumIsBetween(($oShape.ShearAngle() / 100), $nSlant - .01, $nSlant + .01)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeRotateSlant

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeTextboxCreateTextCursor
; Description ...: Create a Text Cursor in a Shape's Textbox for inserting text etc.
; Syntax ........: _LOImpress_DrawShapeTextboxCreateTextCursor(ByRef $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oShape not a shape supporting a Textbox.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Text Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. A Text Cursor Object located in the Textbox.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeTextboxCreateTextCursor(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oShape.PropertySetInfo.hasPropertyByName("Text") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextCursor = $oShape.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOImpress_DrawShapeTextboxCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeTypeSize
; Description ...: Set or Retrieve Shape Size related settings.
; Syntax ........: _LOImpress_DrawShapeTypeSize(ByRef $oShape[, $iWidth = Null[, $iHeight = Null[, $bProtectSize = Null]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Shape, in Micrometers(uM). Min. 51.
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the Shape, in Micrometers(uM). Min. 51.
;                  $bProtectSize        - [optional] a boolean value. Default is Null. If True, Locks the size of the Shape.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer, or less than 51.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer, or less than 51.
;                  @Error 1 @Extended 4 Return 0 = $bProtectSize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iHeight
;                  |                               4 = Error setting $bProtectSize
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I have skipped "Keep Ratio", as currently it seems unable to be set for shapes.
; Related .......: _LOImpress_DrawShapeInsert, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeTypeSize(ByRef $oShape, $iWidth = Null, $iHeight = Null, $bProtectSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSize[3]
	Local $tSize

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iHeight, $bProtectSize) Then
		__LO_ArrayFill($avSize, $tSize.Width(), $tSize.Height(), $oShape.SizeProtect())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSize)
	EndIf

	If ($iWidth <> Null) Or ($iHeight <> Null) Then
		If ($iWidth <> Null) Then ; Min 51
			If Not __LO_IntIsBetween($iWidth, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

			$tSize.Width = $iWidth
		EndIf

		If ($iHeight <> Null) Then
			If Not __LO_IntIsBetween($iHeight, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$tSize.Height = $iHeight
		EndIf

		$oShape.Size = $tSize

		$iError = ($iWidth = Null) ? ($iError) : ((__LO_IntIsBetween($oShape.Size.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = ($iHeight = Null) ? ($iError) : ((__LO_IntIsBetween($oShape.Size.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($bProtectSize <> Null) Then
		If Not IsBool($bProtectSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.SizeProtect = $bProtectSize
		$iError = ($oShape.SizeProtect() = $bProtectSize) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeTypeSize
