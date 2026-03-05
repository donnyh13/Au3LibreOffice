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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Deleting, etc. general Impress Shapes, such as Text Boxes.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Note...........: Many functions included in this file can be used to set Drawing shape properties as well.
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_ShapeAreaColor
; _LOImpress_ShapeAreaFillStyle
; _LOImpress_ShapeAreaGradient
; _LOImpress_ShapeAreaGradientMulticolor
; _LOImpress_ShapeAreaShadow
; _LOImpress_ShapeAreaTransparency
; _LOImpress_ShapeAreaTransparencyGradient
; _LOImpress_ShapeAreaTransparencyGradientMulti
; _LOImpress_ShapeCreateTextCursor
; _LOImpress_ShapeLineArrowStyles
; _LOImpress_ShapeLineProperties
; _LOImpress_ShapePosition
; _LOImpress_ShapeRotateSlant
; _LOImpress_ShapeSize
; _LOImpress_ShapeTextAttrAnimation
; _LOImpress_ShapeTextAttrColumns
; _LOImpress_ShapeTextAttrFit
; _LOImpress_ShapeTextAttrSettings
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaColor
; Description ...: Set or Retrieve the Fill color settings for a Shape.
; Syntax ........: _LOImpress_ShapeAreaColor(ByRef $oShape[, $iColor = Null])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is Null. The Fill color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current Fill color as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaColor(ByRef $oShape, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; If $iColor is Null, and Fill Style is set to solid, then return current color value, else return LO_COLOR_OFF.
	If __LO_VarsAreNull($iColor) Then Return SetError($__LO_STATUS_SUCCESS, 1, ($oShape.FillStyle() = $LOI_AREA_FILL_STYLE_SOLID) ? (__LOImpress_ColorRemoveAlpha($oShape.FillColor())) : ($LO_COLOR_OFF))

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
EndFunc   ;==>_LOImpress_ShapeAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOImpress_ShapeAreaFillStyle(ByRef $oShape)
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
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaFillStyle(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oShape.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOImpress_ShapeAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaGradient
; Description ...: Modify or retrieve the settings for Shape Background color Gradient.
; Syntax ........: _LOImpress_ShapeAreaGradient(ByRef $oShape[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
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
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sGradientName not a String.
;                  @Error 1 @Extended 3 Return 0 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error 1 @Extended 5 Return 0 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error 1 @Extended 8 Return 0 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iFromColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 10 Return 0 = $iToColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 11 Return 0 = $iFromIntense not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 12 Return 0 = $iToIntense not an Integer, less than 0 or greater than 100.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaGradient(ByRef $oShape, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
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
	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oShape.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oShape.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oShape.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oShape.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oShape.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = (__LO_VarsAreNull($iFromColor)) ? ($iError) : (($oShape.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = (__LO_VarsAreNull($iToColor)) ? ($iError) : (($oShape.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = (__LO_VarsAreNull($iFromIntense)) ? ($iError) : (($oShape.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = (__LO_VarsAreNull($iToIntense)) ? ($iError) : (($oShape.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaGradientMulticolor
; Description ...: Set or Retrieve a Shape's Multicolor Gradient settings. See remarks.
; Syntax ........: _LOImpress_ShapeAreaGradientMulticolor(ByRef $oShape[, $avColorStops = Null])
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
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, as a RGB Color Integer.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......: _LOImpress_GradientMulticolorAdd, _LOImpress_GradientMulticolorDelete, _LOImpress_GradientMulticolorModify, _LOImpress_ShapeAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaGradientMulticolor(ByRef $oShape, $avColorStops = Null)
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
EndFunc   ;==>_LOImpress_ShapeAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaShadow
; Description ...: Set or Retrieve the shadow settings for a Shape.
; Syntax ........: _LOImpress_ShapeAreaShadow(ByRef $oShape[, $bShadow = Null[, $iLocation = Null[, $iColor = Null[, $iDistance = Null[, $iBlur = Null[, $iTransparency = Null]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $bShadow             - [optional] a boolean value. Default is Null. If True, a Shadow is present for the Shape.
;                  $iLocation           - [optional] an integer value (0-8). Default is Null. The Location of the Shadow, must be one of the Constants, $LOI_SHAPE_SHADOW_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Shadow color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iDistance           - [optional] an integer value. Default is Null. The distance of the Shadow from the Shape's edges, set in Hundredths of a Millimeter (HMM).
;                  $iBlur               - [optional] an integer value (0-150). Default is Null. The amount of blur applied to the Shadow, set in Printer's Points.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The percentage of Shadow transparency. 100% means completely transparent.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bShadow not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $iLocation not an Integer, less than 0 or greater than 8. See Constants, $LOI_SHAPE_SHADOW_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iDistance not an Integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iBlur not an Integer, less than 0 or greater than 150 Printer's Points.
;                  @Error 1 @Extended 7 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotation Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Distance and Location Values.
;                  @Error 3 @Extended 3 Return 0 = Failed to modify Location property.
;                  @Error 3 @Extended 4 Return 0 = Failed to modify Distance property.
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
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaShadow(ByRef $oShape, $bShadow = Null, $iLocation = Null, $iColor = Null, $iDistance = Null, $iBlur = Null, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iInternalLocation, $iInternalDistance
	Local $avShadow[6]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bShadow, $iLocation, $iColor, $iDistance, $iBlur, $iTransparency) Then
		$iInternalDistance = __LOImpress_ShapeAreaShadowModify($oShape)
		$iInternalLocation = @extended
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		__LO_ArrayFill($avShadow, $oShape.Shadow(), $iInternalLocation, $oShape.ShadowColor(), $iInternalDistance, _
				_LO_UnitConvert($oShape.ShadowBlur(), $LO_CONVERT_UNIT_HMM_PT), _
				$oShape.ShadowTransparence())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($bShadow <> Null) Then
		If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oShape.Shadow = $bShadow
		$iError = ($oShape.Shadow() = $bShadow) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOI_SHAPE_SHADOW_TOP_LEFT, $LOI_SHAPE_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		__LOImpress_ShapeAreaShadowModify($oShape, $iLocation)
		If (@error = $__LO_STATUS_PROP_SETTING_ERROR) Then
			$iError = BitOR($iError, 2)

		ElseIf @error Then

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		EndIf
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.ShadowColor = $iColor
		$iError = ($oShape.ShadowColor() = $iColor) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iDistance <> Null) Then
		If Not __LO_IntIsBetween($iDistance, 0, $iDistance) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		__LOImpress_ShapeAreaShadowModify($oShape, Null, $iDistance)
		If (@error = $__LO_STATUS_PROP_SETTING_ERROR) Then
			$iError = BitOR($iError, 8)

		ElseIf @error Then

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
		EndIf
	EndIf

	If ($iBlur <> Null) Then
		If Not __LO_IntIsBetween($iBlur, 0, 150) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; 0 - 5292 max Hundredths of a Millimeter (HMM).

		$oShape.ShadowBlur = _LO_UnitConvert($iBlur, $LO_CONVERT_UNIT_PT_HMM)
		$iError = ($oShape.ShadowBlur() = _LO_UnitConvert($iBlur, $LO_CONVERT_UNIT_PT_HMM)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oShape.ShadowTransparence = $iTransparency
		$iError = ($oShape.ShadowTransparence = $iTransparency) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeAreaShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaTransparency
; Description ...: Set or retrieve Transparency settings for a Shape.
; Syntax ........: _LOImpress_ShapeAreaTransparency(ByRef $oShape[, $iTransparency = Null])
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
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaTransparency(ByRef $oShape, $iTransparency = Null)
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
EndFunc   ;==>_LOImpress_ShapeAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaTransparencyGradient
; Description ...: Set or retrieve the Shape transparency gradient settings.
; Syntax ........: _LOImpress_ShapeAreaTransparencyGradient(ByRef $oShape[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
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
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than -1 or greater than 5. See constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 4 Return 0 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 5 Return 0 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error 1 @Extended 6 Return 0 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iStart not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iEnd not an Integer, less than 0 or greater than 100.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaTransparencyGradient(ByRef $oShape, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
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

	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oShape.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oShape.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oShape.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oShape.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oShape.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iStart)) ? ($iError) : (($oShape.FillTransparenceGradient.StartColor() = __LOImpress_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iEnd)) ? ($iError) : (($oShape.FillTransparenceGradient.EndColor() = __LOImpress_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Shape's Multi Transparency Gradient settings. See remarks.
; Syntax ........: _LOImpress_ShapeAreaTransparencyGradientMulti(ByRef $oShape[, $avColorStops = Null])
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
;                  @Error 0 @Extended ? Return Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple Transparency stops in a Gradient rather than just a beginning and an ending value, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the Transparency percentage value between 0 and 100%.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......: _LOImpress_TransparencyGradientMultiModify, _LOImpress_TransparencyGradientMultiDelete, _LOImpress_TransparencyGradientMultiAdd, _LOImpress_ShapeAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaTransparencyGradientMulti(ByRef $oShape, $avColorStops = Null)
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
EndFunc   ;==>_LOImpress_ShapeAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCreateTextCursor
; Description ...: Create a Text Cursor in a Shape's Textbox for inserting text etc.
; Syntax ........: _LOImpress_ShapeCreateTextCursor(ByRef $oShape)
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
; Remarks .......: This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeCreateTextCursor(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oShape.supportsService("com.sun.star.drawing.Text") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextCursor = $oShape.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOImpress_ShapeCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeLineArrowStyles
; Description ...: Set or Retrieve Shape Line Start and End Arrow Style settings.
; Syntax ........: _LOImpress_ShapeLineArrowStyles(ByRef $oShape[, $vStartStyle = Null[, $iStartWidth = Null[, $bStartCenter = Null[, $bSync = Null[, $vEndStyle = Null[, $iEndWidth = Null[, $bEndCenter = Null]]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
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
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $vStartStyle not a String, and not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $vStartStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iStartWidth not an Integer, less than 0 or greater than 5004.
;                  @Error 1 @Extended 5 Return 0 = $bStartCenter not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bSync not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $vEndStyle not a String, and not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $vSEndStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $iEndWidth not an Integer, less than 0 or greater than 5004.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function works for connector shapes also.
;                  When the arrowhead type "Arrow" is set in the LO UI, or upon creation of a line with arrows, the internal name of the arrowhead is set to an incrementing name of "Arrowheads x", where x is an Integer value. Since I have no way to determine if the head is a custom arrowhead or supposed to be the "Arrow" type, the return when this is present will be the name "Arrowheads x", and not $LOI_SHAPE_LINE_ARROW_TYPE_ARROW.
;                  When setting an Arrowhead to be $LOI_SHAPE_LINE_ARROW_TYPE_ARROW, the head is set correctly, but the LibreOffice UI will show "None". The return for Arrowhead type will be correct, $LOI_SHAPE_LINE_ARROW_TYPE_ARROW.
;                  Libre Office has no setting for $bSync, so I have made a manual version of it in this function. It only accepts True, and must be called with True each time you want it to synchronize.
;                  When retrieving the current settings, $bSync will be a Boolean value of whether the Start Arrowhead settings are currently equal to the End Arrowhead setting values.
;                  Both $vStartStyle and $vEndStyle accept a String or an Integer because there is the possibility of a custom Arrowhead being available the user may want to use.
;                  When retrieving the current settings, both $vStartStyle and $vEndStyle could be either an Integer or a String. It will be a String if the current Arrowhead is a custom Arrowhead, else an Integer, corresponding to one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeLineArrowStyles(ByRef $oShape, $vStartStyle = Null, $iStartWidth = Null, $bStartCenter = Null, $bSync = Null, $vEndStyle = Null, $iEndWidth = Null, $bEndCenter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avArrow[7]
	Local $sStartStyle, $sEndStyle

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($vStartStyle, $iStartWidth, $bStartCenter, $bSync, $vEndStyle, $iEndWidth, $bEndCenter) Then
		__LO_ArrayFill($avArrow, __LOImpress_ShapeArrowStyleName(Null, $oShape.LineStartName()), $oShape.LineStartWidth(), $oShape.LineStartCenter(), _
				((($oShape.LineStartName() = $oShape.LineEndName()) And ($oShape.LineStartWidth() = $oShape.LineEndWidth()) And ($oShape.LineStartCenter() = $oShape.LineEndCenter())) ? (True) : (False)), _ ; See if Start and End are the same.
				__LOImpress_ShapeArrowStyleName(Null, $oShape.LineEndName()), $oShape.LineEndWidth(), $oShape.LineEndCenter())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avArrow)
	EndIf

	If ($vStartStyle <> Null) Then
		If Not IsString($vStartStyle) And Not IsInt($vStartStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If IsInt($vStartStyle) Then
			If Not __LO_IntIsBetween($vStartStyle, $LOI_SHAPE_LINE_ARROW_TYPE_NONE, $LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$sStartStyle = __LOImpress_ShapeArrowStyleName($vStartStyle)
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
			If Not __LO_IntIsBetween($vEndStyle, $LOI_SHAPE_LINE_ARROW_TYPE_NONE, $LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$sEndStyle = __LOImpress_ShapeArrowStyleName($vEndStyle)
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
EndFunc   ;==>_LOImpress_ShapeLineArrowStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeLineProperties
; Description ...: Set or Retrieve Shape Line settings.
; Syntax ........: _LOImpress_ShapeLineProperties(ByRef $oShape[, $vStyle = Null[, $iColor = Null[, $iWidth = Null[, $iTransparency = Null[, $iCornerStyle = Null[, $iCapStyle = Null]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $vStyle              - [optional] a variant value (0-31, or String). Default is Null. The Line Style to use. Can be a Custom Line Style name, or one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Line color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iWidth              - [optional] an integer value (0-5004). Default is Null. The line Width, set in Hundredths of a Millimeter (HMM).
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The Line transparency percentage. 100% = fully transparent.
;                  $iCornerStyle        - [optional] an integer value (0,2-4). Default is Null. The Line Corner Style. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  $iCapStyle           - [optional] an integer value (0-2). Default is Null. The Line Cap Style. See Constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $vStyle not a String, and not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $vStyle is an Integer, but less than 0 or greater than 31. See constants $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iWidth not an Integer, less than 0 or greater than 5004.
;                  @Error 1 @Extended 6 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iCornerStyle not an Integer, not equal to 0, equal to 1, not equal to 2 or greater than 4. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 8 Return 0 = $iCapStyle is an Integer, but less than 0 or greater than 2. See constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $vStyle accepts a String or an Integer because there is the possibility of a custom Line Style being available that the user may want to use.
;                  When retrieving the current settings, $vStyle could be either an Integer or a String. It will be a String if the current Line Style is a custom Line Style, else an Integer, corresponding to one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeLineProperties(ByRef $oShape, $vStyle = Null, $iColor = Null, $iWidth = Null, $iTransparency = Null, $iCornerStyle = Null, $iCapStyle = Null)
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
				$vReturn = $LOI_SHAPE_LINE_STYLE_NONE

			Case $__LOI_SHAPE_LINE_STYLE_SOLID
				$vReturn = $LOI_SHAPE_LINE_STYLE_CONTINUOUS

			Case $__LOI_SHAPE_LINE_STYLE_DASH
				$vReturn = __LOImpress_ShapeLineStyleName(Null, $oShape.LineDashName())
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		EndSwitch

		__LO_ArrayFill($avLine, $vReturn, $oShape.LineColor(), $oShape.LineWidth(), $oShape.LineTransparence(), $oShape.LineJoint(), $oShape.LineCap())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avLine)
	EndIf

	If ($vStyle <> Null) Then
		If Not IsString($vStyle) And Not IsInt($vStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If IsInt($vStyle) Then
			If Not __LO_IntIsBetween($vStyle, $LOI_SHAPE_LINE_STYLE_NONE, $LOI_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			Switch $vStyle
				Case $LOI_SHAPE_LINE_STYLE_NONE
					$oShape.LineStyle = $__LOI_SHAPE_LINE_STYLE_NONE
					$iError = ($oShape.LineStyle() = $__LOI_SHAPE_LINE_STYLE_NONE) ? ($iError) : (BitOR($iError, 1))

				Case $LOI_SHAPE_LINE_STYLE_CONTINUOUS
					$oShape.LineStyle = $__LOI_SHAPE_LINE_STYLE_SOLID
					$iError = ($oShape.LineStyle() = $__LOI_SHAPE_LINE_STYLE_SOLID) ? ($iError) : (BitOR($iError, 1))

				Case Else
					$sStyle = __LOImpress_ShapeLineStyleName($vStyle)
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
		If Not __LO_IntIsBetween($iCornerStyle, $LOI_SHAPE_LINE_JOINT_NONE, $LOI_SHAPE_LINE_JOINT_ROUND, $LOI_SHAPE_LINE_JOINT_MIDDLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oShape.LineJoint = $iCornerStyle
		$iError = ($oShape.LineJoint() = $iCornerStyle) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iCapStyle <> Null) Then
		If Not __LO_IntIsBetween($iCapStyle, $LOI_SHAPE_LINE_CAP_FLAT, $LOI_SHAPE_LINE_CAP_SQUARE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oShape.LineCap = $iCapStyle
		$iError = ($oShape.LineCap() = $iCapStyle) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeLineProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePosition
; Description ...: Set or Retrieve the Shape's position settings.
; Syntax ........: _LOImpress_ShapePosition(ByRef $oShape[, $iX = Null[, $iY = Null[, $bProtectPos = Null]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iX                  - [optional] an integer value. Default is Null. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - [optional] an integer value. Default is Null. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePosition(ByRef $oShape, $iX = Null, $iY = Null, $bProtectPos = Null)
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

		$iError = (__LO_VarsAreNull($iX)) ? ($iError) : ((__LO_IntIsBetween($oShape.Position.X(), $iX - 1, $iX + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = (__LO_VarsAreNull($iY)) ? ($iError) : ((__LO_IntIsBetween($oShape.Position.Y(), $iY - 1, $iY + 1)) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($bProtectPos <> Null) Then
		If Not IsBool($bProtectPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.MoveProtect = $bProtectPos
		$iError = ($oShape.MoveProtect() = $bProtectPos) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapePosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeRotateSlant
; Description ...: Set or retrieve Rotation and Slant settings for a Shape.
; Syntax ........: _LOImpress_ShapeRotateSlant(ByRef $oShape[, $nRotate = Null[, $nSlant = Null]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $nRotate             - [optional] a general number value (0-359.99). Default is Null. The Degrees to rotate the shape. See remarks.
;                  $nSlant              - [optional] a general number value (-89-89.00). Default is Null. The Degrees to slant the shape. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $nRotate not a Number, less than 0 or greater than 359.99.
;                  @Error 1 @Extended 3 Return 0 = $nSlant not a Number, less than -89 or greater than 89.00.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $nRotate
;                  |                               2 = Error setting $nSlant
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you attempt to apply rotation to an already slanted Shape, or vice versa, a property setting error will occur, and the values will be very inaccurately applied.
;                  This function uses the deprecated Libre Office methods RotateAngle, and ShearAngle, and may stop working in future Libre Office versions, after 7.3.4.2.
;                  At the present time Control Point settings are not included as they are too complex to manipulate.
;                  At the present time Corner Radius setting is not included, as I was unable to identify a shape that utilized this setting.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeRotateSlant(ByRef $oShape, $nRotate = Null, $nSlant = Null)
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
		$iError = (__LO_NumIsBetween(($oShape.ShearAngle() / 100), $nSlant - .01, $nSlant + .01)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeRotateSlant

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeSize
; Description ...: Set or Retrieve Shape Size related settings.
; Syntax ........: _LOImpress_ShapeSize(ByRef $oShape[, $iWidth = Null[, $iHeight = Null[, $bProtectSize = Null]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Shape, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the Shape, in Hundredths of a Millimeter (HMM). Min. 51.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I have skipped "Keep Ratio", as currently it seems unable to be set for shapes.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeSize(ByRef $oShape, $iWidth = Null, $iHeight = Null, $bProtectSize = Null)
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

		$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($oShape.Size.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = (__LO_VarsAreNull($iHeight)) ? ($iError) : ((__LO_IntIsBetween($oShape.Size.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($bProtectSize <> Null) Then
		If Not IsBool($bProtectSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.SizeProtect = $bProtectSize
		$iError = ($oShape.SizeProtect() = $bProtectSize) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeTextAttrAnimation
; Description ...: Set or Retrieve Shape Text Attribute Animation settings.
; Syntax ........: _LOImpress_ShapeTextAttrAnimation(ByRef $oShape[, $iEffect = Null[, $iDirection = Null[, $bStartInside = Null[, $bVisibleOnExit = Null[, $iCycles = Null[, $iInc = Null[, $bPixels = Null[, $iDelay = Null]]]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Dimension Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iEffect             - [optional] an integer value (0-4). Default is Null. The Animation type. See Constants, $LOI_ANIMATION_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDirection          - [optional] an integer value (0-3). Default is Null. The Direction of the text's movement, if applicable. See Constants, $LOI_ANIMATION_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bStartInside        - [optional] a boolean value. Default is Null. If True, Text is visible and inside the shape when the effect is applied.
;                  $bVisibleOnExit      - [optional] a boolean value. Default is Null. If True, Text remains visible after the effect is applied.
;                  $iCycles             - [optional] an integer value (0-100). Default is Null. The number of times to repeat the animation. 0 = Continuous.
;                  $iInc                - [optional] an integer value (1-100px/25-32,766). Default is Null. the increment value for scrolling the text, in Hundredths of a Millimeter (HMM), or pixels.
;                  $bPixels             - [optional] a boolean value. Default is Null. If True, $iInc is set in pixels, else in Hundredths of a Millimeter (HMM).
;                  $iDelay              - [optional] an integer value (0-30,000). Default is Null. The amount time (ms) to wait before repeating the effect.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iEffect not an Integer, less than 0 or greater than 4. See Constants, $LOI_ANIMATION_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iDirection not an Integer, less than 0 or greater than 3. See Constants, $LOI_ANIMATION_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $bStartInside not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bVisibleOnExit not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iCycles not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iInc not an Integer, less than 1 or greater than 100 pixels, less than 25 or greater than 32,766 Hundredths of a Millimeter (HMM).
;                  @Error 1 @Extended 8 Return 0 = $bPixels not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iDelay not an Integer, less than 0 or greater than 30,000.
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
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeTextAttrAnimation(ByRef $oShape, $iEffect = Null, $iDirection = Null, $bStartInside = Null, $bVisibleOnExit = Null, $iCycles = Null, $iInc = Null, $bPixels = Null, $iDelay = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iValue
	Local $avTextAttr[8]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iEffect, $iDirection, $bStartInside, $bVisibleOnExit, $iCycles, $iInc, $bPixels, $iDelay) Then
		__LO_ArrayFill($avTextAttr, $oShape.TextAnimationKind(), $oShape.TextAnimationDirection(), $oShape.TextAnimationStartInside(), _
				$oShape.TextAnimationStopInside(), $oShape.TextAnimationCount(), _
				($oShape.TextAnimationAmount() < 0) ? ($oShape.TextAnimationAmount() * -1) : ($oShape.TextAnimationAmount()), _ ; If TextAnimationAmount is negative, Pixels are used, if positive Hundredths of a Millimeter (HMM).
				($oShape.TextAnimationAmount() < 0) ? (True) : (False), _ ; $bPixels
				$oShape.TextAnimationDelay())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTextAttr)
	EndIf

	If ($iEffect <> Null) Then
		If Not __LO_IntIsBetween($iEffect, $LOI_ANIMATION_TYPE_NONE, $LOI_ANIMATION_TYPE_SCROLL_IN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oShape.TextAnimationKind = $iEffect
		$iError = ($oShape.TextAnimationKind() = $iEffect) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iDirection <> Null) Then
		If Not __LO_IntIsBetween($iDirection, $LOI_ANIMATION_DIR_LEFT, $LOI_ANIMATION_DIR_DOWN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oShape.TextAnimationDirection = $iDirection
		$iError = ($oShape.TextAnimationDirection() = $iDirection) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bStartInside <> Null) Then
		If Not IsBool($bStartInside) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.TextAnimationStartInside = $bStartInside
		$iError = ($oShape.TextAnimationStartInside() = $bStartInside) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bVisibleOnExit <> Null) Then
		If Not IsBool($bVisibleOnExit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oShape.TextAnimationStopInside = $bVisibleOnExit
		$iError = ($oShape.TextAnimationStopInside() = $bVisibleOnExit) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iCycles <> Null) Then
		If Not __LO_IntIsBetween($iCycles, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oShape.TextAnimationCount = $iCycles
		$iError = ($oShape.TextAnimationCount() = $iCycles) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iInc <> Null) Then
		If (($oShape.TextAnimationAmount() < 0) And ($bPixels <> False)) Or ($bPixels = True) Then ; Set in Pixels
			If Not __LO_IntIsBetween($iInc, 1, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oShape.TextAnimationAmount = ($iInc * -1) ; Multiply by -1 to change to negative, as Pixels are set in negative values.
			$iError = ($oShape.TextAnimationAmount() = ($iInc * -1)) ? ($iError) : (BitOR($iError, 32))

		Else ; Set in Hundredths of a Millimeter (HMM).
			If Not __LO_IntIsBetween($iInc, 25, 32766) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oShape.TextAnimationAmount = $iInc
			$iError = ($oShape.TextAnimationAmount() = $iInc) ? ($iError) : (BitOR($iError, 32))
		EndIf
	EndIf

	If ($bPixels <> Null) Then
		If Not IsBool($bPixels) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$iValue = $oShape.TextAnimationAmount()
		If Not IsInt($iValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If $bPixels Then
			If ($iValue > 0) Then ; Not set to pixels yet, convert to pixels
				$iValue = ($iValue > 100) ? (-100) : ($iValue * -1) ; If greater than 100 pixels (the max) set to 100, otherwise convert the value to negative for pixels.
				$oShape.TextAnimationAmount = $iValue
			EndIf

		Else
			If ($iValue < 0) Then ; Set to pixels, convert to Hundredths of a Millimeter (HMM).
				$iValue = ($iValue * -1) ; Convert the value to positive for Hundredths of a Millimeter (HMM).
				$oShape.TextAnimationAmount = $iValue
			EndIf
		EndIf
		$iError = ($oShape.TextAnimationAmount() = $iValue) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 30000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oShape.TextAnimationDelay = $iDelay
		$iError = ($oShape.TextAnimationDelay() = $iDelay) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeTextAttrAnimation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeTextAttrColumns
; Description ...: Set or Retrieve Shape Text Attribute Column settings. (L.O. 7.2+)
; Syntax ........: _LOImpress_ShapeTextAttrColumns(ByRef $oShape[, $iColumns = Null[, $iSpacing = Null]])
; Parameters ....: $oShape              - [in/out] an object. A Dimension Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iColumns            - [optional] an integer value (1-16). Default is Null. The number of columns.
;                  $iSpacing            - [optional] an integer value. Default is Null. The spacing between each column, in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumns not an Integer, less than 1 or greater than 16.
;                  @Error 1 @Extended 3 Return 0 = $iSpacing not an Integer, less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.text.TextColumns Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve parent Document Object.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current version is less than 7.2.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColumns
;                  |                               2 = Error setting $iSpacing
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeTextAttrColumns(ByRef $oShape, $iColumns = Null, $iSpacing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oDoc, $oTextColumns
	Local $aiColumns[2]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$oDoc = $oShape.Parent.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTextColumns = $oShape.TextColumns()

	If __LO_VarsAreNull($iColumns, $iSpacing) Then
		__LO_ArrayFill($aiColumns, _
				(IsObj($oTextColumns)) ? ($oTextColumns.ColumnCount()) : (1), _ ; If No text columns are set for a new shape, TextColumns will be Null, return default values.
				(IsObj($oTextColumns)) ? ($oTextColumns.AutomaticDistance()) : (0)) ; If No text columns are set for a new shape, TextColumns will be Null, return default values.

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiColumns)
	EndIf

	If Not IsObj($oTextColumns) Then ; Create a TextColumns service if there was none.
		$oTextColumns = $oDoc.createInstance("com.sun.star.text.TextColumns")
		If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oTextColumns.ColumnCount = 1
	EndIf

	If ($iColumns <> Null) Then
		If Not __LO_IntIsBetween($iColumns, 1, 16) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oTextColumns.ColumnCount = $iColumns
		$oShape.TextColumns = $oTextColumns
		$iError = ($oShape.TextColumns.ColumnCount() = $iColumns) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iSpacing <> Null) Then
		If Not __LO_IntIsBetween($iSpacing, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTextColumns.AutomaticDistance = $iSpacing
		$oShape.TextColumns = $oTextColumns
		$iError = (__LO_IntIsBetween($oShape.TextColumns.AutomaticDistance(), $iSpacing - 1, $iSpacing + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeTextAttrColumns

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeTextAttrFit
; Description ...: Set or Retrieve Shape Text Attribute Fit properties. See Remarks.
; Syntax ........: _LOImpress_ShapeTextAttrFit(ByRef $oShape[, $bFitToFrame = Null[, $bAdjustContour = Null[, $bFitWidth = Null[, $bFitHeight = Null[, $bWordWrap = Null[, $bResizeShape = Null]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $bFitToFrame         - [optional] a boolean value. Default is Null. If True, Resizes the text to fit the entire area of the drawing object.
;                  $bAdjustContour      - [optional] a boolean value. Default is Null. If True, Adapts the text flow so that it matches the contours of the drawing object.
;                  $bFitWidth           - [optional] a boolean value. Default is Null. If True, Expands the width of the object to the width of the text.
;                  $bFitHeight          - [optional] a boolean value. Default is Null. If True, Expands the height of the object to the height of the text.
;                  $bWordWrap           - [optional] a boolean value. Default is Null. If True, Wraps the text to fit inside the shape.
;                  $bResizeShape        - [optional] a boolean value. Default is Null. If True, Resizes a custom shape to fit the text that you enter.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFitToFrame not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bAdjustContour not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bFitWidth not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bFitHeight not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bWordWrap not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bResizeShape not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFitToFrame
;                  |                               2 = Error setting $bAdjustContour
;                  |                               4 = Error setting $bFitWidth
;                  |                               8 = Error setting $bFitHeight
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
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeTextAttrFit(ByRef $oShape, $bFitToFrame = Null, $bAdjustContour = Null, $bFitWidth = Null, $bFitHeight = Null, $bWordWrap = Null, $bResizeShape = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__LOI_TEXT_FIT_NONE = 0, $__LOI_TEXT_FIT_PROP = 1 ; com.sun.star.drawing.TextFitToSizeType
	Local $iError = 0
	Local $avTextAttr[6]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bFitToFrame, $bAdjustContour, $bFitWidth, $bFitHeight, $bWordWrap, $bResizeShape) Then
		__LO_ArrayFill($avTextAttr, ($oShape.TextFitToSize() = $__LOI_TEXT_FIT_PROP) ? (True) : (False), _
				$oShape.TextContourFrame(), $oShape.TextAutoGrowWidth(), $oShape.TextAutoGrowHeight(), $oShape.TextWordWrap(), $oShape.TextAutoGrowHeight())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTextAttr)
	EndIf

	If ($bFitToFrame <> Null) Then
		If Not IsBool($bFitToFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oShape.TextFitToSize = ($bFitToFrame) ? ($__LOI_TEXT_FIT_PROP) : ($__LOI_TEXT_FIT_NONE)
		$iError = ($oShape.TextFitToSize() = ($bFitToFrame) ? ($__LOI_TEXT_FIT_PROP) : ($__LOI_TEXT_FIT_NONE)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bAdjustContour <> Null) Then
		If Not IsBool($bAdjustContour) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oShape.TextContourFrame = $bAdjustContour
		$iError = ($oShape.TextContourFrame() = $bAdjustContour) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bFitWidth <> Null) Then
		If Not IsBool($bFitWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.TextAutoGrowWidth = $bFitWidth
		$iError = ($oShape.TextAutoGrowWidth() = $bFitWidth) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bFitHeight <> Null) Then
		If Not IsBool($bFitHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oShape.TextAutoGrowHeight = $bFitHeight
		$iError = ($oShape.TextAutoGrowHeight() = $bFitHeight) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bWordWrap <> Null) Then
		If Not IsBool($bWordWrap) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oShape.TextWordWrap = $bWordWrap
		$iError = ($oShape.TextWordWrap() = $bWordWrap) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bResizeShape <> Null) Then
		If Not IsBool($bResizeShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oShape.TextAutoGrowHeight = $bResizeShape
		$iError = ($oShape.TextAutoGrowHeight() = $bResizeShape) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeTextAttrFit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeTextAttrSettings
; Description ...: Set or Retrieve Shape text Attribute settings.
; Syntax ........: _LOImpress_ShapeTextAttrSettings(ByRef $oShape[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null[, $iAnchor = Null[, $bFullWidth = Null]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Dimension Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $iLeft               - [optional] an integer value (-100,000-100,000). Default is Null. The space between the left edge of the drawing object and the left border of the text, in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value (-100,000-100,000). Default is Null. The space between the right edge of the drawing object and the right border of the text, in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value (-100,000-100,000). Default is Null. The space between the top edge of the drawing object and the top border of the text, in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value (-100,000-100,000). Default is Null. The space between the bottom edge of the drawing object and the bottom border of the text, in Hundredths of a Millimeter (HMM).
;                  $iAnchor             - [optional] an integer value (0-8). Default is Null. The text anchor position. See Constants, $LOI_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bFullWidth          - [optional] a boolean value. Default is Null. If True, Anchors the text to the full width of the drawing object.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iLeft not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 3 Return 0 = $iRight not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 4 Return 0 = $iTop not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 5 Return 0 = $iBottom not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 6 Return 0 = $iAnchor  not an Integer, less than 0 or greater than 8. See Constants, $LOI_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
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
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_SlideShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeTextAttrSettings(ByRef $oShape, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null, $iAnchor = Null, $bFullWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurAnchor
	Local $avTextAttr[6]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLeft, $iRight, $iTop, $iBottom, $iAnchor, $bFullWidth) Then
		Select
			Case ($oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_TOP) And ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_LEFT)
				$iCurAnchor = $LOI_TEXT_ANCHOR_TOP_LEFT

			Case ($oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_TOP) And (($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_CENTER) Or ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_BLOCK))
				$iCurAnchor = $LOI_TEXT_ANCHOR_TOP_CENTER

			Case ($oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_TOP) And ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_RIGHT)
				$iCurAnchor = $LOI_TEXT_ANCHOR_TOP_RIGHT

			Case ($oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_CENTER) And ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_LEFT)
				$iCurAnchor = $LOI_TEXT_ANCHOR_MIDDLE_LEFT

			Case ($oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_CENTER) And (($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_CENTER) Or ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_BLOCK))
				$iCurAnchor = $LOI_TEXT_ANCHOR_MIDDLE_CENTER

			Case ($oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_CENTER) And ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_RIGHT)
				$iCurAnchor = $LOI_TEXT_ANCHOR_MIDDLE_RIGHT

			Case ($oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_BOTTOM) And ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_LEFT)
				$iCurAnchor = $LOI_TEXT_ANCHOR_BOTTOM_LEFT

			Case ($oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_BOTTOM) And (($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_CENTER) Or ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_BLOCK))
				$iCurAnchor = $LOI_TEXT_ANCHOR_BOTTOM_CENTER

			Case ($oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_BOTTOM) And ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_RIGHT)
				$iCurAnchor = $LOI_TEXT_ANCHOR_BOTTOM_RIGHT
		EndSelect

		__LO_ArrayFill($avTextAttr, $oShape.TextLeftDistance(), $oShape.TextRightDistance(), $oShape.TextUpperDistance(), $oShape.TextLowerDistance(), _
				$iCurAnchor, ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_BLOCK) ? (True) : (False))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTextAttr)
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oShape.TextLeftDistance = $iLeft
		$iError = (__LO_IntIsBetween($oShape.TextLeftDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oShape.TextRightDistance = $iRight
		$iError = (__LO_IntIsBetween($oShape.TextRightDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.TextUpperDistance = $iTop
		$iError = (__LO_IntIsBetween($oShape.TextUpperDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oShape.TextLowerDistance = $iBottom
		$iError = (__LO_IntIsBetween($oShape.TextLowerDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iAnchor <> Null) Then
		If Not __LO_IntIsBetween($iAnchor, $LOI_TEXT_ANCHOR_TOP_LEFT, $LOI_TEXT_ANCHOR_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		Switch $iAnchor
			Case $LOI_TEXT_ANCHOR_TOP_LEFT
				$oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_TOP
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_LEFT
				$iError = ($oShape.TextVerticalAdjust() = $LOI_TEXT_ALIGN_VERT_TOP) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_LEFT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_TEXT_ANCHOR_TOP_CENTER
				$oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_TOP
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_CENTER
				$iError = ($oShape.TextVerticalAdjust() = $LOI_TEXT_ALIGN_VERT_TOP) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_CENTER) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_TEXT_ANCHOR_TOP_RIGHT
				$oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_TOP
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_RIGHT
				$iError = ($oShape.TextVerticalAdjust() = $LOI_TEXT_ALIGN_VERT_TOP) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_RIGHT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_TEXT_ANCHOR_MIDDLE_LEFT
				$oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_CENTER
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_LEFT
				$iError = ($oShape.TextVerticalAdjust() = $LOI_TEXT_ALIGN_VERT_CENTER) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_LEFT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_TEXT_ANCHOR_MIDDLE_CENTER
				$oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_CENTER
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_CENTER
				$iError = ($oShape.TextVerticalAdjust() = $LOI_TEXT_ALIGN_VERT_CENTER) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_CENTER) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_TEXT_ANCHOR_MIDDLE_RIGHT
				$oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_CENTER
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_RIGHT
				$iError = ($oShape.TextVerticalAdjust() = $LOI_TEXT_ALIGN_VERT_CENTER) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_RIGHT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_TEXT_ANCHOR_BOTTOM_LEFT
				$oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_BOTTOM
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_LEFT
				$iError = ($oShape.TextVerticalAdjust() = $LOI_TEXT_ALIGN_VERT_BOTTOM) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_LEFT) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_TEXT_ANCHOR_BOTTOM_CENTER
				$oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_BOTTOM
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_CENTER
				$iError = ($oShape.TextVerticalAdjust() = $LOI_TEXT_ALIGN_VERT_BOTTOM) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_CENTER) ? ($iError) : (BitOR($iError, 16))

			Case $LOI_TEXT_ANCHOR_BOTTOM_RIGHT
				$oShape.TextVerticalAdjust = $LOI_TEXT_ALIGN_VERT_BOTTOM
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_RIGHT
				$iError = ($oShape.TextVerticalAdjust() = $LOI_TEXT_ALIGN_VERT_BOTTOM) ? ($iError) : (BitOR($iError, 16))
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_RIGHT) ? ($iError) : (BitOR($iError, 16))
		EndSwitch
	EndIf

	If ($bFullWidth <> Null) Then
		If Not IsBool($bFullWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		If $bFullWidth Then
			$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_BLOCK
			$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_BLOCK) ? ($iError) : (BitOR($iError, 32))

		Else
			If ($oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_BLOCK) Then ; Only set Horizontal Adjust to Center if it was set to Block already when setting $bFullWidth to False.
				$oShape.TextHorizontalAdjust = $LOI_TEXT_ALIGN_HORI_CENTER
				$iError = ($oShape.TextHorizontalAdjust() = $LOI_TEXT_ALIGN_HORI_CENTER) ? ($iError) : (BitOR($iError, 32))
			EndIf
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeTextAttrSettings
