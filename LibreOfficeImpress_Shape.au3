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
; _LOImpress_ShapeCharEffect
; _LOImpress_ShapeCharFont
; _LOImpress_ShapeCharFontColor
; _LOImpress_ShapeCharOverLine
; _LOImpress_ShapeCharPosition
; _LOImpress_ShapeCharScaling
; _LOImpress_ShapeCharSpacing
; _LOImpress_ShapeCharStrikeOut
; _LOImpress_ShapeCharUnderLine
; _LOImpress_ShapeCreateTextCursor
; _LOImpress_ShapeDelete
; _LOImpress_ShapeExists
; _LOImpress_ShapeImageAltText
; _LOImpress_ShapeImageCrop
; _LOImpress_ShapeImageInsert
; _LOImpress_ShapeImageModify
; _LOImpress_ShapeImageReplace
; _LOImpress_ShapeInteraction
; _LOImpress_ShapeLineArrowStyles
; _LOImpress_ShapeLineProperties
; _LOImpress_ShapeName
; _LOImpress_ShapeParAlignment
; _LOImpress_ShapeParIndent
; _LOImpress_ShapeParSpacing
; _LOImpress_ShapeParTabStopCreate
; _LOImpress_ShapeParTabStopDelete
; _LOImpress_ShapeParTabStopMod
; _LOImpress_ShapeParTabStopsGetList
; _LOImpress_ShapePosition
; _LOImpress_ShapePresStyleAreaColor
; _LOImpress_ShapePresStyleAreaFillStyle
; _LOImpress_ShapePresStyleAreaGradient
; _LOImpress_ShapePresStyleAreaGradientMulticolor
; _LOImpress_ShapePresStyleAreaShadow
; _LOImpress_ShapePresStyleAreaTransparency
; _LOImpress_ShapePresStyleAreaTransparencyGradient
; _LOImpress_ShapePresStyleAreaTransparencyGradientMulti
; _LOImpress_ShapePresStyleCharEffect
; _LOImpress_ShapePresStyleCharFont
; _LOImpress_ShapePresStyleCharFontColor
; _LOImpress_ShapePresStyleCharOverLine
; _LOImpress_ShapePresStyleCharStrikeOut
; _LOImpress_ShapePresStyleCharUnderLine
; _LOImpress_ShapePresStyleGetObjByName
; _LOImpress_ShapePresStyleLineArrowStyles
; _LOImpress_ShapePresStyleLineProperties
; _LOImpress_ShapePresStyleNumCustomize
; _LOImpress_ShapePresStyleParAlignment
; _LOImpress_ShapePresStyleParIndent
; _LOImpress_ShapePresStyleParSpacing
; _LOImpress_ShapePresStyleParTabStopCreate
; _LOImpress_ShapePresStyleParTabStopDelete
; _LOImpress_ShapePresStyleParTabStopMod
; _LOImpress_ShapePresStyleParTabStopsGetList
; _LOImpress_ShapePresStylesGetNames
; _LOImpress_ShapePresStyleTextAttrFit
; _LOImpress_ShapePresStyleTextAttrSettings
; _LOImpress_ShapeRotateSlant
; _LOImpress_ShapesGetList
; _LOImpress_ShapeSize
; _LOImpress_ShapeStyleAreaColor
; _LOImpress_ShapeStyleAreaFillStyle
; _LOImpress_ShapeStyleAreaGradient
; _LOImpress_ShapeStyleAreaGradientMulticolor
; _LOImpress_ShapeStyleAreaShadow
; _LOImpress_ShapeStyleAreaTransparency
; _LOImpress_ShapeStyleAreaTransparencyGradient
; _LOImpress_ShapeStyleAreaTransparencyGradientMulti
; _LOImpress_ShapeStyleCharEffect
; _LOImpress_ShapeStyleCharFont
; _LOImpress_ShapeStyleCharFontColor
; _LOImpress_ShapeStyleCharOverLine
; _LOImpress_ShapeStyleCharStrikeOut
; _LOImpress_ShapeStyleCharUnderLine
; _LOImpress_ShapeStyleConnectorSettings
; _LOImpress_ShapeStyleCreate
; _LOImpress_ShapeStyleCurrent
; _LOImpress_ShapeStyleDelete
; _LOImpress_ShapeStyleDimensionSettings
; _LOImpress_ShapeStyleExists
; _LOImpress_ShapeStyleGetObjByName
; _LOImpress_ShapeStyleLineArrowStyles
; _LOImpress_ShapeStyleLineProperties
; _LOImpress_ShapeStyleOrganizer
; _LOImpress_ShapeStyleParAlignment
; _LOImpress_ShapeStyleParIndent
; _LOImpress_ShapeStyleParSpacing
; _LOImpress_ShapeStyleParTabStopCreate
; _LOImpress_ShapeStyleParTabStopDelete
; _LOImpress_ShapeStyleParTabStopMod
; _LOImpress_ShapeStyleParTabStopsGetList
; _LOImpress_ShapeStylesGetNames
; _LOImpress_ShapeStyleTextAttrAnimation
; _LOImpress_ShapeStyleTextAttrFit
; _LOImpress_ShapeStyleTextAttrSettings
; _LOImpress_ShapeTextAttrAnimation
; _LOImpress_ShapeTextAttrColumns
; _LOImpress_ShapeTextAttrFit
; _LOImpress_ShapeTextAttrSettings
; _LOImpress_ShapeTextBoxInsert
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaColor
; Description ...: Set or Retrieve the Fill color settings for a Shape.
; Syntax ........: _LOImpress_ShapeAreaColor(ByRef $oShape[, $iColor = Null])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iColor              - [optional] (-2-16777215) Default is Null. The Fill color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None", or $LOI_SHAPE_COLOR_USE_SLIDE_BACKGROUND (-2) to use the Slide's background color (L.O. 7.5 +).
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current Fill color as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $iColor not an Integer, less than -2 or greater than 16777215.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current color value.
;                  @Error: 3, @Extended: 2 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iColor
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version less than 7.5.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
;                  So far, I have found that Textboxes and all drawing shapes support the $LOI_SHAPE_COLOR_USE_SLIDE_BACKGROUND flag. Images and Tables do not, and will throw a property setting error.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaColor(ByRef $oShape, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency, $iCurColor

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; If $iColor is Null, and Fill Style is set to solid, then return current color value, else return LO_COLOR_OFF.
	If __LO_VarsAreNull($iColor) Then
		If ($oShape.FillStyle() = $LOI_AREA_FILL_STYLE_SOLID) Then ; If FillStyle is set to solid, then return current color value, else return $LO_COLOR_OFF (Probably a Gradient is used or otherwise).
			$iCurColor = __LOImpress_ColorRemoveAlpha($oShape.FillColor())
			If Not IsInt($iCurColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		ElseIf ($oShape.FillStyle() = $LOI_AREA_FILL_STYLE_OFF) And $oShape.PropertySetInfo.hasPropertyByName("FillUseSlideBackground") And $oShape.FillUseSlideBackground() Then
			$iCurColor = $LOI_SHAPE_COLOR_USE_SLIDE_BACKGROUND

		Else
			$iCurColor = $LO_COLOR_OFF
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurColor)
	EndIf

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE, "", $LOI_SHAPE_COLOR_USE_SLIDE_BACKGROUND) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($iColor = $LO_COLOR_OFF) Then
		$oShape.FillStyle = $LOI_AREA_FILL_STYLE_OFF
		$oShape.FillUseSlideBackground = False

	ElseIf ($iColor = $LOI_SHAPE_COLOR_USE_SLIDE_BACKGROUND) Then
		If Not __LO_VersionCheck(7.5) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		If ($oShape.PropertySetInfo.hasPropertyByName("FillUseSlideBackground")) Then
			$oShape.FillStyle = $LOI_AREA_FILL_STYLE_OFF
			$oShape.FillUseSlideBackground = True
			$iError = ($oShape.FillUseSlideBackground() = True) ? ($iError) : (BitOR($iError, 1))

		Else
			$iError = BitOR($iError, 1)
		EndIf

	Else
		$iOldTransparency = $oShape.FillTransparence()
		If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oShape.FillStyle = $LOI_AREA_FILL_STYLE_SOLID
		$oShape.FillUseSlideBackground = False
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
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOI_AREA_FILL_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Fill Style.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
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
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $sGradientName       - [optional] Default is Null. A Preset Gradient Name. See remarks. See constants, $LOI_GRAD_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iType               - [optional] (-1-5) Default is Null. The gradient type to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iIncrement          - [optional] (0, 3-256) Default is Null. The number of steps of color change. 0 = Automatic.
;                  $iXCenter            - [optional] (0-100) Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] (0-100) Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] (0-359) Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] (0-100) Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] (0-16777215) Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] (0-16777215) Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] (0-100) Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] (0-100) Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. Gradient has been successfully turned off.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $sGradientName not a String.
;                  @Error: 1, @Extended: 3 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error: 1, @Extended: 5 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 6 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 8 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 9 = $iFromColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 10 = $iToColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 11 = $iFromIntense not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 12 = $iToIntense not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillGradient" Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Parent Document Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 4 = Error creating Gradient Name.
;                  @Error: 3, @Extended: 5 = Error setting Gradient Name.
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
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
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

	If ($oShape.FillGradientName() = "") Or __LOImpress_GradientIsModified($tStyleGradient, $oShape.FillGradientName()) Then
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
; Description ...: Set or Retrieve a Shape's Multicolor Gradient settings.
; Syntax ........: _LOImpress_ShapeAreaGradientMulticolor(ByRef $oShape[, $avColorStops = Null])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $avColorStops        - [optional] Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: ?, Return: Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $avColorStops not an Array, or does not contain two columns.
;                  @Error: 1, @Extended: 3 = $avColorStops contains less than two rows.
;                  @Error: 1, @Extended: 4 = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error: 1, @Extended: 5 = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve FillGradient Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Failed to retrieve StopColor Struct.
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
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......: _LO_GradientMulticolorAdd, _LO_GradientMulticolorDelete, _LO_GradientMulticolorModify, _LOImpress_ShapeAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaGradientMulticolor(ByRef $oShape, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaGradientMulticolor($oShape, $avColorStops)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaShadow
; Description ...: Set or Retrieve the shadow settings for a Shape.
; Syntax ........: _LOImpress_ShapeAreaShadow(ByRef $oShape[, $bShadow = Null[, $iLocation = Null[, $iColor = Null[, $iDistance = Null[, $iBlur = Null[, $iTransparency = Null]]]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $bShadow             - [optional] Default is Null. If True, a Shadow is present for the Shape.
;                  $iLocation           - [optional] (0-8) Default is Null. The Location of the Shadow, must be one of the Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Shadow color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iDistance           - [optional] Default is Null. The distance of the Shadow from the Shape's edges, set in Hundredths of a Millimeter (HMM).
;                  $iBlur               - [optional] (0-150) Default is Null. The amount of blur applied to the Shadow, set in Printer's Points.
;                  $iTransparency       - [optional] (0-100) Default is Null. The percentage of Shadow transparency. 100% means completely transparent.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $bShadow not a Boolean.
;                  @Error: 1, @Extended: 3 = $iLocation not an Integer, less than 0 or greater than 8. See Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iDistance not an Integer, or less than 0.
;                  @Error: 1, @Extended: 6 = $iBlur not an Integer, less than 0 or greater than 150 Printer's Points.
;                  @Error: 1, @Extended: 7 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Distance and Location Values.
;                  @Error: 3, @Extended: 2 = Failed to modify Location property.
;                  @Error: 3, @Extended: 3 = Failed to modify Distance property.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bShadow
;                  |                               2 = Error setting $iLocation
;                  |                               4 = Error setting $iColor
;                  |                               8 = Error setting $iDistance
;                  |                               16 = Error setting $iBlur
;                  |                               32 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  LibreOffice may change the shadow distance +/- a Hundredth of a Millimeter (HMM).
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaShadow(ByRef $oShape, $bShadow = Null, $iLocation = Null, $iColor = Null, $iDistance = Null, $iBlur = Null, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaShadow($oShape, $bShadow, $iLocation, $iColor, $iDistance, $iBlur, $iTransparency)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeAreaShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaTransparency
; Description ...: Set or retrieve Transparency settings for a Shape.
; Syntax ........: _LOImpress_ShapeAreaTransparency(ByRef $oShape[, $iTransparency = Null])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iTransparency       - [optional] (0-100) Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Transparency value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaTransparency(ByRef $oShape, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaTransparency($oShape, $iTransparency)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeAreaTransparencyGradient
; Description ...: Set or retrieve the Shape transparency gradient settings.
; Syntax ........: _LOImpress_ShapeAreaTransparencyGradient(ByRef $oShape[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iType               - [optional] (-1-5) Default is Null. The type of transparency gradient that you want to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3. Call with $LOI_GRAD_TYPE_OFF to turn Transparency Gradient off.
;                  $iXCenter            - [optional] (0-100) Default is Null. The horizontal offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] (0-100) Default is Null. The vertical offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] (0-359) Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] (0-100) Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iStart              - [optional] (0-100) Default is Null. The transparency value for the beginning point of the gradient, where 0% is fully opaque and 100% is fully transparent.
;                  $iEnd                - [optional] (0-100) Default is Null. The transparency value for the endpoint of the gradient, where 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. Transparency Gradient has been successfully turned off.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $iType not an Integer, less than -1 or greater than 5. See constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 4 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 5 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 6 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iEnd not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillTransparenceGradient" Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Parent Document Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 4 = Error creating Transparency Gradient Name.
;                  @Error: 3, @Extended: 5 = Error setting Transparency Gradient Name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iType
;                  |                               2 = Error setting $iXCenter
;                  |                               4 = Error setting $iYCenter
;                  |                               8 = Error setting $iAngle
;                  |                               16 = Error setting $iTransitionStart
;                  |                               32 = Error setting $iStart
;                  |                               64 = Error setting $iEnd
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
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
; Description ...: Set or Retrieve a Shape's Multi Transparency Gradient settings.
; Syntax ........: _LOImpress_ShapeAreaTransparencyGradientMulti(ByRef $oShape[, $avColorStops = Null])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $avColorStops        - [optional] Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: ?, Return: Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $avColorStops not an Array, or does not contain two columns.
;                  @Error: 1, @Extended: 3 = $avColorStops contains less than two rows.
;                  @Error: 1, @Extended: 4 = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error: 1, @Extended: 5 = ColorStop Transparency value not an Integer, less than 0 or greater than 100. Returning problem element index.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve FillTransparenceGradient Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Failed to retrieve StopColor Struct.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current version less than 7.6.
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
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeAreaTransparencyGradientMulti(ByRef $oShape, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaTransparencyGradientMulti($oShape, $avColorStops)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCharEffect
; Description ...: Set or Retrieve the Font Effect settings for a Shape.
; Syntax ........: _LOImpress_ShapeCharEffect(ByRef $oShape[, $iCase = Null[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iCase               - [optional] (0-4) Default is Null. The Character Case Style. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRelief             - [optional] (0-2) Default is Null. The Character Relief style. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOutline            - [optional] Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeCharEffect(ByRef $oShape, $iCase = Null, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharEffect($oShape, $iCase, $iRelief, $bOutline, $bShadow)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeCharEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCharFont
; Description ...: Set and Retrieve the Font Settings for a Shape.
; Syntax ........: _LOImpress_ShapeCharFont(ByRef $oShape[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $sFontName           - [optional] Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] Default is Null. The new Font size.
;                  $iPosture            - [optional] (0-5) Default is Null. The Font Italic setting. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] (0, 50-200) Default is Null. The Font Bold settings see Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeCharFont(ByRef $oShape, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharFont($oShape, $sFontName, $nFontSize, $iPosture, $iWeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeCharFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCharFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting values for a Shape.
; Syntax ........: _LOImpress_ShapeCharFontColor(ByRef $oShape[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iFontColor          - [optional] (-1-16777215) Default is Null. The font Color value, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - [optional] (0-100) Default is Null. Transparency percentage. 0 is visible, 100 is invisible. Available for LibreOffice 7.0 and up.
;                  $iHighlight          - [optional] (-1-16777215) Default is Null. The highlight Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters. If The current LibreOffice version is below 7.0 the $iTransparency parameter will return a Null value.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeCharFontColor(ByRef $oShape, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharFontColor($oShape, $iFontColor, $iTransparency, $iHighlight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeCharFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCharOverLine
; Description ...: Set and retrieve the OverLine settings for a Shape.
; Syntax ........: _LOImpress_ShapeCharOverLine(ByRef $oShape[, $iOverLineStyle = Null[, $iOLColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iOverLineStyle      - [optional] (0-18) Default is Null. The style of the Overline line, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iOLColor            - [optional] (-1-16777215) Default is Null. The Overline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not Overlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeCharOverLine(ByRef $oShape, $iOverLineStyle = Null, $iOLColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharOverLine($oShape, $iOverLineStyle, $iOLColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeCharOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size for a Shape.
; Syntax ........: _LOImpress_ShapeCharPosition(ByRef $oShape[, $iSuperScript = Null[, $iSubScript = Null[, $iRelativeSize = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iSuperScript        - [optional] (-1-100) Default is Null. The Superscript percentage value. Call with -1 for Automatic SuperScript. See Remarks.
;                  $iSubScript          - [optional] (-1-100) Default is Null. Subscript percentage value. Call with -1 for Automatic SubScript. See Remarks.
;                  $iRelativeSize       - [optional] (1-100) Default is Null. The size percentage relative to current font size.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $oShape does not support Character properties.
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
Func _LOImpress_ShapeCharPosition(ByRef $oShape, $iSuperScript = Null, $iSubScript = Null, $iRelativeSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharPosition($oShape, $iSuperScript, $iSubScript, $iRelativeSize)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeCharPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCharScaling
; Description ...: Set or retrieve the character Scale settings for a Shape.
; Syntax ........: _LOImpress_ShapeCharScaling(ByRef $oShape[, $iScaleWidth = Null])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iScaleWidth         - [optional] (1-100) Default is Null. The percentage to horizontally stretch or compress the text. 100 is normal sizing.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current Scale Width value as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeCharScaling(ByRef $oShape, $iScaleWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharScaling($oShape, $iScaleWidth)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeCharScaling

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning) for a Shape.
; Syntax ........: _LOImpress_ShapeCharSpacing(ByRef $oShape[, $bAutoKerning = Null[, $nKerning = Null]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $bAutoKerning        - [optional] Default is Null. If True, applies a spacing in between certain pairs of characters.
;                  $nKerning            - [optional] (-928.8-928.8) Default is Null. The kerning value of the characters. See Remarks. Values are in Printer's Points as set in the LibreOffice UI.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeCharSpacing(ByRef $oShape, $bAutoKerning = Null, $nKerning = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharSpacing($oShape, $bAutoKerning, $nKerning)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeCharSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCharStrikeOut
; Description ...: Set or Retrieve the Strikeout settings for a Shape.
; Syntax ........: _LOImpress_ShapeCharStrikeOut(ByRef $oShape[, $iStrikeLineStyle = Null[, $bWordOnly = Null]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iStrikeLineStyle    - [optional] (0-6) Default is Null. The Strikeout Line Style, see constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bWordOnly           - [optional] Default is Null. If True, strike out is applied to words only, skipping whitespaces.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeCharStrikeOut(ByRef $oShape, $iStrikeLineStyle = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharStrikeOut($oShape, $iStrikeLineStyle, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeCharStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCharUnderLine
; Description ...: Set and retrieve the Underline settings for a Shape.
; Syntax ........: _LOImpress_ShapeCharUnderLine(ByRef $oShape[, $iUnderLineStyle = Null[, $iULColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iUnderLineStyle     - [optional] (0-18) Default is Null. The Underline line style, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iULColor            - [optional] (-1-16777215) Default is Null. The underline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not underlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape an Object.
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
Func _LOImpress_ShapeCharUnderLine(ByRef $oShape, $iUnderLineStyle = Null, $iULColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharUnderLine($oShape, $iUnderLineStyle, $iULColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeCharUnderLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeCreateTextCursor
; Description ...: Create a Text Cursor in a Shape's Textbox for inserting text etc.
; Syntax ........: _LOImpress_ShapeCreateTextCursor(ByRef $oShape)
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
; Return values .: Success: Object.
;                  @Error: 0, @Extended: 0, Return: Object = Success. A Text Cursor Object located in the Textbox.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oShape not a shape supporting a Textbox.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a Text Cursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
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
; Name ..........: _LOImpress_ShapeDelete
; Description ...: Delete a Shape.
; Syntax ........: _LOImpress_ShapeDelete(ByRef $oShape)
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Shape was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Shape's containing Slide.
;                  @Error: 3, @Extended: 2 = Failed to retrieve count of shapes.
;                  @Error: 3, @Extended: 3 = Same number of shapes still present. Failed to delete the Shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeDelete(ByRef $oShape)
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
	If ($oDrawPage.getCount() = $iShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Count of shapes the same, shape wasn't deleted.

	$oShape = Null

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_ShapeDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeExists
; Description ...: Check if a Document contains a DrawShape with the specified name.
; Syntax ........: _LOImpress_ShapeExists(ByRef $oDoc, $sShapeName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sShapeName          - The Shape name to search for.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. If a Shape was found matching $sShapeName, True is returned, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sShapeName not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Slide Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Shape Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: For all shapes that have not been renamed by the user, the name value is blank, even though the shape in the UI has a name. Therefore this function will only work for user-renamed shapes.
;                  This function searches all slides, because a Shape name must be unique for an entire slideshow document.
;                  This function will work for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeExists(ByRef $oDoc, $sShapeName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSlide, $oShape

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	For $i = 0 To $oDoc.DrawPages.getCount() - 1
		$oSlide = $oDoc.DrawPages.getByIndex($i)
		If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If $oSlide.hasElements() Then
			For $j = 0 To $oSlide.getCount() - 1
				$oShape = $oSlide.getByIndex($j)
				If Not IsObj($oShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				If ($oShape.Name() <> "") And ($oShape.Name() = $sShapeName) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

				Sleep((IsInt($j / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
			Next
		EndIf

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; No matches
EndFunc   ;==>_LOImpress_ShapeExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeImageAltText
; Description ...: Set or Retrieve Image Alternate text settings.
; Syntax ........: _LOImpress_ShapeImageAltText(ByRef $oImage[, $sText = Null[, $sAltText = Null[, $bDecorative = Null]]])
; Parameters ....: $oImage              - A Image object returned by a previous _LOImpress_ShapeImageInsert, or _LOImpress_ShapesGetList function.
;                  $sText               - [optional] Default is Null. Enter alternative text to display when the image isn't available.
;                  $sAltText            - [optional] Default is Null. Detailed alternative text of the Image.
;                  $bDecorative         - [optional] Default is Null. If True, the image is considered decorative and is ignored by assistive readers. L.O. 7.6+.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters. If The current LibreOffice version is below 7.6 the $bDecorative parameter will return a Null value.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oImage not an Object.
;                  @Error: 1, @Extended: 2 = Shape called in $oImage not an Image shape.
;                  @Error: 1, @Extended: 3 = $sText not a string.
;                  @Error: 1, @Extended: 4 = $sAltText not a string.
;                  @Error: 1, @Extended: 5 = $bDecorative not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sText
;                  |                               2 = Error setting $sAltText
;                  |                               4 = Error setting $bDecorative
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version less than 7.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeImageAltText(ByRef $oImage, $sText = Null, $sAltText = Null, $bDecorative = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asName[3]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oImage.supportsService("com.sun.star.drawing.GraphicObjectShape") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sText, $sAltText, $bDecorative) Then
		If __LO_VersionCheck(7.6) Then
			__LO_ArrayFill($asName, $oImage.Title(), $oImage.Description(), $oImage.Decorative())

		Else
			__LO_ArrayFill($asName, $oImage.Title(), $oImage.Description(), Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $asName)
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.Title = $sText
		$iError = ($oImage.Title() = $sText) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sAltText <> Null) Then
		If Not IsString($sAltText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.Description = $sAltText
		$iError = ($oImage.Description() = $sAltText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bDecorative <> Null) Then
		If Not IsBool($bDecorative) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oImage.Decorative = $bDecorative
		$iError = ($oImage.Decorative() = $bDecorative) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeImageAltText

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeImageCrop
; Description ...: Set or retrieve Image crop settings.
; Syntax ........: _LOImpress_ShapeImageCrop(ByRef $oImage[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null[, $bKeepScale = Null]]]]])
; Parameters ....: $oImage              - A Image object returned by a previous _LOImpress_ShapeImageInsert, or _LOImpress_ShapesGetList function.
;                  $iLeft               - [optional] Default is Null. The amount in Hundredths of a Millimeter (HMM) to either extend the background of the image, (negative numbers), or to crop, (positive numbers) from the Left side.
;                  $iRight              - [optional] Default is Null. The amount in Hundredths of a Millimeter (HMM) to either extend the background of the image, (negative numbers), or to crop, (positive numbers) from the Right side.
;                  $iTop                - [optional] Default is Null. The amount in Hundredths of a Millimeter (HMM) to either extend the background of the image, (negative numbers), or to crop, (positive numbers) from the Top side.
;                  $iBottom             - [optional] Default is Null. The amount in Hundredths of a Millimeter (HMM) to either extend the background of the image, (negative numbers), or to crop, (positive numbers) from the Bottom side.
;                  $bKeepScale          - [optional] Default is Null. If True, crop amounts are removed or added to the image, while keeping the scaling. If False, crop values are removed or added while retaining the image size. See remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oImage not an Object.
;                  @Error: 1, @Extended: 2 = Shape called in $oImage not an Image shape.
;                  @Error: 1, @Extended: 3 = $bKeepScale not a Boolean.
;                  @Error: 1, @Extended: 4 = $iLeft not an Integer.
;                  @Error: 1, @Extended: 5 = $iRight not an Integer.
;                  @Error: 1, @Extended: 6 = $iTop not an Integer.
;                  @Error: 1, @Extended: 7 = $iBottom not an Integer.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve the image Crop structure.
;                  @Error: 3, @Extended: 2 = Failed to retrieve the image Size structure.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
; Author ........: donnyh13
; Modified ......:
; Remarks .......: There is no setting for $bKeepScale in LibreOffice's API. Therefore I have made this function behave as follows:
;                  - Unless $bKeepScale is called with False, $bKeepScale is assumed to be True.
;                  - Calling $bKeepScale alone, without setting a crop value does nothing.
;                  - The return value of $bKeepScale is always Null.
;                  Maximum crop values are based on page width. You cannot exceed the size of the page, nor crop too much of the image away.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOImpress_ShapeImageInsert, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeImageCrop(ByRef $oImage, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null, $bKeepScale = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avImage[5]
	Local $tCrop, $tSize

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oImage.supportsService("com.sun.star.drawing.GraphicObjectShape") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tCrop = $oImage.GraphicCrop()
	If Not IsObj($tCrop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tSize = $oImage.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iLeft, $iRight, $iTop, $iBottom, $bKeepScale) Then
		__LO_ArrayFill($avImage, $tCrop.Left(), $tCrop.Right(), $tCrop.Top(), $tCrop.Bottom(), Null)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avImage)
	EndIf

	If ($bKeepScale = Null) Then $bKeepScale = True

	If Not IsBool($bKeepScale) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($iLeft <> Null) Then
		If Not IsInt($iLeft) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If $bKeepScale Then $tSize.Width = ($tSize.Width() + $tCrop.Left() - $iLeft)
		$tCrop.Left = $iLeft
	EndIf

	If ($iRight <> Null) Then
		If Not IsInt($iRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		If $bKeepScale Then $tSize.Width = ($tSize.Width() + $tCrop.Right() - $iRight)
		$tCrop.Right = $iRight
	EndIf

	If ($iTop <> Null) Then
		If Not IsInt($iTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		If $bKeepScale Then $tSize.Height = ($tSize.Height() + $tCrop.Top() - $iTop)
		$tCrop.Top = $iTop
	EndIf

	If ($iBottom <> Null) Then
		If Not IsInt($iBottom) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		If $bKeepScale Then $tSize.Height = ($tSize.Height() + $tCrop.Bottom() - $iBottom)
		$tCrop.Bottom = $iBottom
	EndIf

	$oImage.GraphicCrop = $tCrop

	If $bKeepScale Then $oImage.Size = $tSize

	; Error checking
	$iError = (__LO_VarsAreNull($iLeft)) ? ($iError) : ((__LO_IntIsBetween($oImage.GraphicCrop.Left(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iRight)) ? ($iError) : ((__LO_IntIsBetween($oImage.GraphicCrop.Right(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iTop)) ? ($iError) : ((__LO_IntIsBetween($oImage.GraphicCrop.Top(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iBottom)) ? ($iError) : ((__LO_IntIsBetween($oImage.GraphicCrop.Bottom(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeImageCrop

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeImageInsert
; Description ...: Insert an image into a slide.
; Syntax ........: _LOImpress_ShapeImageInsert(ByRef $oSlide, $sURL[, $iWidth = -1[, $iHeight = -1[, $iX = -1[, $iY = -1]]]])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $sURL                - The file path to the image to insert.
;                  $iWidth              - [optional] Default is -1. The Images's Width in Hundredths of a Millimeter (HMM). Call with -1 for automatic width.
;                  $iHeight             - [optional] Default is -1. The Images's Height in Hundredths of a Millimeter (HMM). Call with -1 for automatic height.
;                  $iX                  - [optional] Default is -1. The X position from the top-left of the page, in Hundredths of a Millimeter (HMM). Call with -1 to center the image horizontally.
;                  $iY                  - [optional] Default is -1. The Y position from the top-left of the page, in Hundredths of a Millimeter (HMM). Call with -1 to center the image vertically.
; Return values .: Success: Object.
;                  @Error: 0, @Extended: 0, Return: Object = Success. Image was successfully inserted, returning image Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $sImage not a String.
;                  @Error: 1, @Extended: 3 = Image called in $sImage doesn't exist at given path.
;                  @Error: 1, @Extended: 4 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 5 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 6 = $iX not an Integer.
;                  @Error: 1, @Extended: 7 = $iY not an Integer.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failure creating "com.sun.star.drawing.GraphicObjectShape" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error converting Image Path to LibreOffice URL.
;                  @Error: 3, @Extended: 2 = Error retrieving Document Object.
;                  @Error: 3, @Extended: 3 = Error retrieving image's size structure.
;                  @Error: 3, @Extended: 4 = Error retrieving Bitmap size.
;                  @Error: 3, @Extended: 5 = Error calculating image's ratio.
;                  @Error: 3, @Extended: 6 = Error calculating Slide's ratio.
;                  @Error: 3, @Extended: 7 = Error retrieving image's Position structure.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Image is Auto-Sized and centered thanks to method by A. Pitonyak, OOME 4.1, PDF pg 320.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeImageInsert(ByRef $oSlide, $sURL, $iWidth = -1, $iHeight = -1, $iX = -1, $iY = -1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oImage, $oDoc
	Local $tBitmapSize, $tNewSize, $tPos
	Local $nImageRatio, $nPageRatio

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not FileExists($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$sURL = _LO_PathConvert($sURL, $LO_PATHCONV_OFFICE_RETURN)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oImage = $oDoc.createInstance("com.sun.star.drawing.GraphicObjectShape")
	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oImage.GraphicURL = $sURL

	$oSlide.add($oImage)

	$tNewSize = $oImage.Size()
	If Not IsObj($tNewSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	; Auto-Size and center image. Thanks to method by A. Pitonyak, OOME 4.1, PDF pg 320.
	$tBitmapSize = $oImage.GraphicObjectFillBitmap.GetSize()
	If Not IsObj($tBitmapSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$nImageRatio = (Number($tBitmapSize.Height()) / Number($tBitmapSize.Width()))
	If Not IsNumber($nImageRatio) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$nPageRatio = (Number($oSlide.Height()) / Number($oSlide.Width()))
	If Not IsNumber($nPageRatio) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	If ($nPageRatio > $nImageRatio) Then ; Compare the ratios to see which is wider.
		$tNewSize.Width = ($iWidth = -1) ? ($oSlide.Width()) : ($iWidth)
		$tNewSize.Height = ($iHeight = -1) ? (($iWidth = -1) ? (Int($oSlide.Width() * $nImageRatio)) : ($iWidth * $nImageRatio)) : ($iHeight) ;

	Else
		$tNewSize.Width = ($iWidth = -1) ? (Int($oSlide.Width() / $nImageRatio)) : ($iWidth)
		$tNewSize.Height = ($iHeight = -1) ? ($oSlide.Height()) : ($iHeight)
	EndIf

	$tPos = $oImage.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

	$tPos.X = ($iX = -1) ? (Int(($oSlide.Width() - $tNewSize.Width()) / 2)) : ($iX)
	$tPos.Y = ($iY = -1) ? (Int(($oSlide.Height() - $tNewSize.Height()) / 2)) : ($iY)

	$oImage.Size = $tNewSize
	$oImage.Position = $tPos

	Return SetError($__LO_STATUS_SUCCESS, 0, $oImage)
EndFunc   ;==>_LOImpress_ShapeImageInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeImageModify
; Description ...: Set or retrieve Image modification settings.
; Syntax ........: _LOImpress_ShapeImageModify(ByRef $oImage[, $bFlipVert = Null[, $bFlipHori = Null]])
; Parameters ....: $oImage              - A Image object returned by a previous _LOImpress_ShapeImageInsert, or _LOImpress_ShapesGetList function.
;                  $bFlipVert           - [optional] Default is Null. If True, the image is flipped vertically.
;                  $bFlipHori           - [optional] Default is Null. If True, the image is flipped horizontally.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oImage not an Object.
;                  @Error: 1, @Extended: 2 = Shape called in $oImage not an Image shape.
;                  @Error: 1, @Extended: 3 = $bFlipVert not a Boolean.
;                  @Error: 1, @Extended: 4 = $bFlipHori not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFlipVert
;                  |                               2 = Error setting $bFlipHori
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOImpress_ShapeImageInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeImageModify(ByRef $oImage, $bFlipVert = Null, $bFlipHori = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avImage[2]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oImage.supportsService("com.sun.star.drawing.GraphicObjectShape") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bFlipVert, $bFlipHori) Then
		__LO_ArrayFill($avImage, ($oImage.RotateAngle() = 18000) ? (True) : (False), $oImage.IsMirrored()) ; If image is rotated 18000 degrees, image is flipped vertically.

		Return SetError($__LO_STATUS_SUCCESS, 1, $avImage)
	EndIf

	If ($bFlipVert <> Null) Then
		If Not IsBool($bFlipVert) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If $bFlipVert Then
			$oImage.RotateAngle = 18000 ; Image is vertically flipped.
			$iError = ($oImage.RotateAngle() = 18000) ? ($iError) : (BitOR($iError, 1))

		Else
			$oImage.RotateAngle = 0
			$iError = ($oImage.RotateAngle() = 0) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($bFlipHori <> Null) Then
		If Not IsBool($bFlipHori) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.IsMirrored = $bFlipHori
		$iError = ($oImage.IsMirrored() = $bFlipHori) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeImageModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeImageReplace
; Description ...: Replace an image with another image.
; Syntax ........: _LOImpress_ShapeImageReplace(ByRef $oImage, $sNewImage)
; Parameters ....: $oImage              - A Image object returned by a previous _LOImpress_ShapeImageInsert, or _LOImpress_ShapesGetList function.
;                  $sNewImage           - The file path to the new image.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Image was successfully replaced.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oImage not an Object.
;                  @Error: 1, @Extended: 2 = Shape called in $oImage not an Image shape.
;                  @Error: 1, @Extended: 3 = $sNewImage not a string.
;                  @Error: 1, @Extended: 4 = File called in $sNewImage doesn't exist.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to convert $sNewImage Path to LibreOffice URL.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_ShapeImageInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeImageReplace(ByRef $oImage, $sNewImage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oImage.supportsService("com.sun.star.drawing.GraphicObjectShape") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sNewImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not FileExists($sNewImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$sNewImage = _LO_PathConvert($sNewImage, $LO_PATHCONV_OFFICE_RETURN)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oImage.GraphicURL = $sNewImage

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_ShapeImageReplace

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeInteraction
; Description ...: Set or Retrieve a Shape's current Interaction settings.
; Syntax ........: _LOImpress_ShapeInteraction(ByRef $oShape[, $iAction = Null[, $sTarget = Null[, $iVerb = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iAction             - [optional] (0-13) Default is Null. The action to perform when the shape is clicked. See Constants, $LOI_SHAPE_INTERACTION_ACTION_* as defined in LibreOfficeImpress_Constants.au3.
;                  $sTarget             - [optional] Default is Null. The target for the action. See remarks.
;                  $iVerb               - [optional] Default is Null. If $iAction is set to $LOI_SHAPE_INTERACTION_ACTION_OBJ_ACTION, this is the action to perform on the OLE Object. See remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $iAction not an Integer, less than 0 or greater than 13. See Constants, $LOI_SHAPE_INTERACTION_ACTION_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $sTarget not a String.
;                  @Error: 1, @Extended: 4 = Slide or shape does not exist with name called in $sTarget.
;                  @Error: 1, @Extended: 5 = File called in $sTarget does not exist.
;                  @Error: 1, @Extended: 6 = $iVerb not an Integer.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Target value.
;                  @Error: 3, @Extended: 2 = Failed to convert current Target path.
;                  @Error: 3, @Extended: 3 = Failed to retrieve parent Slide Object.
;                  @Error: 3, @Extended: 4 = Failed to retrieve parent Document Object.
;                  @Error: 3, @Extended: 5 = Failed to convert target path.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iAction
;                  |                               2 = Error setting $sTarget
;                  |                               4 = Error setting $iVerb
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  When $iAction is set to $LOI_SHAPE_INTERACTION_ACTION_OBJ_ACTION, call $sTarget with the appropriate flag as a string for the action to perform on the OLE Object, and call $iVerb with the appropriate flag as an integer.
;                  As an example for values to use with $LOI_SHAPE_INTERACTION_ACTION_OBJ_ACTION, I have observed the following values:
;                  - When setting the action to "edit", $sTarget has a value of "-1" (as a string), and $iVerb has a value of 65535.
;                  - When setting the action to "Save a Copy As", $sTarget has a value of "-8" (as a string), and $iVerb has a value of 65528.
;                  $iVerb determines the action performed, and $sTarget determines the action showing selected in the UI.
;                  User is responsible for ensuring values are correctly called (i.e. that a shape, or slide etc exists by that name) for $LOI_SHAPE_INTERACTION_ACTION_GOTO_PAGE_OBJ, $LOI_SHAPE_INTERACTION_ACTION_OBJ_ACTION, and $LOI_SHAPE_INTERACTION_ACTION_MACRO.
;                  See comments for each $LOI_SHAPE_INTERACTION_ACTION_* Constant for what values are expected in $sTarget otherwise.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeInteraction(ByRef $oShape, $iAction = Null, $sTarget = Null, $iVerb = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurVal
	Local $oSlide, $oDoc
	Local $avInteraction[3]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iAction, $sTarget) Then
		$sCurVal = $oShape.Bookmark()
		If Not IsString($sCurVal) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Switch $oShape.OnClick()
			Case $LOI_SHAPE_INTERACTION_ACTION_DOCUMENT, $LOI_SHAPE_INTERACTION_ACTION_SOUND, $LOI_SHAPE_INTERACTION_ACTION_PROGRAM
				$sCurVal = _LO_PathConvert($oShape.Bookmark(), $LO_PATHCONV_PCPATH_RETURN)
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndSwitch

		__LO_ArrayFill($avInteraction, $oShape.OnClick(), $sCurVal, $oShape.Verb())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avInteraction)
	EndIf

	If ($iAction <> Null) Then
		If Not __LO_IntIsBetween($iAction, $LOI_SHAPE_INTERACTION_ACTION_NONE, $LOI_SHAPE_INTERACTION_ACTION_EXIT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oShape.OnClick = $iAction
		$iError = ($oShape.OnClick() = $iAction) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sTarget <> Null) Then
		If Not IsString($sTarget) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		Switch $oShape.OnClick()
			Case $LOI_SHAPE_INTERACTION_ACTION_GOTO_PAGE_OBJ
				$oSlide = $oShape.Parent()
				If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

				$oDoc = $oSlide.MasterPage.Forms.Parent()
				If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

				If Not _LOImpress_ShapeExists($oDoc, $sTarget) And Not _
						$oDoc.Links.getByName("Slide").Links.hasByName($sTarget) And Not _
						$oDoc.Links.getByName("Notes").Links.hasByName($sTarget) And Not _
						$oDoc.Links.getByName("Master Page").Links.hasByName($sTarget) And Not _
						$oDoc.Links.getByName("Handouts").Links.hasByName($sTarget) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Not sure if I need to check Handouts?

			Case $LOI_SHAPE_INTERACTION_ACTION_DOCUMENT, $LOI_SHAPE_INTERACTION_ACTION_SOUND, $LOI_SHAPE_INTERACTION_ACTION_PROGRAM
				If Not FileExists($sTarget) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

				$sTarget = _LO_PathConvert($sTarget, $LO_PATHCONV_OFFICE_RETURN)
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
		EndSwitch

		$oShape.Bookmark = $sTarget
		$iError = ($oShape.Bookmark() = $sTarget) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iVerb <> Null) Then
		If Not IsInt($iVerb) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oShape.Verb = $iVerb
		$iError = ($oShape.Verb() = $iVerb) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeInteraction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeLineArrowStyles
; Description ...: Set or Retrieve Shape Line Start and End Arrow Style settings.
; Syntax ........: _LOImpress_ShapeLineArrowStyles(ByRef $oShape[, $vStartStyle = Null[, $iStartWidth = Null[, $bStartCenter = Null[, $bSync = Null[, $vEndStyle = Null[, $iEndWidth = Null[, $bEndCenter = Null]]]]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $vStartStyle         - [optional] (0-32, or String) Default is Null. The Arrow head to apply to the start of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iStartWidth         - [optional] (0-5004) Default is Null. The Width of the Starting Arrowhead, in Hundredths of a Millimeter (HMM).
;                  $bStartCenter        - [optional] Default is Null. If True, Places the center of the Start arrowhead on the endpoint of the line.
;                  $bSync               - [optional] Default is Null. If True, Synchronizes the Start Arrowhead settings with the end Arrowhead settings. See remarks.
;                  $vEndStyle           - [optional] (0-32, or String) Default is Null. The Arrow head to apply to the end of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iEndWidth           - [optional] (0-5004) Default is Null. The Width of the Ending Arrowhead, in Hundredths of a Millimeter (HMM).
;                  $bEndCenter          - [optional] Default is Null. If True, Places the center of the End arrowhead on the endpoint of the line.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $vStartStyle not a String, and not an Integer.
;                  @Error: 1, @Extended: 3 = $vStartStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iStartWidth not an Integer, less than 0 or greater than 5004.
;                  @Error: 1, @Extended: 5 = $bStartCenter not a Boolean.
;                  @Error: 1, @Extended: 6 = $bSync not a Boolean.
;                  @Error: 1, @Extended: 7 = $vEndStyle not a String, and not an Integer.
;                  @Error: 1, @Extended: 8 = $vSEndStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 9 = $iEndWidth not an Integer, less than 0 or greater than 5004.
;                  @Error: 1, @Extended: 10 = $bEndCenter not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to convert Constant to Arrowhead name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $vStartStyle
;                  |                               2 = Error setting $iStartWidth
;                  |                               4 = Error setting $bStartCenter
;                  |                               8 = Error setting $bSync
;                  |                               16 = Error setting $vEndStyle
;                  |                               32 = Error setting $iEndWidth
;                  |                               64 = Error setting $bEndCenter
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
		__LO_ArrayFill($avArrow, __LOImpress_ShapeLineArrowStyleName(Null, $oShape.LineStartName()), $oShape.LineStartWidth(), $oShape.LineStartCenter(), _
				((($oShape.LineStartName() = $oShape.LineEndName()) And ($oShape.LineStartWidth() = $oShape.LineEndWidth()) And ($oShape.LineStartCenter() = $oShape.LineEndCenter())) ? (True) : (False)), _ ; See if Start and End are the same.
				__LOImpress_ShapeLineArrowStyleName(Null, $oShape.LineEndName()), $oShape.LineEndWidth(), $oShape.LineEndCenter())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avArrow)
	EndIf

	If ($vStartStyle <> Null) Then
		If Not IsString($vStartStyle) And Not IsInt($vStartStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If IsInt($vStartStyle) Then
			If Not __LO_IntIsBetween($vStartStyle, $LOI_SHAPE_LINE_ARROW_TYPE_NONE, $LOI_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$sStartStyle = __LOImpress_ShapeLineArrowStyleName($vStartStyle)
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

			$sEndStyle = __LOImpress_ShapeLineArrowStyleName($vEndStyle)
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
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $vStyle              - [optional] (0-31, or String) Default is Null. The Line Style to use. Can be a Custom Line Style name, or one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Line color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iWidth              - [optional] (0-5004) Default is Null. The line Width, set in Hundredths of a Millimeter (HMM).
;                  $iTransparency       - [optional] (0-100) Default is Null. The Line transparency percentage. 100% = fully transparent.
;                  $iCornerStyle        - [optional] (0, 2-4) Default is Null. The Line Corner Style. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  $iCapStyle           - [optional] (0-2) Default is Null. The Line Cap Style. See Constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $vStyle not a String, and not an Integer.
;                  @Error: 1, @Extended: 3 = $vStyle is an Integer, but less than 0 or greater than 31. See constants $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iWidth not an Integer, less than 0 or greater than 5004.
;                  @Error: 1, @Extended: 6 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iCornerStyle not an Integer, not equal to 0, equal to 1, not equal to 2 or greater than 4. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error: 1, @Extended: 8 = $iCapStyle is an Integer, but less than 0 or greater than 2. See constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to convert Constant to Line Style name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $vStyle
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $iWidth
;                  |                               8 = Error setting $iTransparency
;                  |                               16 = Error setting $iCornerStyle
;                  |                               32 = Error setting $iCapStyle
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $vStyle accepts a String or an Integer because there is the possibility of a custom Line Style being available that the user may want to use.
;                  When retrieving the current settings, $vStyle could be either an Integer or a String. It will be a String if the current Line Style is a custom Line Style, else an Integer, corresponding to one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
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
; Name ..........: _LOImpress_ShapeName
; Description ...: Set or Retrieve a Shape's Name.
; Syntax ........: _LOImpress_ShapeName(ByRef $oShape[, $sName = Null])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $sName               - [optional] Default is Null. The new, unique Name for the Shape.
; Return values .: Success: 1 or String
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Shape's name was successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. All optional parameters were called with Null, returning the Shape's current name.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Document already contains a Shape with the same name as called in $sName.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Shape's name.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Parent Slide Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Parent Document Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sName
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  For all shapes that have not been renamed by the user, the name value is blank, even though the shape in the UI has a name.
;                  When renaming a shape, the Shape name must be unique to the entire slideshow (at least in the LibreOffice UI), however due to the above issue, it is possible to have two shapes with the same name in the UI (and also internally if I don't make a safety check).
;                  This function will work for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeName(ByRef $oShape, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oSlide, $oDoc
	Local $sCurrName

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sName) Then
		$sCurrName = $oShape.Name()
		If Not IsString($sCurrName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrName)
	EndIf

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSlide = $oShape.Parent()
	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If _LOImpress_ShapeExists($oDoc, $sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oShape.Name = $sName
	$iError = ($oShape.Name() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeParAlignment
; Description ...: Set and Retrieve Paragraph Alignment settings for a Shape.
; Syntax ........: _LOImpress_ShapeParAlignment(ByRef $oShape[, $iHorAlign = Null[, $iLastLineAlign = Null[, $iTxtDirection = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iHorAlign           - [optional] (0-3) Default is Null. The Horizontal alignment of the paragraph. See Constants, $LOI_PAR_ALIGN_HOR_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iLastLineAlign      - [optional] (0-3) Default is Null. Specify the alignment for the last line in the paragraph. See Constants, $LOI_PAR_LAST_LINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iTxtDirection       - [optional] (0-5) Default is Null. The Text Writing Direction. See Constants, $LOI_PAR_TXT_DIR_* as defined in LibreOfficeImpress_Constants.au3. [LibreOffice Default is 4]
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeParAlignment(ByRef $oShape, $iHorAlign = Null, $iLastLineAlign = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParAlignment($oShape, $iHorAlign, $iLastLineAlign, $iTxtDirection)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeParAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeParIndent
; Description ...: Set or Retrieve Paragraph Indent settings for a Shape.
; Syntax ........: _LOImpress_ShapeParIndent(ByRef $oShape[, $iBeforeTxt = Null[, $iAfterTxt = Null[, $iFirstLine = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iBeforeTxt          - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM).
;                  $iAfterTxt           - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM)
;                  $iFirstLine          - [optional] (0-1162202) Default is Null. Indentation distance of the first line of a paragraph. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeParIndent(ByRef $oShape, $iBeforeTxt = Null, $iAfterTxt = Null, $iFirstLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParIndent($oShape, $iBeforeTxt, $iAfterTxt, $iFirstLine)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeParIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeParSpacing
; Description ...: Set and Retrieve Line Spacing settings for a Shape.
; Syntax ........: _LOImpress_ShapeParSpacing(ByRef $oShape[, $iAbovePar = Null[, $iBelowPar = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iAbovePar           - [optional] (0-100000) Default is Null. The Space above a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iBelowPar           - [optional] (0-100000) Default is Null. The Space Below a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iLineSpcMode        - [optional] (0-3) Default is Null. The line spacing type of the paragraph. See Constants, $LOI_PAR_LINE_SPC_MODE_* as defined in LibreOfficeImpress_Constants.au3, also notice min and max values for each.
;                  $iLineSpcHeight      - [optional] Default is Null. This value specifies the height in regard to Mode. See Remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeParSpacing(ByRef $oShape, $iAbovePar = Null, $iBelowPar = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParSpacing($oShape, $iAbovePar, $iBelowPar, $iLineSpcMode, $iLineSpcHeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeParSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeParTabStopCreate
; Description ...: Create a new TabStop for a Shape.
; Syntax ........: _LOImpress_ShapeParTabStopCreate(ByRef $oShape, $iPosition[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iPosition           - The TabStop position to set the new TabStop to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Settings were successfully set. New TabStop position is returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeParTabStopCreate(ByRef $oShape, $iPosition, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopCreate($oShape, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeParTabStopCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeParTabStopDelete
; Description ...: Delete a TabStop from a Shape.
; Syntax ........: _LOImpress_ShapeParTabStopDelete(ByRef $oShape, $iTabStop)
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Returning True if TabStop was successfully deleted, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeParTabStopDelete(ByRef $oShape, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopDelete($oShape, $iTabStop)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeParTabStopDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop in a Shape.
; Syntax ........: _LOImpress_ShapeParTabStopMod(ByRef $oShape, $iTabStop[, $iPosition = Null[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - [optional] Default is Null. The New position to set the input position to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: ?, Return: 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
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
Func _LOImpress_ShapeParTabStopMod(ByRef $oShape, $iTabStop, $iPosition = Null, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopMod($oShape, $iTabStop, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeParTabStopMod

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeParTabStopsGetList
; Description ...: Retrieve an array of TabStops available in a Shape.
; Syntax ........: _LOImpress_ShapeParTabStopsGetList(ByRef $oShape)
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array of TabStops. @Extended set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeParTabStopsGetList(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopsGetList($oShape)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeParTabStopsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePosition
; Description ...: Set or Retrieve the Shape's position settings.
; Syntax ........: _LOImpress_ShapePosition(ByRef $oShape[, $iX = Null[, $iY = Null[, $bProtectPos = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iX                  - [optional] Default is Null. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - [optional] Default is Null. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $bProtectPos         - [optional] Default is Null. If True, the Shape's position is locked.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $iX not an Integer.
;                  @Error: 1, @Extended: 3 = $iY not an Integer.
;                  @Error: 1, @Extended: 4 = $bProtectPos not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Shape's Position Structure.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iX
;                  |                               2 = Error setting $iY
;                  |                               4 = Error setting $bProtectPos
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
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
; Name ..........: _LOImpress_ShapePresStyleAreaColor
; Description ...: Set or Retrieve the Fill color settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleAreaColor(ByRef $oPresStyle[, $iColor = Null])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iColor              - [optional] (-1-16777215) Default is Null. The Fill color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current Fill color as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current color value.
;                  @Error: 3, @Extended: 2 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iColor
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleAreaColor(ByRef $oPresStyle, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeStyleAreaColor($oPresStyle, $iColor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOImpress_ShapePresStyleAreaFillStyle(ByRef $oPresStyle)
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOI_AREA_FILL_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Fill Style.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleAreaFillStyle(ByRef $oPresStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oPresStyle.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOImpress_ShapePresStyleAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleAreaGradient
; Description ...: Modify or retrieve the settings for Presentation Style Background color Gradient.
; Syntax ........: _LOImpress_ShapePresStyleAreaGradient(ByRef $oDoc, ByRef $oPresStyle[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $sGradientName       - [optional] Default is Null. A Preset Gradient Name. See remarks. See constants, $LOI_GRAD_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iType               - [optional] (-1-5) Default is Null. The gradient type to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iIncrement          - [optional] (0, 3-256) Default is Null. The number of steps of color change. 0 = Automatic.
;                  $iXCenter            - [optional] (0-100) Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] (0-100) Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] (0-359) Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] (0-100) Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] (0-16777215) Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] (0-16777215) Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] (0-100) Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] (0-100) Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. Gradient has been successfully turned off.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 3 = $sGradientName not a String.
;                  @Error: 1, @Extended: 4 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error: 1, @Extended: 6 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 9 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 10 = $iFromColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 11 = $iToColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 12 = $iFromIntense not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 13 = $iToIntense not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillGradient" Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Error creating Gradient Name.
;                  @Error: 3, @Extended: 4 = Error setting Gradient Name.
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
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleAreaGradient(ByRef $oDoc, ByRef $oPresStyle, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeStyleAreaGradient($oDoc, $oPresStyle, $sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleAreaGradientMulticolor
; Description ...: Set or Retrieve a Presentation Style's Multicolor Gradient settings.
; Syntax ........: _LOImpress_ShapePresStyleAreaGradientMulticolor(ByRef $oPresStyle[, $avColorStops = Null])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $avColorStops        - [optional] Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: ?, Return: Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 2 = $avColorStops not an Array, or does not contain two columns.
;                  @Error: 1, @Extended: 3 = $avColorStops contains less than two rows.
;                  @Error: 1, @Extended: 4 = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error: 1, @Extended: 5 = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve FillGradient Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Failed to retrieve StopColor Struct.
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
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LO_GradientMulticolorAdd, _LO_GradientMulticolorDelete, _LO_GradientMulticolorModify, _LOImpress_ShapeAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleAreaGradientMulticolor(ByRef $oPresStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaGradientMulticolor($oPresStyle, $avColorStops)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleAreaShadow
; Description ...: Set or Retrieve the shadow settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleAreaShadow(ByRef $oPresStyle[, $bShadow = Null[, $iLocation = Null[, $iColor = Null[, $iDistance = Null[, $iBlur = Null[, $iTransparency = Null]]]]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $bShadow             - [optional] Default is Null. If True, a Shadow is present for the Shape.
;                  $iLocation           - [optional] (0-8) Default is Null. The Location of the Shadow, must be one of the Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Shadow color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iDistance           - [optional] Default is Null. The distance of the Shadow from the Shape's edges, set in Hundredths of a Millimeter (HMM).
;                  $iBlur               - [optional] (0-150) Default is Null. The amount of blur applied to the Shadow, set in Printer's Points.
;                  $iTransparency       - [optional] (0-100) Default is Null. The percentage of Shadow transparency. 100% means completely transparent.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 2 = $bShadow not a Boolean.
;                  @Error: 1, @Extended: 3 = $iLocation not an Integer, less than 0 or greater than 8. See Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iDistance not an Integer, or less than 0.
;                  @Error: 1, @Extended: 6 = $iBlur not an Integer, less than 0 or greater than 150 Printer's Points.
;                  @Error: 1, @Extended: 7 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Distance and Location Values.
;                  @Error: 3, @Extended: 2 = Failed to modify Location property.
;                  @Error: 3, @Extended: 3 = Failed to modify Distance property.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bShadow
;                  |                               2 = Error setting $iLocation
;                  |                               4 = Error setting $iColor
;                  |                               8 = Error setting $iDistance
;                  |                               16 = Error setting $iBlur
;                  |                               32 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  LibreOffice may change the shadow distance +/- a Hundredth of a Millimeter (HMM).
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleAreaShadow(ByRef $oPresStyle, $bShadow = Null, $iLocation = Null, $iColor = Null, $iDistance = Null, $iBlur = Null, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaShadow($oPresStyle, $bShadow, $iLocation, $iColor, $iDistance, $iBlur, $iTransparency)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleAreaShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleAreaTransparency
; Description ...: Set or retrieve Transparency settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleAreaTransparency(ByRef $oPresStyleStyle[, $iTransparency = Null])
; Parameters ....: $oPresStyleStyle     - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iTransparency       - [optional] (0-100) Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyleStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Transparency value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleAreaTransparency(ByRef $oPresStyleStyle, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyleStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaTransparency($oPresStyleStyle, $iTransparency)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleAreaTransparencyGradient
; Description ...: Set or retrieve the Presentation Style transparency gradient settings.
; Syntax ........: _LOImpress_ShapePresStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oPresStyle[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iType               - [optional] (-1-5) Default is Null. The type of transparency gradient that you want to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3. Call with $LOI_GRAD_TYPE_OFF to turn Transparency Gradient off.
;                  $iXCenter            - [optional] (0-100) Default is Null. The horizontal offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] (0-100) Default is Null. The vertical offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] (0-359) Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] (0-100) Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iStart              - [optional] (0-100) Default is Null. The transparency value for the beginning point of the gradient, where 0% is fully opaque and 100% is fully transparent.
;                  $iEnd                - [optional] (0-100) Default is Null. The transparency value for the endpoint of the gradient, where 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. Transparency Gradient has been successfully turned off.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 3 = $iType not an Integer, less than -1 or greater than 5. See constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 5 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 6 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 7 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 9 = $iEnd not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillTransparenceGradient" Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Error creating Transparency Gradient Name.
;                  @Error: 3, @Extended: 4 = Error setting Transparency Gradient Name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iType
;                  |                               2 = Error setting $iXCenter
;                  |                               4 = Error setting $iYCenter
;                  |                               8 = Error setting $iAngle
;                  |                               16 = Error setting $iTransitionStart
;                  |                               32 = Error setting $iStart
;                  |                               64 = Error setting $iEnd
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oPresStyle, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOImpress_ShapeStyleAreaTransparencyGradient($oDoc, $oPresStyle, $iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Presentation Style's Multi Transparency Gradient settings.
; Syntax ........: _LOImpress_ShapePresStyleAreaTransparencyGradientMulti(ByRef $oPresStyle[, $avColorStops = Null])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $avColorStops        - [optional] Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: ?, Return: Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 2 = $avColorStops not an Array, or does not contain two columns.
;                  @Error: 1, @Extended: 3 = $avColorStops contains less than two rows.
;                  @Error: 1, @Extended: 4 = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error: 1, @Extended: 5 = ColorStop Transparency value not an Integer, less than 0 or greater than 100. Returning problem element index.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve FillTransparenceGradient Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Failed to retrieve StopColor Struct.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current version less than 7.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple Transparency stops in a Gradient rather than just a beginning and an ending value, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the Transparency percentage value between 0 and 100%.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LO_TransparencyGradientMultiModify, _LO_TransparencyGradientMultiDelete, _LO_TransparencyGradientMultiAdd, _LOImpress_ShapeAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleAreaTransparencyGradientMulti(ByRef $oPresStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaTransparencyGradientMulti($oPresStyle, $avColorStops)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleCharEffect
; Description ...: Set or Retrieve the Font Effect settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleCharEffect(ByRef $oPresStyle[, $iCase = Null[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iCase               - [optional] (0-4) Default is Null. The Character Case Style. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRelief             - [optional] (0-2) Default is Null. The Character Relief style. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOutline            - [optional] Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleCharEffect(ByRef $oPresStyle, $iCase = Null, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharEffect($oPresStyle, $iCase, $iRelief, $bOutline, $bShadow)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleCharEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleCharFont
; Description ...: Set and Retrieve the Font Settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleCharFont(ByRef $oPresStyle[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $sFontName           - [optional] Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] Default is Null. The new Font size.
;                  $iPosture            - [optional] (0-5) Default is Null. The Font Italic setting. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] (0, 50-200) Default is Null. The Font Bold settings see Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleCharFont(ByRef $oPresStyle, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharFont($oPresStyle, $sFontName, $nFontSize, $iPosture, $iWeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleCharFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleCharFontColor
; Description ...: Set or retrieve the font color and highlighting values for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleCharFontColor(ByRef $oPresStyle[, $iFontColor = Null[, $iHighlight = Null]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iFontColor          - [optional] (-1-16777215) Default is Null. The font Color value, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iHighlight          - [optional] (-1-16777215) Default is Null. The highlight Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 3 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $FontColor
;                  |                               2 = Error setting $iHighlight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleCharFontColor(ByRef $oPresStyle, $iFontColor = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_StyleCharFontColor($oPresStyle, $iFontColor, $iHighlight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleCharFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleCharOverLine
; Description ...: Set and retrieve the OverLine settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleCharOverLine(ByRef $oPresStyle[, $iOverLineStyle = Null[, $iOLColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iOverLineStyle      - [optional] (0-18) Default is Null. The style of the Overline line, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iOLColor            - [optional] (-1-16777215) Default is Null. The Overline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not Overlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleCharOverLine(ByRef $oPresStyle, $iOverLineStyle = Null, $iOLColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharOverLine($oPresStyle, $iOverLineStyle, $iOLColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleCharOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleCharStrikeOut
; Description ...: Set or Retrieve the Strikeout settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleCharStrikeOut(ByRef $oPresStyle[, $iStrikeLineStyle = Null[, $bWordOnly = Null]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iStrikeLineStyle    - [optional] (0-6) Default is Null. The Strikeout Line Style, see constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bWordOnly           - [optional] Default is Null. If True, strike out is applied to words only, skipping whitespaces.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleCharStrikeOut(ByRef $oPresStyle, $iStrikeLineStyle = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharStrikeOut($oPresStyle, $iStrikeLineStyle, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleCharStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleCharUnderLine
; Description ...: Set and retrieve the Underline settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleCharUnderLine(ByRef $oPresStyle[, $iUnderLineStyle = Null[, $iULColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iUnderLineStyle     - [optional] (0-18) Default is Null. The Underline line style, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iULColor            - [optional] (-1-16777215) Default is Null. The underline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not underlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle an Object.
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
Func _LOImpress_ShapePresStyleCharUnderLine(ByRef $oPresStyle, $iUnderLineStyle = Null, $iULColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharUnderLine($oPresStyle, $iUnderLineStyle, $iULColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleCharUnderLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleGetObjByName
; Description ...: Retrieve a Presentation Style Object for use with other Presentation Style functions.
; Syntax ........: _LOImpress_ShapePresStyleGetObjByName(ByRef $oDoc, $sPresStyle)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sPresStyle          - The Presentation Style name to retrieve the Object for.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Presentation Style successfully retrieved, returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sPresStyle not a String.
;                  @Error: 1, @Extended: 3 = Presentation Style called in $sPresStyle not found in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Presentation Style Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_ShapePresStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleGetObjByName(ByRef $oDoc, $sPresStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPresStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.StyleFamilies.getByName("Default").hasByName($sPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPresStyle = $oDoc.StyleFamilies().getByName("Default").getByName($sPresStyle)
	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPresStyle)
EndFunc   ;==>_LOImpress_ShapePresStyleGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleLineArrowStyles
; Description ...: Set or Retrieve Presentation Style Line Start and End Arrow Style settings.
; Syntax ........: _LOImpress_ShapePresStyleLineArrowStyles(ByRef $oDoc, ByRef $oPresStyle[, $vStartStyle = Null[, $iStartWidth = Null[, $bStartCenter = Null[, $bSync = Null[, $vEndStyle = Null[, $iEndWidth = Null[, $bEndCenter = Null]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $vStartStyle         - [optional] (0-32, or String) Default is Null. The Arrow head to apply to the start of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iStartWidth         - [optional] (0-5004) Default is Null. The Width of the Starting Arrowhead, in Hundredths of a Millimeter (HMM).
;                  $bStartCenter        - [optional] Default is Null. If True, Places the center of the Start arrowhead on the endpoint of the line.
;                  $bSync               - [optional] Default is Null. If True, Synchronizes the Start Arrowhead settings with the end Arrowhead settings. See remarks.
;                  $vEndStyle           - [optional] (0-32, or String) Default is Null. The Arrow head to apply to the end of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iEndWidth           - [optional] (0-5004) Default is Null. The Width of the Ending Arrowhead, in Hundredths of a Millimeter (HMM).
;                  $bEndCenter          - [optional] Default is Null. If True, Places the center of the End arrowhead on the endpoint of the line.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 3 = $vStartStyle not a String, and not an Integer.
;                  @Error: 1, @Extended: 4 = $vStartStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iStartWidth not an Integer, less than 0 or greater than 5004.
;                  @Error: 1, @Extended: 6 = $bStartCenter not a Boolean.
;                  @Error: 1, @Extended: 7 = $bSync not a Boolean.
;                  @Error: 1, @Extended: 8 = $vEndStyle not a String, and not an Integer.
;                  @Error: 1, @Extended: 9 = $vSEndStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 10 = $iEndWidth not an Integer, less than 0 or greater than 5004.
;                  @Error: 1, @Extended: 11 = $bEndCenter not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to convert Constant to Arrowhead name.
;                  @Error: 3, @Extended: 2 = Failed to insert preset Arrowhead name and style.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $vStartStyle
;                  |                               2 = Error setting $iStartWidth
;                  |                               4 = Error setting $bStartCenter
;                  |                               8 = Error setting $bSync
;                  |                               16 = Error setting $vEndStyle
;                  |                               32 = Error setting $iEndWidth
;                  |                               64 = Error setting $bEndCenter
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When the arrowhead type "Arrow" is set in the LO UI, or upon creation of a line with arrows, the internal name of the arrowhead is set to an incrementing name of "Arrowheads x", where x is an Integer value. Since I have no way to determine if the head is a custom arrowhead or supposed to be the "Arrow" type, the return when this is present will be the name "Arrowheads x", and not $LOI_SHAPE_LINE_ARROW_TYPE_ARROW.
;                  When setting an Arrowhead to be $LOI_SHAPE_LINE_ARROW_TYPE_ARROW, the head is set correctly, but the LibreOffice UI will show "None". The return for Arrowhead type will be correct, $LOI_SHAPE_LINE_ARROW_TYPE_ARROW.
;                  LibreOffice has no setting for $bSync, so I have made a manual version of it in this function. It only accepts True, and must be called with True each time you want it to synchronize.
;                  When retrieving the current settings, $bSync will be a Boolean value of whether the Start Arrowhead settings are currently equal to the End Arrowhead setting values.
;                  Both $vStartStyle and $vEndStyle accept a String or an Integer because there is the possibility of a custom Arrowhead being available the user may want to use.
;                  When retrieving the current settings, both $vStartStyle and $vEndStyle could be either an Integer or a String. It will be a String if the current Arrowhead is a custom Arrowhead, else an Integer, corresponding to one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleLineArrowStyles(ByRef $oDoc, ByRef $oPresStyle, $vStartStyle = Null, $iStartWidth = Null, $bStartCenter = Null, $bSync = Null, $vEndStyle = Null, $iEndWidth = Null, $bEndCenter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOImpress_ShapeStyleLineArrowStyles($oDoc, $oPresStyle, $vStartStyle, $iStartWidth, $bStartCenter, $bSync, $vEndStyle, $iEndWidth, $bEndCenter)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleLineArrowStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleLineProperties
; Description ...: Set or Retrieve Presentation Style Line settings.
; Syntax ........: _LOImpress_ShapePresStyleLineProperties(ByRef $oDoc, ByRef $oPresStyle[, $vStyle = Null[, $iColor = Null[, $iWidth = Null[, $iTransparency = Null[, $iCornerStyle = Null[, $iCapStyle = Null]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $vStyle              - [optional] (0-31, or String) Default is Null. The Line Style to use. Can be a Custom Line Style name, or one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Line color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iWidth              - [optional] (0-5004) Default is Null. The line Width, set in Hundredths of a Millimeter (HMM).
;                  $iTransparency       - [optional] (0-100) Default is Null. The Line transparency percentage. 100% = fully transparent.
;                  $iCornerStyle        - [optional] (0, 2-4) Default is Null. The Line Corner Style. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  $iCapStyle           - [optional] (0-2) Default is Null. The Line Cap Style. See Constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 3 = $vStyle not a String, and not an Integer.
;                  @Error: 1, @Extended: 4 = $vStyle is an Integer, but less than 0 or greater than 31. See constants $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 6 = $iWidth not an Integer, less than 0 or greater than 5004.
;                  @Error: 1, @Extended: 7 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iCornerStyle not an Integer, not equal to 0, equal to 1, not equal to 2 or greater than 4. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error: 1, @Extended: 9 = $iCapStyle is an Integer, but less than 0 or greater than 2. See constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to convert Constant to Line Style name.
;                  @Error: 3, @Extended: 2 =  Failed to insert Line Style name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $vStyle
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $iWidth
;                  |                               8 = Error setting $iTransparency
;                  |                               16 = Error setting $iCornerStyle
;                  |                               32 = Error setting $iCapStyle
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $vStyle accepts a String or an Integer because there is the possibility of a custom Line Style being available that the user may want to use.
;                  When retrieving the current settings, $vStyle could be either an Integer or a String. It will be a String if the current Line Style is a custom Line Style, else an Integer, corresponding to one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleLineProperties(ByRef $oDoc, ByRef $oPresStyle, $vStyle = Null, $iColor = Null, $iWidth = Null, $iTransparency = Null, $iCornerStyle = Null, $iCapStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOImpress_ShapeStyleLineProperties($oDoc, $oPresStyle, $vStyle, $iColor, $iWidth, $iTransparency, $iCornerStyle, $iCapStyle)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleLineProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleNumCustomize
; Description ...: Retrieve and Set Numbering Customize settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleNumCustomize(ByRef $oDoc, ByRef $oPresStyle, $iLevel[, $iNumFormat = Null[, $iStartAt = Null[, $iColor = Null[, $iRelSize = Null[, $sSepBefore = Null[, $sSepAfter = Null[, $iCharDecimal = Null]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iLevel              - (0-10) The Numbering Level to modify; enter 0 to modify all levels.
;                  $iNumFormat          - [optional] (0-71) Default is Null. The numbering scheme for the selected levels. See Constants, $LOI_NUM_FRMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iStartAt            - [optional] Default is Null. A new starting number for the current level
;                  $iColor              - [optional] (-1-16777215) Default is Null. The color of the numbering symbol, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRelSize            - [optional] (25-400) Default is Null. The percentage to resize the numbering symbol, relative to the paragraph font size.
;                  $sSepBefore          - [optional] Default is Null. A character or the text to display in front of the number in the list.
;                  $sSepAfter           - [optional] Default is Null. A character or the text to display behind the number in the list.
;                  $iCharDecimal        - [optional] Default is Null. The ASCII Decimal character code value (See ASC function) of the desired character. Note: $iNumFormat must be set to $LOI_NUM_FRMT_CHAR_SPECIAL(6) before these can be set.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully set the requested Properties.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters. See remarks.
;                  @Error: 0, @Extended: 2, Return: Array = Success. All optional parameters were called with Null, returning a 10 Element Array containing arrays of settings for each Numbering level corresponding to their position in the array. Each array will be as described above. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 3 = $oPresStyle not a Presentation Style Object.
;                  @Error: 1, @Extended: 4 = $iLevel not between 0 - 10.
;                  @Error: 1, @Extended: 5 = $iNumFormat not an Integer, less than 0 or greater than 71. See Constants, $LOI_NUM_FRMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iStartAt not an Integer.
;                  @Error: 1, @Extended: 7 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 8 = $iRelSize not an Integer, less than 25 or greater than 400.
;                  @Error: 1, @Extended: 9 = $sSepBefore not a string.
;                  @Error: 1, @Extended: 10 = $sSepAfter not a string.
;                  @Error: 1, @Extended: 11 = $iCharDecimal not an Integer.
;                  @Error: 1, @Extended: 12 = $iCharDecimal was called and Number Format not set to $LOI_NUM_FRMT_CHAR_SPECIAL.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error mapping setting values.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Numbering Rules Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Numbering Rule Array for level.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iNumFormat
;                  |                               2 = Error setting $iStartAt
;                  |                               4 = Error setting $iColor
;                  |                               8 = Error setting $iRelSize
;                  |                               16 = Error setting $sSepBefore
;                  |                               32 = Error setting $sSepAfter
;                  |                               64 = Error setting $iCharDecimal
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function should work just fine as the others do for modifying styles, but for setting Numbering Style settings, it would seem that the Array of Setting Objects passed by AutoIt is not recognized as an appropriate array/sequence by LibreOffice, and consequently causes a com.sun.star.lang.IllegalArgumentException COM error. See __LOImpress_ShapePresStyleNumModify function for a more detailed explanation. This function can still be used to set and retrieve, setting values, however now, this function either inserts a temporary macro into $oDoc for performing the needed procedure, or if that fails, it invisibly opens an .odt Libre document and inserts a macro, see __LOImpress_ShapePresStyleNumInitiateDocument which is then called with the necessary parameters to set.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  If Current numbering type is set to Bullet, the returned array will be a 7 Element Array with values in order of function parameters, the parameters $iStartAt, $sSepBefore, and $sSepAfter will return a Null value, as they are not valid for Bullets.
;                  If the current numbering type is other than bullet style, a 7 element array will be returned, the $iCharDecimal parameter will return a Null value.
;                  You can request setting values for one numbering level at a time, or all at once (see below).
;                  If you retrieve the current settings for all levels (by calling $iLevel with 0), the return will be a 10 element array containing an array of settings for each Numbering Level.
;                  Call any optional parameter with Null keyword to skip it.
;                  When a lot of settings are set, especially for all levels, this function can be a bit slow.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleNumCustomize(ByRef $oDoc, ByRef $oPresStyle, $iLevel, $iNumFormat = Null, $iStartAt = Null, $iColor = Null, $iRelSize = Null, $sSepBefore = Null, $sSepAfter = Null, $iCharDecimal = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumRules
	Local $iError = 0
	Local $avCustomize[7], $aaAllLevels[10]
	Local $atNumLevel[0]
	Local $mNumLevel[]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oPresStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iLevel, 0, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iLevel = ($iLevel - 1) ; Numbering Levels are  0 based, minus 1 to compensate.

	$oNumRules = $oPresStyle.NumberingRules()
	If Not IsObj($oNumRules) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iNumFormat, $iStartAt, $iColor, $iRelSize, $sSepBefore, $sSepAfter, $iCharDecimal) Then
		For $i = (($iLevel = -1) ? (0) : ($iLevel)) To (($iLevel = -1) ? (9) : ($iLevel)) ; Determine if I'm retrieving settings for all levels or just one.
			$atNumLevel = $oNumRules.getByIndex($i)
			If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$mNumLevel = __LOImpress_NumRuleCreateMap($atNumLevel)
			If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			If MapExists($mNumLevel, "BulletChar") Then
				__LO_ArrayFill($avCustomize, $atNumLevel[$mNumLevel["NumberingType"]].Value(), _
						Null, _
						$atNumLevel[$mNumLevel["BulletColor"]].Value(), _
						$atNumLevel[$mNumLevel["BulletRelSize"]].Value(), _
						Null, Null, _
						Asc($atNumLevel[$mNumLevel["BulletChar"]].Value()))

			Else ; If not set for Bullet style, return only these settings as BulletChar doesn't exist.
				__LO_ArrayFill($avCustomize, $atNumLevel[$mNumLevel["NumberingType"]].Value(), _
						$atNumLevel[$mNumLevel["StartWith"]].Value(), _
						$atNumLevel[$mNumLevel["BulletColor"]].Value(), _
						$atNumLevel[$mNumLevel["BulletRelSize"]].Value(), _
						$atNumLevel[$mNumLevel["Prefix"]].Value(), _
						$atNumLevel[$mNumLevel["Suffix"]].Value(), Null)
			EndIf

			If ($iLevel = -1) Then $aaAllLevels[$i] = $avCustomize
		Next

		Return ($iLevel = -1) ? (SetError($__LO_STATUS_SUCCESS, 2, $aaAllLevels)) : (SetError($__LO_STATUS_SUCCESS, 1, $avCustomize))
	EndIf

	For $i = (($iLevel = -1) ? (0) : ($iLevel)) To (($iLevel = -1) ? (9) : ($iLevel)) ; Determine if I am setting settings for all levels or just one.
		$atNumLevel = $oNumRules.getByIndex($i)
		If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$mNumLevel = __LOImpress_NumRuleCreateMap($atNumLevel) ; Map what elements each setting is located at.
		If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		If ($iNumFormat <> Null) Then
			If Not __LO_IntIsBetween($iNumFormat, $LOI_NUM_FRMT_CHARS_UPPER_LETTER, $LOI_NUM_FRMT_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$atNumLevel[$mNumLevel["NumberingType"]].Value = $iNumFormat

			__LOImpress_ShapePresStyleNumModify($oDoc, $oNumRules, $i, $atNumLevel) ; Modify the Setting in case it is switching from/to a bullet type.

			$atNumLevel = $oNumRules.getByIndex($i)
			If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$mNumLevel = __LOImpress_NumRuleCreateMap($atNumLevel)
			If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		EndIf

		If ($iStartAt <> Null) Then
			If Not IsInt($iStartAt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$atNumLevel[$mNumLevel["StartWith"]].Value = $iStartAt
		EndIf

		If ($iColor <> Null) Then
			If Not __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$atNumLevel[$mNumLevel["BulletColor"]].Value = $iColor
		EndIf

		If ($iRelSize <> Null) Then
			If Not __LO_IntIsBetween($iRelSize, 25, 400) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$atNumLevel[$mNumLevel["BulletRelSize"]].Value = $iRelSize
		EndIf

		If ($sSepBefore <> Null) Then
			If Not IsString($sSepBefore) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$atNumLevel[$mNumLevel["Prefix"]].Value = $sSepBefore
		EndIf

		If ($sSepAfter <> Null) Then
			If Not IsString($sSepAfter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$atNumLevel[$mNumLevel["Suffix"]].Value = $sSepAfter
		EndIf

		If ($iCharDecimal <> Null) Then
			If Not IsInt($iCharDecimal) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
			If Not MapExists($mNumLevel, "BulletChar") Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

			$atNumLevel[$mNumLevel["BulletChar"]].Value = Chr($iCharDecimal)
		EndIf

		__LOImpress_ShapePresStyleNumModify($oDoc, $oNumRules, $i, $atNumLevel)
		$oPresStyle.NumberingRules = $oNumRules

		$atNumLevel = $oPresStyle.NumberingRules.getByIndex($i)
		If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$mNumLevel = __LOImpress_NumRuleCreateMap($atNumLevel)
		If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		; Error Checking
		$iError = (__LO_VarsAreNull($iNumFormat)) ? ($iError) : (($atNumLevel[$mNumLevel["NumberingType"]].Value() = $iNumFormat) ? ($iError) : (BitOR($iError, 1)))
		$iError = (__LO_VarsAreNull($iStartAt)) ? ($iError) : (($atNumLevel[$mNumLevel["StartWith"]].Value() = $iStartAt) ? ($iError) : (BitOR($iError, 2)))
		$iError = (__LO_VarsAreNull($iColor)) ? ($iError) : (($atNumLevel[$mNumLevel["BulletColor"]].Value() = $iColor) ? ($iError) : (BitOR($iError, 4)))
		$iError = (__LO_VarsAreNull($iRelSize)) ? ($iError) : (($atNumLevel[$mNumLevel["BulletRelSize"]].Value() = $iRelSize) ? ($iError) : (BitOR($iError, 8)))
		$iError = (__LO_VarsAreNull($sSepBefore)) ? ($iError) : (($atNumLevel[$mNumLevel["Prefix"]].Value() = $sSepBefore) ? ($iError) : (BitOR($iError, 16)))
		$iError = (__LO_VarsAreNull($sSepAfter)) ? ($iError) : (($atNumLevel[$mNumLevel["Suffix"]].Value() = $sSepAfter) ? ($iError) : (BitOR($iError, 32)))
		$iError = (__LO_VarsAreNull($iCharDecimal)) ? ($iError) : ((Asc($atNumLevel[$mNumLevel["BulletChar"]].Value()) = $iCharDecimal) ? ($iError) : (BitOR($iError, 64)))
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapePresStyleNumCustomize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleParAlignment
; Description ...: Set and Retrieve Paragraph Alignment settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleParAlignment(ByRef $oPresStyle[, $iHorAlign = Null[, $iLastLineAlign = Null[, $iTxtDirection = Null]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iHorAlign           - [optional] (0-3) Default is Null. The Horizontal alignment of the paragraph. See Constants, $LOI_PAR_ALIGN_HOR_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iLastLineAlign      - [optional] (0-3) Default is Null. Specify the alignment for the last line in the paragraph. See Constants, $LOI_PAR_LAST_LINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iTxtDirection       - [optional] (0-5) Default is Null. The Text Writing Direction. See Constants, $LOI_PAR_TXT_DIR_* as defined in LibreOfficeImpress_Constants.au3. [LibreOffice Default is 4]
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleParAlignment(ByRef $oPresStyle, $iHorAlign = Null, $iLastLineAlign = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParAlignment($oPresStyle, $iHorAlign, $iLastLineAlign, $iTxtDirection)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleParAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleParIndent
; Description ...: Set or Retrieve Paragraph Indent settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleParIndent(ByRef $oPresStyle[, $iBeforeTxt = Null[, $iAfterTxt = Null[, $iFirstLine = Null]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iBeforeTxt          - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM).
;                  $iAfterTxt           - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM)
;                  $iFirstLine          - [optional] (0-1162202) Default is Null. Indentation distance of the first line of a paragraph. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleParIndent(ByRef $oPresStyle, $iBeforeTxt = Null, $iAfterTxt = Null, $iFirstLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParIndent($oPresStyle, $iBeforeTxt, $iAfterTxt, $iFirstLine)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleParIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleParSpacing
; Description ...: Set and Retrieve Line Spacing settings for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleParSpacing(ByRef $oPresStyle[, $iAbovePar = Null[, $iBelowPar = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null]]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iAbovePar           - [optional] (0-100000) Default is Null. The Space above a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iBelowPar           - [optional] (0-100000) Default is Null. The Space Below a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iLineSpcMode        - [optional] (0-3) Default is Null. The line spacing type of the paragraph. See Constants, $LOI_PAR_LINE_SPC_MODE_* as defined in LibreOfficeImpress_Constants.au3, also notice min and max values for each.
;                  $iLineSpcHeight      - [optional] Default is Null. This value specifies the height in regard to Mode. See Remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleParSpacing(ByRef $oPresStyle, $iAbovePar = Null, $iBelowPar = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParSpacing($oPresStyle, $iAbovePar, $iBelowPar, $iLineSpcMode, $iLineSpcHeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleParSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleParTabStopCreate
; Description ...: Create a new TabStop for a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleParTabStopCreate(ByRef $oPresStyle, $iPosition[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iPosition           - The TabStop position to set the new TabStop to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Settings were successfully set. New TabStop position is returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleParTabStopCreate(ByRef $oPresStyle, $iPosition, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopCreate($oPresStyle, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleParTabStopCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleParTabStopDelete
; Description ...: Delete a TabStop from a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleParTabStopDelete(ByRef $oPresStyle, $iTabStop)
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Returning True if TabStop was successfully deleted, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleParTabStopDelete(ByRef $oPresStyle, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopDelete($oPresStyle, $iTabStop)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleParTabStopDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop in a Shape Style.
; Syntax ........: _LOImpress_ShapePresStyleParTabStopMod(ByRef $oPresStyle, $iTabStop[, $iPosition = Null[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - [optional] Default is Null. The New position to set the input position to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: ?, Return: 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
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
Func _LOImpress_ShapePresStyleParTabStopMod(ByRef $oPresStyle, $iTabStop, $iPosition = Null, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopMod($oPresStyle, $iTabStop, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleParTabStopMod

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleParTabStopsGetList
; Description ...: Retrieve an array of TabStops available in a Presentation Style.
; Syntax ........: _LOImpress_ShapePresStyleParTabStopsGetList(ByRef $oPresStyle)
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array of TabStops. @Extended set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleParTabStopsGetList(ByRef $oPresStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopsGetList($oPresStyle)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleParTabStopsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStylesGetNames
; Description ...: Retrieve an array of all Presentation Style names available for a document.
; Syntax ........: _LOImpress_ShapePresStylesGetNames(ByRef $oDoc[, $bAppliedOnly = False[, $bDisplayName = False]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bAppliedOnly        - [optional] Default is False. If True, only Applied Presentation Styles are returned.
;                  $bDisplayName        - [optional] Default is False. If True, the style name displayed in the UI (Display Name), instead of the programmatic style name, is returned. See remarks.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array containing all Presentation Styles matching the called parameters. See remarks. @Extended contains the count of results returned.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bAppliedOnly not a Boolean.
;                  @Error: 1, @Extended: 3 = $bDisplayName not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Array of Presentation Style names.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If Only a Document object is called, all available Presentation styles will be returned.
;                  If $bAppliedOnly is called with True, only styles that are applied are returned.
;                  Ten Presentation styles have different internal names:
;                  - "Background objects" is internally called "backgroundobjects".
;                  - "Outline 1" is internally called "outline1".
;                  - "Outline 2" is internally called "outline2".
;                  - "Outline 3" is internally called "outline3".
;                  - "Outline 4" is internally called "outline4".
;                  - "Outline 5" is internally called "outline5".
;                  - "Outline 6" is internally called "outline6".
;                  - "Outline 7" is internally called "outline7".
;                  - "Outline 8" is internally called "outline8".
;                  - "Outline 9" is internally called "outline9".
;                  Previous to LibreOffice 25.2 either name would work when setting a Style, however after 25.2 only the internal, or programmatic style names, will work.
;                  Calling $bDisplayName with True will return a list of Style names, as the user sees them in the UI, in the same order as they are returned if $bDisplayName is False. It is best not to use these when setting Styling.
; Related .......: _LOImpress_ShapePresStyleGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStylesGetNames(ByRef $oDoc, $bAppliedOnly = False, $bDisplayName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asStyles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bDisplayName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$asStyles = __LO_StylesGetNames($oDoc, "Default", False, $bAppliedOnly, $bDisplayName)
	If Not IsArray($asStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asStyles), $asStyles)
EndFunc   ;==>_LOImpress_ShapePresStylesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleTextAttrFit
; Description ...: Set or Retrieve Presentation Style Text Attribute Fit properties.
; Syntax ........: _LOImpress_ShapePresStyleTextAttrFit(ByRef $oPresStyle[, $bFitWidth = Null[, $bFitHeight = Null[, $bFitToFrame = Null]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $bFitWidth           - [optional] Default is Null. If True, Expands the width of the object to the width of the text.
;                  $bFitHeight          - [optional] Default is Null. If True, Expands the height of the object to the height of the text.
;                  $bFitToFrame         - [optional] Default is Null. If True, Resizes the text to fit the entire area of the drawing object.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 2 = $bFitWidth not a Boolean.
;                  @Error: 1, @Extended: 3 = $bFitHeight not a Boolean.
;                  @Error: 1, @Extended: 4 = $bFitToFrame not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFitWidth
;                  |                               2 = Error setting $bFitHeight
;                  |                               4 = Error setting $bFitToFrame
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Properties as found in the UI, and their equivalent: "Fit Width to Text" = $bFitWidth. "Fit Height to Text" = $bFitHeight. "Fit to Frame" = $bFitToFrame.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleTextAttrFit(ByRef $oPresStyle, $bFitWidth = Null, $bFitHeight = Null, $bFitToFrame = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__LOI_TEXT_FIT_NONE = 0, $__LOI_TEXT_FIT_PROP = 1 ; com.sun.star.drawing.TextFitToSizeType
	Local $iError = 0
	Local $avTextAttr[3]

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bFitWidth, $bFitHeight, $bFitToFrame) Then
		__LO_ArrayFill($avTextAttr, $oPresStyle.TextAutoGrowWidth(), $oPresStyle.TextAutoGrowHeight(), ($oPresStyle.TextFitToSize() = $__LOI_TEXT_FIT_PROP) ? (True) : (False))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTextAttr)
	EndIf

	; I could use the internal function __LOImpress_ShapeTextAttrFit, but I would have to ReDim the Array returned, this is simpler.
	If ($bFitWidth <> Null) Then
		If Not IsBool($bFitWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oPresStyle.TextAutoGrowWidth = $bFitWidth
		$iError = ($oPresStyle.TextAutoGrowWidth() = $bFitWidth) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bFitHeight <> Null) Then
		If Not IsBool($bFitHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPresStyle.TextAutoGrowHeight = $bFitHeight
		$iError = ($oPresStyle.TextAutoGrowHeight() = $bFitHeight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bFitToFrame <> Null) Then
		If Not IsBool($bFitToFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPresStyle.TextFitToSize = ($bFitToFrame) ? ($__LOI_TEXT_FIT_PROP) : ($__LOI_TEXT_FIT_NONE)
		$iError = ($oPresStyle.TextFitToSize() = ($bFitToFrame) ? ($__LOI_TEXT_FIT_PROP) : ($__LOI_TEXT_FIT_NONE)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapePresStyleTextAttrFit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapePresStyleTextAttrSettings
; Description ...: Set or Retrieve Presentation Style text Attribute settings.
; Syntax ........: _LOImpress_ShapePresStyleTextAttrSettings(ByRef $oPresStyle[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null[, $iAnchor = Null[, $bFullWidth = Null]]]]]])
; Parameters ....: $oPresStyle          - A Presentation Style object returned by a previous _LOImpress_ShapePresStyleGetObjByName function.
;                  $iLeft               - [optional] (-100000-100000) Default is Null. The space between the left edge of the drawing object and the left border of the text, in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] (-100000-100000) Default is Null. The space between the right edge of the drawing object and the right border of the text, in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] (-100000-100000) Default is Null. The space between the top edge of the drawing object and the top border of the text, in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] (-100000-100000) Default is Null. The space between the bottom edge of the drawing object and the bottom border of the text, in Hundredths of a Millimeter (HMM).
;                  $iAnchor             - [optional] (0-8) Default is Null. The text anchor position. See Constants, $LOI_PAR_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bFullWidth          - [optional] Default is Null. If True, Anchors the text to the full width of the drawing object.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPresStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iLeft not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 3 = $iRight not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 4 = $iTop not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 5 = $iBottom not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 6 = $iAnchor  not an Integer, less than 0 or greater than 8. See Constants, $LOI_PAR_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 7 = $bFullWidth not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
;                  |                               16 = Error setting $iAnchor
;                  |                               32 = Error setting $bFullWidth
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapePresStyleTextAttrSettings(ByRef $oPresStyle, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null, $iAnchor = Null, $bFullWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPresStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeTextAttrSettings($oPresStyle, $iLeft, $iRight, $iTop, $iBottom, $iAnchor, $bFullWidth)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapePresStyleTextAttrSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeRotateSlant
; Description ...: Set or retrieve Rotation and Slant settings for a Shape.
; Syntax ........: _LOImpress_ShapeRotateSlant(ByRef $oShape[, $nRotate = Null[, $nSlant = Null]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $nRotate             - [optional] (0-359.99) Default is Null. The Degrees to rotate the shape. See remarks.
;                  $nSlant              - [optional] (-89-89.00) Default is Null. The Degrees to slant the shape. See remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $nRotate not a Number, less than 0 or greater than 359.99.
;                  @Error: 1, @Extended: 3 = $nSlant not a Number, less than -89 or greater than 89.00.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $nRotate
;                  |                               2 = Error setting $nSlant
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you attempt to apply rotation to an already slanted Shape, or vice versa, a property setting error will occur, and the values will be very inaccurately applied.
;                  This function uses the deprecated LibreOffice methods RotateAngle, and ShearAngle, and may stop working in future LibreOffice versions, after 7.3.4.2.
;                  At the present time Control Point settings are not included as they are too complex to manipulate.
;                  At the present time Corner Radius setting is not included, as I was unable to identify a shape that utilized this setting.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
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
; Name ..........: _LOImpress_ShapesGetList
; Description ...: Retrieve an array of Shapes (Text Boxes, DrawShapes, Images etc) contained in a Slide.
; Syntax ........: _LOImpress_ShapesGetList(ByRef $oSlide[, $iTypes = $LOI_SHAPE_TYPE_ALL])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iTypes              - [optional] (0-1023) Default is $LOI_SHAPE_TYPE_ALL. The type of Shapes to return in the Array. Can be BitOR'd. See Constants, $LOI_SHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. A two columned Array containing the Shape Objects contained in the Slide. See Remarks. @Extended is set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iTypes not an Integer, less than 1 or greater than 1023. See Constants, $LOI_SHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Shape Object.
;                  @Error: 3, @Extended: 2 = Failed to identify Shape Type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array returned has two columns. The first column is the shape Object. The second column is the Shape Type, corresponding to one of the Constants $LOI_SHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Related .......: _LOImpress_DrawShapeGetType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapesGetList(ByRef $oSlide, $iTypes = $LOI_SHAPE_TYPE_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avShapes[0][2]
	Local $oShape
	Local $iShapeType, $iCount = 0

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iTypes, $LOI_SHAPE_TYPE_DRAWING_SHAPE, $LOI_SHAPE_TYPE_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oSlide.hasElements() Then
		ReDim $avShapes[$oSlide.getCount()][2]

		For $i = 0 To $oSlide.getCount() - 1
			$oShape = $oSlide.getByIndex($i)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$iShapeType = __LOImpress_ShapeGetType($oShape)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			If (BitAND($iTypes, $iShapeType) = $iShapeType) Then
				$avShapes[$iCount][0] = $oShape
				$avShapes[$iCount][1] = $iShapeType
				$iCount += 1
			EndIf
			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		ReDim $avShapes[$iCount][2]
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $avShapes)
EndFunc   ;==>_LOImpress_ShapesGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeSize
; Description ...: Set or Retrieve Shape Size related settings.
; Syntax ........: _LOImpress_ShapeSize(ByRef $oShape[, $iWidth = Null[, $iHeight = Null[, $bProtectSize = Null]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iWidth              - [optional] Default is Null. The width of the Shape, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iHeight             - [optional] Default is Null. The height of the Shape, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $bProtectSize        - [optional] Default is Null. If True, Locks the size of the Shape.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer, or less than 51.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer, or less than 51.
;                  @Error: 1, @Extended: 4 = $bProtectSize not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Shape Structure.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iHeight
;                  |                               4 = Error setting $bProtectSize
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I have skipped "Keep Ratio", as currently it seems unable to be set for shapes.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
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
; Name ..........: _LOImpress_ShapeStyleAreaColor
; Description ...: Set or Retrieve the Fill color settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleAreaColor(ByRef $oShapeStyle[, $iColor = Null])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iColor              - [optional] (-1-16777215) Default is Null. The Fill color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current Fill color as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current color value.
;                  @Error: 3, @Extended: 2 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iColor
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleAreaColor(ByRef $oShapeStyle, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeStyleAreaColor($oShapeStyle, $iColor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOImpress_ShapeStyleAreaFillStyle(ByRef $oShapeStyle)
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOI_AREA_FILL_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Fill Style.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleAreaFillStyle(ByRef $oShapeStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oShapeStyle.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOImpress_ShapeStyleAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleAreaGradient
; Description ...: Modify or retrieve the settings for Shape Style Background color Gradient.
; Syntax ........: _LOImpress_ShapeStyleAreaGradient(ByRef $oDoc, ByRef $oShapeStyle[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $sGradientName       - [optional] Default is Null. A Preset Gradient Name. See remarks. See constants, $LOI_GRAD_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iType               - [optional] (-1-5) Default is Null. The gradient type to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iIncrement          - [optional] (0, 3-256) Default is Null. The number of steps of color change. 0 = Automatic.
;                  $iXCenter            - [optional] (0-100) Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] (0-100) Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] (0-359) Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] (0-100) Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] (0-16777215) Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] (0-16777215) Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] (0-100) Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] (0-100) Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. Gradient has been successfully turned off.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 3 = $sGradientName not a String.
;                  @Error: 1, @Extended: 4 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error: 1, @Extended: 6 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 9 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 10 = $iFromColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 11 = $iToColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 12 = $iFromIntense not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 13 = $iToIntense not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillGradient" Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Error creating Gradient Name.
;                  @Error: 3, @Extended: 4 = Error setting Gradient Name.
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
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleAreaGradient(ByRef $oDoc, ByRef $oShapeStyle, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeStyleAreaGradient($oDoc, $oShapeStyle, $sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleAreaGradientMulticolor
; Description ...: Set or Retrieve a Shape Style's Multicolor Gradient settings.
; Syntax ........: _LOImpress_ShapeStyleAreaGradientMulticolor(ByRef $oShapeStyle[, $avColorStops = Null])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $avColorStops        - [optional] Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: ?, Return: Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $avColorStops not an Array, or does not contain two columns.
;                  @Error: 1, @Extended: 3 = $avColorStops contains less than two rows.
;                  @Error: 1, @Extended: 4 = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error: 1, @Extended: 5 = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve FillGradient Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Failed to retrieve StopColor Struct.
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
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LO_GradientMulticolorAdd, _LO_GradientMulticolorDelete, _LO_GradientMulticolorModify, _LOImpress_ShapeAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleAreaGradientMulticolor(ByRef $oShapeStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaGradientMulticolor($oShapeStyle, $avColorStops)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleAreaShadow
; Description ...: Set or Retrieve the shadow settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleAreaShadow(ByRef $oShapeStyle[, $bShadow = Null[, $iLocation = Null[, $iColor = Null[, $iDistance = Null[, $iBlur = Null[, $iTransparency = Null]]]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $bShadow             - [optional] Default is Null. If True, a Shadow is present for the Shape.
;                  $iLocation           - [optional] (0-8) Default is Null. The Location of the Shadow, must be one of the Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Shadow color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iDistance           - [optional] Default is Null. The distance of the Shadow from the Shape's edges, set in Hundredths of a Millimeter (HMM).
;                  $iBlur               - [optional] (0-150) Default is Null. The amount of blur applied to the Shadow, set in Printer's Points.
;                  $iTransparency       - [optional] (0-100) Default is Null. The percentage of Shadow transparency. 100% means completely transparent.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $bShadow not a Boolean.
;                  @Error: 1, @Extended: 3 = $iLocation not an Integer, less than 0 or greater than 8. See Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iDistance not an Integer, or less than 0.
;                  @Error: 1, @Extended: 6 = $iBlur not an Integer, less than 0 or greater than 150 Printer's Points.
;                  @Error: 1, @Extended: 7 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Distance and Location Values.
;                  @Error: 3, @Extended: 2 = Failed to modify Location property.
;                  @Error: 3, @Extended: 3 = Failed to modify Distance property.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bShadow
;                  |                               2 = Error setting $iLocation
;                  |                               4 = Error setting $iColor
;                  |                               8 = Error setting $iDistance
;                  |                               16 = Error setting $iBlur
;                  |                               32 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  LibreOffice may change the shadow distance +/- a Hundredth of a Millimeter (HMM).
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleAreaShadow(ByRef $oShapeStyle, $bShadow = Null, $iLocation = Null, $iColor = Null, $iDistance = Null, $iBlur = Null, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaShadow($oShapeStyle, $bShadow, $iLocation, $iColor, $iDistance, $iBlur, $iTransparency)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleAreaShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleAreaTransparency
; Description ...: Set or retrieve Transparency settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleAreaTransparency(ByRef $oShapeStyle[, $iTransparency = Null])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iTransparency       - [optional] (0-100) Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Transparency value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleAreaTransparency(ByRef $oShapeStyle, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaTransparency($oShapeStyle, $iTransparency)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleAreaTransparencyGradient
; Description ...: Set or retrieve the Shape Style transparency gradient settings.
; Syntax ........: _LOImpress_ShapeStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oShapeStyle[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iType               - [optional] (-1-5) Default is Null. The type of transparency gradient that you want to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3. Call with $LOI_GRAD_TYPE_OFF to turn Transparency Gradient off.
;                  $iXCenter            - [optional] (0-100) Default is Null. The horizontal offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] (0-100) Default is Null. The vertical offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] (0-359) Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] (0-100) Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iStart              - [optional] (0-100) Default is Null. The transparency value for the beginning point of the gradient, where 0% is fully opaque and 100% is fully transparent.
;                  $iEnd                - [optional] (0-100) Default is Null. The transparency value for the endpoint of the gradient, where 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. Transparency Gradient has been successfully turned off.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 3 = $iType not an Integer, less than -1 or greater than 5. See constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 5 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 6 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 7 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 9 = $iEnd not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillTransparenceGradient" Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Error creating Transparency Gradient Name.
;                  @Error: 3, @Extended: 4 = Error setting Transparency Gradient Name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iType
;                  |                               2 = Error setting $iXCenter
;                  |                               4 = Error setting $iYCenter
;                  |                               8 = Error setting $iAngle
;                  |                               16 = Error setting $iTransitionStart
;                  |                               32 = Error setting $iStart
;                  |                               64 = Error setting $iEnd
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleAreaTransparencyGradient(ByRef $oDoc, ByRef $oShapeStyle, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOImpress_ShapeStyleAreaTransparencyGradient($oDoc, $oShapeStyle, $iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Shape Style's Multi Transparency Gradient settings.
; Syntax ........: _LOImpress_ShapeStyleAreaTransparencyGradientMulti(ByRef $oShapeStyle[, $avColorStops = Null])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $avColorStops        - [optional] Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: ?, Return: Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $avColorStops not an Array, or does not contain two columns.
;                  @Error: 1, @Extended: 3 = $avColorStops contains less than two rows.
;                  @Error: 1, @Extended: 4 = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error: 1, @Extended: 5 = ColorStop Transparency value not an Integer, less than 0 or greater than 100. Returning problem element index.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve FillTransparenceGradient Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Failed to retrieve StopColor Struct.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current version less than 7.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple Transparency stops in a Gradient rather than just a beginning and an ending value, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the Transparency percentage value between 0 and 100%.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LO_TransparencyGradientMultiModify, _LO_TransparencyGradientMultiDelete, _LO_TransparencyGradientMultiAdd, _LOImpress_ShapeAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleAreaTransparencyGradientMulti(ByRef $oShapeStyle, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaTransparencyGradientMulti($oShapeStyle, $avColorStops)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleCharEffect
; Description ...: Set or Retrieve the Font Effect settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleCharEffect(ByRef $oShapeStyle[, $iCase = Null[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iCase               - [optional] (0-4) Default is Null. The Character Case Style. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRelief             - [optional] (0-2) Default is Null. The Character Relief style. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOutline            - [optional] Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleCharEffect(ByRef $oShapeStyle, $iCase = Null, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharEffect($oShapeStyle, $iCase, $iRelief, $bOutline, $bShadow)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleCharEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleCharFont
; Description ...: Set and Retrieve the Font Settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleCharFont(ByRef $oShapeStyle[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $sFontName           - [optional] Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] Default is Null. The new Font size.
;                  $iPosture            - [optional] (0-5) Default is Null. The Font Italic setting. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] (0, 50-200) Default is Null. The Font Bold settings see Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleCharFont(ByRef $oShapeStyle, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharFont($oShapeStyle, $sFontName, $nFontSize, $iPosture, $iWeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleCharFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleCharFontColor
; Description ...: Set or retrieve the font color and highlighting values for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleCharFontColor(ByRef $oShapeStyle[, $iFontColor = Null[, $iHighlight = Null]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iFontColor          - [optional] (-1-16777215) Default is Null. The font Color value, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iHighlight          - [optional] (-1-16777215) Default is Null. The highlight Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 3 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $FontColor
;                  |                               2 = Error setting $iHighlight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleCharFontColor(ByRef $oShapeStyle, $iFontColor = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_StyleCharFontColor($oShapeStyle, $iFontColor, $iHighlight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleCharFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleCharOverLine
; Description ...: Set and retrieve the OverLine settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleCharOverLine(ByRef $oShapeStyle[, $iOverLineStyle = Null[, $iOLColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iOverLineStyle      - [optional] (0-18) Default is Null. The style of the Overline line, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iOLColor            - [optional] (-1-16777215) Default is Null. The Overline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not Overlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleCharOverLine(ByRef $oShapeStyle, $iOverLineStyle = Null, $iOLColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharOverLine($oShapeStyle, $iOverLineStyle, $iOLColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleCharOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleCharStrikeOut
; Description ...: Set or Retrieve the Strikeout settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleCharStrikeOut(ByRef $oShapeStyle[, $iStrikeLineStyle = Null[, $bWordOnly = Null]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iStrikeLineStyle    - [optional] (0-6) Default is Null. The Strikeout Line Style, see constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bWordOnly           - [optional] Default is Null. If True, strike out is applied to words only, skipping whitespaces.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleCharStrikeOut(ByRef $oShapeStyle, $iStrikeLineStyle = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharStrikeOut($oShapeStyle, $iStrikeLineStyle, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleCharStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleCharUnderLine
; Description ...: Set and retrieve the Underline settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleCharUnderLine(ByRef $oShapeStyle[, $iUnderLineStyle = Null[, $iULColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iUnderLineStyle     - [optional] (0-18) Default is Null. The Underline line style, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iULColor            - [optional] (-1-16777215) Default is Null. The underline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not underlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle an Object.
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
Func _LOImpress_ShapeStyleCharUnderLine(ByRef $oShapeStyle, $iUnderLineStyle = Null, $iULColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharUnderLine($oShapeStyle, $iUnderLineStyle, $iULColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleCharUnderLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleConnectorSettings
; Description ...: Set or Retrieve Connector line settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleConnectorSettings(ByRef $oShapeStyle[, $iType = Null[, $iHoriBeg = Null[, $iHoriEnd = Null[, $iVertBeg = Null[, $iVertEnd = Null]]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iType               - [optional] (0-3) Default is Null. The connector line type. See Constants, $LOI_DRAWSHAPE_CONNECTOR_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iHoriBeg            - [optional] (0-10008) Default is Null. The amount of horizontal spacing, in Hundredths of a Millimeter (HMM), at the beginning of the connector.
;                  $iHoriEnd            - [optional] (0-10008) Default is Null. The amount of horizontal spacing, in Hundredths of a Millimeter (HMM), at the end of the connector.
;                  $iVertBeg            - [optional] (0-10008) Default is Null. The amount of vertical spacing, in Hundredths of a Millimeter (HMM), at the beginning of the connector.
;                  $iVertEnd            - [optional] (0-10008) Default is Null. The amount of vertical spacing, in Hundredths of a Millimeter (HMM), at the end of the connector.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iType not an Integer, less than 0 or greater than 3. See Constants, $LOI_DRAWSHAPE_CONNECTOR_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iHoriBeg not an Integer, less than 0 or greater than 10,008.
;                  @Error: 1, @Extended: 4 = $iHoriEnd not an Integer, less than 0 or greater than 10,008.
;                  @Error: 1, @Extended: 5 = $iVertBeg not an Integer, less than 0 or greater than 10,008.
;                  @Error: 1, @Extended: 6 = $iVertEnd not an Integer, less than 0 or greater than 10,008.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iType
;                  |                               2 = Error setting $iHoriBeg
;                  |                               4 = Error setting $iHoriEnd
;                  |                               8 = Error setting $iVertBeg
;                  |                               16 = Error setting $iVertEnd
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Line 1, 2, and 3 Skew setting is not available for Shape Styles.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleConnectorSettings(ByRef $oShapeStyle, $iType = Null, $iHoriBeg = Null, $iHoriEnd = Null, $iVertBeg = Null, $iVertEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avConnector[5]

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iType, $iHoriBeg, $iHoriEnd, $iVertBeg, $iVertEnd) Then
		__LO_ArrayFill($avConnector, $oShapeStyle.EdgeKind(), $oShapeStyle.EdgeNode1HorzDist(), $oShapeStyle.EdgeNode2HorzDist(), $oShapeStyle.EdgeNode1VertDist(), _
				$oShapeStyle.EdgeNode2VertDist())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avConnector)
	EndIf

	If ($iType <> Null) Then
		If Not __LO_IntIsBetween($iType, $LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD, $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oShapeStyle.EdgeKind = $iType
		$iError = ($oShapeStyle.EdgeKind() = $iType) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iHoriBeg <> Null) Then
		If Not __LO_IntIsBetween($iHoriBeg, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oShapeStyle.EdgeNode1HorzDist = $iHoriBeg
		$iError = (__LO_IntIsBetween($oShapeStyle.EdgeNode1HorzDist(), $iHoriBeg - 1, $iHoriBeg + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iHoriEnd <> Null) Then
		If Not __LO_IntIsBetween($iHoriEnd, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShapeStyle.EdgeNode2HorzDist = $iHoriEnd
		$iError = (__LO_IntIsBetween($oShapeStyle.EdgeNode2HorzDist(), $iHoriEnd - 1, $iHoriEnd + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iVertBeg <> Null) Then
		If Not __LO_IntIsBetween($iVertBeg, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oShapeStyle.EdgeNode1VertDist = $iVertBeg
		$iError = (__LO_IntIsBetween($oShapeStyle.EdgeNode1VertDist(), $iVertBeg - 1, $iVertBeg + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iVertEnd <> Null) Then
		If Not __LO_IntIsBetween($iVertEnd, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oShapeStyle.EdgeNode2VertDist = $iVertEnd
		$iError = (__LO_IntIsBetween($oShapeStyle.EdgeNode2VertDist(), $iVertEnd - 1, $iVertEnd + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeStyleConnectorSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleCreate
; Description ...: Create a new Drawing/Shape Style in a Document.
; Syntax ........: _LOImpress_ShapeStyleCreate(ByRef $oDoc, $sShapeStyle)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sShapeStyle         - The Name of the new Drawing/Shape Style to Create.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. New Shape Style successfully created. Returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sShapeStyle not a String.
;                  @Error: 1, @Extended: 3 = Shape Style name called in $sShapeStyle already exists in document.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating new Shape Style Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error Retrieving "graphics" style family Object.
;                  @Error: 3, @Extended: 2 = Error creating new Shape Style by name.
;                  @Error: 3, @Extended: 3 = Error Retrieving Created Shape Style Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_ShapeStyleDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleCreate(ByRef $oDoc, $sShapeStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShapeStyles, $oStyle, $oShapeStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If _LOImpress_ShapeStyleExists($oDoc, $sShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oShapeStyles = $oDoc.StyleFamilies.getByName("graphics")
	If Not IsObj($oShapeStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oStyle = $oShapeStyles.createInstance()
	If Not IsObj($oStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oShapeStyles.insertByName($sShapeStyle, $oStyle)

	If Not $oShapeStyles.hasByName($sShapeStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oShapeStyle = $oShapeStyles.getByName($sShapeStyle)
	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShapeStyle)
EndFunc   ;==>_LOImpress_ShapeStyleCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleCurrent
; Description ...: Set or Retrieve the current Drawing/Shape style for a Shape.
; Syntax ........: _LOImpress_ShapeStyleCurrent(ByRef $oDoc, ByRef $oShape[, $sShapeStyle = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function. See remarks.
;                  $sShapeStyle         - [optional] Default is Null. The Drawing/Shape Style name to set the Shape to. See remarks.
; Return values .: Success: 1 or String.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Drawing/Shape Style successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. All optional parameters were called with Null, returning current Drawing/Shape Style name set for the Shape.
;                  @Error: 0, @Extended: 2, Return: String = Success. All optional parameters were called with Null, returning current Presentation Style name set for the Shape.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oShape not an Object.
;                  @Error: 1, @Extended: 3 = $oShape does not support Shape Service.
;                  @Error: 1, @Extended: 4 = $sShapeStyle not a String.
;                  @Error: 1, @Extended: 5 = Drawing/Shape Style called in $sShapeStyle not found in Document.
;                  @Error: 1, @Extended: 6 = Can't set Drawing/Shape Style for Title/Subtitle/Outline text box, or other non-user created shapes.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Drawing/Shape Style name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sShapeStyle
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Because a Shape can have either a Presentation Style or a Drawing/Shape style applied to them, this function has two different @Extended values depending on whether the current style applied is a Presentation style or a Drawing/Shape style.
;                  You cannot set the style for a Presentation shape, which is a shape that is not user-created. These include Title, Subtitle, Outline and Note TextBoxes, also background shapes.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleCurrent(ByRef $oDoc, ByRef $oShape, $sShapeStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sCurrStyle
	Local $oShapeStyle
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oShape.supportsService("com.sun.star.presentation.Shape") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($sShapeStyle) Then
		$sCurrStyle = $oShape.Style.Name()
		If Not IsString($sCurrStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If $oShape.IsPresentationObject() Then ; Presentation Objects (Title, Subtitle, Outline textboxes etc) only have Presentation styles.

			Return SetError($__LO_STATUS_SUCCESS, 2, $sCurrStyle) ; Style is a Presentation Style.
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrStyle) ; Style is a Graphics/Drawing/Shape Style.
	EndIf

	If Not IsString($sShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not _LOImpress_ShapeStyleExists($oDoc, $sShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If $oShape.IsPresentationObject() Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; If this is a presentation object, this property is TRUE, Presentation objects are objects like TitleTextShape and OutlinerShape. Can't modify Presentation Objects.

	$oShapeStyle = $oDoc.StyleFamilies().getByName("graphics").getByName($sShapeStyle)
	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape.Style = $oShapeStyle
	$iError = ($oShape.Style.Name() = $oShapeStyle.Name()) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeStyleCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleDelete
; Description ...: Delete a User-Created Shape Style from a Document.
; Syntax ........: _LOImpress_ShapeStyleDelete(ByRef $oDoc, $oShapeStyle[, $bForceDelete = False[, $sReplacementStyle = "Standard"]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oShapeStyle         - A Drawing/Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function. Must be a User-Created Style, not a built-in Style native to LibreOffice.
;                  $bForceDelete        - [optional] Default is False. If True, Drawing/Shape style will be deleted regardless of whether it is in use or not.
;                  $sReplacementStyle   - [optional] Default is "standard". The Drawing/Shape style to use instead of the one being deleted if the Drawing/Shape style being deleted is applied to text in the document.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Drawing/Shape Style called in $oShapeStyle was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 3 = $oShapeStyle not a Drawing/Shape style Object.
;                  @Error: 1, @Extended: 4 = $bForceDelete not a Boolean.
;                  @Error: 1, @Extended: 5 = $sReplacementStyle not a String.
;                  @Error: 1, @Extended: 6 = Drawing/Shape Style called in $sReplacementStyle doesn't exist in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "graphics" style family Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Shape Style Name.
;                  @Error: 3, @Extended: 3 = $oShapeStyle is not a User-Created Shape Style and cannot be deleted.
;                  @Error: 3, @Extended: 4 = $oShapeStyle is in use and $bForceDelete is False.
;                  @Error: 3, @Extended: 5 = $oShapeStyle still exists after deletion attempt.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, _LOImpress_ShapeStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleDelete(ByRef $oDoc, ByRef $oShapeStyle, $bForceDelete = False, $sReplacementStyle = "standard")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShapeStyles
	Local $sShapeStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oShapeStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bForceDelete) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplacementStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($sReplacementStyle <> "") And Not _LOImpress_ShapeStyleExists($oDoc, $sReplacementStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oShapeStyles = $oDoc.StyleFamilies().getByName("graphics")
	If Not IsObj($oShapeStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sShapeStyle = $oShapeStyle.Name()
	If Not IsString($sShapeStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oShapeStyle.isUserDefined() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If $oShapeStyle.isInUse() And Not ($bForceDelete) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; If Style is in use return an error unless force delete is true.

	If ($oShapeStyle.getParentStyle() = Null) Or ($sReplacementStyle <> "Standard") Then $oShapeStyle.setParentStyle($sReplacementStyle)
	; If Parent style is blank set it to "Default Drawing Style" (Standard), Or if not but User has called a specific style set it to that.

	$oShapeStyles.removeByName($sShapeStyle)
	If $oShapeStyles.hasByName($sShapeStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$oShapeStyle = Null

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_ShapeStyleDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleDimensionSettings
; Description ...: Set or Retrieve Dimension line settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleDimensionSettings(ByRef $oShapeStyle[, $iDistance = Null[, $iGuideOverhang = Null[, $iGuideDistance = Null[, $iLGuide = Null[, $iRGuide = Null[, $bBelow = Null[, $iDecimal = Null[, $iVertPos = Null[, $iHoriPos = Null[, $bParallel = Null[, $iUnitType = Null]]]]]]]]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iDistance           - [optional] (-10008-10008) Default is Null. The distance between the dimension line and the baseline, in Hundredths of a Millimeter (HMM).
;                  $iGuideOverhang      - [optional] (-10008-10008) Default is Null. The length of the left and right guides starting at the baseline. Positive values extend the guides above the baseline and negative values extend the guides below the baseline, in Hundredths of a Millimeter (HMM).
;                  $iGuideDistance      - [optional] (-10008-10008) Default is Null. The length of the right and left guides starting at the dimension line. Positive values extend the guides above the dimension line and negative values extend the guides below the dimension line, in Hundredths of a Millimeter (HMM).
;                  $iLGuide             - [optional] (-10008-10008) Default is Null. The length of the left guide starting at the dimension line. Positive values extend the guide below the dimension line and negative values extend the guide above the dimension line, in Hundredths of a Millimeter (HMM).
;                  $iRGuide             - [optional] (-10008-10008) Default is Null. The length of the right guide starting at the dimension line. Positive values extend the guide below the dimension line and negative values extend the guide above the dimension line, in Hundredths of a Millimeter (HMM).
;                  $bBelow              - [optional] Default is Null. If True, the properties set in the Line area are Reversed.
;                  $iDecimal            - [optional] (0-99) Default is Null. The number of decimal places.
;                  $iVertPos            - [optional] (0-4) Default is Null. The position of the dimension line in reference to the text vertically. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iHoriPos            - [optional] (0-3) Default is Null. The position of the dimension text horizontally. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bParallel           - [optional] Default is Null. If True, Displays the text parallel to or at 90 degrees to the dimension line.
;                  $iUnitType           - [optional] (-1-15) Default is Null. The type of measurement units, if any, to display. See Constants, $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iDistance not an Integer, less than -10,008 or greater than 10,008.
;                  @Error: 1, @Extended: 3 = $iGuideOverhang not an Integer, less than -10,008 or greater than 10,008.
;                  @Error: 1, @Extended: 4 = $iGuideDistance not an Integer, less than -10,008 or greater than 10,008.
;                  @Error: 1, @Extended: 5 = $iLGuide not an Integer, less than -10,008 or greater than 10,008.
;                  @Error: 1, @Extended: 6 = $iRGuide not an Integer, less than -10,008 or greater than 10,008.
;                  @Error: 1, @Extended: 7 = $bBelow not a Boolean.
;                  @Error: 1, @Extended: 8 = $iDecimal not an Integer, less than 0 or greater than 99.
;                  @Error: 1, @Extended: 9 = $iVertPos not an Integer, less than 0 or greater than 4. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 10 = $iHoriPos not an Integer, less than 0 or greater than 3. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 11 = $bParallel not a Boolean.
;                  @Error: 1, @Extended: 12 = $iUnitType not an Integer, less than -1 or greater than 15. See Constants, $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
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
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleDimensionSettings(ByRef $oShapeStyle, $iDistance = Null, $iGuideOverhang = Null, $iGuideDistance = Null, $iLGuide = Null, $iRGuide = Null, $bBelow = Null, $iDecimal = Null, $iVertPos = Null, $iHoriPos = Null, $bParallel = Null, $iUnitType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_DimensionSettings($oShapeStyle, $iDistance, $iGuideOverhang, $iGuideDistance, $iLGuide, $iRGuide, $bBelow, $iDecimal, $iVertPos, $iHoriPos, $bParallel, $iUnitType)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleDimensionSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleExists
; Description ...: Check whether a Document contains a specific Drawing/Shape Style by name.
; Syntax ........: _LOImpress_ShapeStyleExists(ByRef $oDoc, $sShapeStyle)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sShapeStyle         - The Drawing/Shape Style Name to search for.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. If the Document contains the Drawing/Shape style called in $sShapeStyle, True is returned, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sShapeStyle not a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleExists(ByRef $oDoc, $sShapeStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oDoc.StyleFamilies.getByName("graphics").hasByName($sShapeStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOImpress_ShapeStyleExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleGetObjByName
; Description ...: Retrieve a Drawing/Shape Style Object for use with other ShapeStyle functions.
; Syntax ........: _LOImpress_ShapeStyleGetObjByName(ByRef $oDoc, $sShapeStyle)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sShapeStyle         - The Drawing/Shape Style name to retrieve the Object for.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Drawing/Shape Style successfully retrieved, returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sShapeStyle not a String.
;                  @Error: 1, @Extended: 3 = Drawing/Shape Style called in $sShapeStyle not found in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Drawing/Shape Style Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_ShapeStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleGetObjByName(ByRef $oDoc, $sShapeStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShapeStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOImpress_ShapeStyleExists($oDoc, $sShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oShapeStyle = $oDoc.StyleFamilies().getByName("graphics").getByName($sShapeStyle)
	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShapeStyle)
EndFunc   ;==>_LOImpress_ShapeStyleGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleLineArrowStyles
; Description ...: Set or Retrieve Shape Style Line Start and End Arrow Style settings.
; Syntax ........: _LOImpress_ShapeStyleLineArrowStyles(ByRef $oDoc, ByRef $oShapeStyle[, $vStartStyle = Null[, $iStartWidth = Null[, $bStartCenter = Null[, $bSync = Null[, $vEndStyle = Null[, $iEndWidth = Null[, $bEndCenter = Null]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $vStartStyle         - [optional] (0-32, or String) Default is Null. The Arrow head to apply to the start of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iStartWidth         - [optional] (0-5004) Default is Null. The Width of the Starting Arrowhead, in Hundredths of a Millimeter (HMM).
;                  $bStartCenter        - [optional] Default is Null. If True, Places the center of the Start arrowhead on the endpoint of the line.
;                  $bSync               - [optional] Default is Null. If True, Synchronizes the Start Arrowhead settings with the end Arrowhead settings. See remarks.
;                  $vEndStyle           - [optional] (0-32, or String) Default is Null. The Arrow head to apply to the end of the line. Can be a Custom Arrowhead name, or one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iEndWidth           - [optional] (0-5004) Default is Null. The Width of the Ending Arrowhead, in Hundredths of a Millimeter (HMM).
;                  $bEndCenter          - [optional] Default is Null. If True, Places the center of the End arrowhead on the endpoint of the line.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 3 = $vStartStyle not a String, and not an Integer.
;                  @Error: 1, @Extended: 4 = $vStartStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iStartWidth not an Integer, less than 0 or greater than 5004.
;                  @Error: 1, @Extended: 6 = $bStartCenter not a Boolean.
;                  @Error: 1, @Extended: 7 = $bSync not a Boolean.
;                  @Error: 1, @Extended: 8 = $vEndStyle not a String, and not an Integer.
;                  @Error: 1, @Extended: 9 = $vSEndStyle is an Integer, but less than 0 or greater than 32. See constants $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 10 = $iEndWidth not an Integer, less than 0 or greater than 5004.
;                  @Error: 1, @Extended: 11 = $bEndCenter not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to convert Constant to Arrowhead name.
;                  @Error: 3, @Extended: 2 = Failed to insert preset Arrowhead name and style.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $vStartStyle
;                  |                               2 = Error setting $iStartWidth
;                  |                               4 = Error setting $bStartCenter
;                  |                               8 = Error setting $bSync
;                  |                               16 = Error setting $vEndStyle
;                  |                               32 = Error setting $iEndWidth
;                  |                               64 = Error setting $bEndCenter
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When the arrowhead type "Arrow" is set in the LO UI, or upon creation of a line with arrows, the internal name of the arrowhead is set to an incrementing name of "Arrowheads x", where x is an Integer value. Since I have no way to determine if the head is a custom arrowhead or supposed to be the "Arrow" type, the return when this is present will be the name "Arrowheads x", and not $LOI_SHAPE_LINE_ARROW_TYPE_ARROW.
;                  When setting an Arrowhead to be $LOI_SHAPE_LINE_ARROW_TYPE_ARROW, the head is set correctly, but the LibreOffice UI will show "None". The return for Arrowhead type will be correct, $LOI_SHAPE_LINE_ARROW_TYPE_ARROW.
;                  LibreOffice has no setting for $bSync, so I have made a manual version of it in this function. It only accepts True, and must be called with True each time you want it to synchronize.
;                  When retrieving the current settings, $bSync will be a Boolean value of whether the Start Arrowhead settings are currently equal to the End Arrowhead setting values.
;                  Both $vStartStyle and $vEndStyle accept a String or an Integer because there is the possibility of a custom Arrowhead being available the user may want to use.
;                  When retrieving the current settings, both $vStartStyle and $vEndStyle could be either an Integer or a String. It will be a String if the current Arrowhead is a custom Arrowhead, else an Integer, corresponding to one of the constants, $LOI_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleLineArrowStyles(ByRef $oDoc, ByRef $oShapeStyle, $vStartStyle = Null, $iStartWidth = Null, $bStartCenter = Null, $bSync = Null, $vEndStyle = Null, $iEndWidth = Null, $bEndCenter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOImpress_ShapeStyleLineArrowStyles($oDoc, $oShapeStyle, $vStartStyle, $iStartWidth, $bStartCenter, $bSync, $vEndStyle, $iEndWidth, $bEndCenter)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleLineArrowStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleLineProperties
; Description ...: Set or Retrieve Shape Style Line settings.
; Syntax ........: _LOImpress_ShapeStyleLineProperties(ByRef $oDoc, ByRef $oShapeStyle[, $vStyle = Null[, $iColor = Null[, $iWidth = Null[, $iTransparency = Null[, $iCornerStyle = Null[, $iCapStyle = Null]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $vStyle              - [optional] (0-31, or String) Default is Null. The Line Style to use. Can be a Custom Line Style name, or one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3. See remarks.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Line color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iWidth              - [optional] (0-5004) Default is Null. The line Width, set in Hundredths of a Millimeter (HMM).
;                  $iTransparency       - [optional] (0-100) Default is Null. The Line transparency percentage. 100% = fully transparent.
;                  $iCornerStyle        - [optional] (0, 2-4) Default is Null. The Line Corner Style. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  $iCapStyle           - [optional] (0-2) Default is Null. The Line Cap Style. See Constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 3 = $vStyle not a String, and not an Integer.
;                  @Error: 1, @Extended: 4 = $vStyle is an Integer, but less than 0 or greater than 31. See constants $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 6 = $iWidth not an Integer, less than 0 or greater than 5004.
;                  @Error: 1, @Extended: 7 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iCornerStyle not an Integer, not equal to 0, equal to 1, not equal to 2 or greater than 4. See Constants $LOI_SHAPE_LINE_JOINT_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error: 1, @Extended: 9 = $iCapStyle is an Integer, but less than 0 or greater than 2. See constants $LOI_SHAPE_LINE_CAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to convert Constant to Line Style name.
;                  @Error: 3, @Extended: 2 = Failed to insert Line Style name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $vStyle
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $iWidth
;                  |                               8 = Error setting $iTransparency
;                  |                               16 = Error setting $iCornerStyle
;                  |                               32 = Error setting $iCapStyle
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $vStyle accepts a String or an Integer because there is the possibility of a custom Line Style being available that the user may want to use.
;                  When retrieving the current settings, $vStyle could be either an Integer or a String. It will be a String if the current Line Style is a custom Line Style, else an Integer, corresponding to one of the constants, $LOI_SHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleLineProperties(ByRef $oDoc, ByRef $oShapeStyle, $vStyle = Null, $iColor = Null, $iWidth = Null, $iTransparency = Null, $iCornerStyle = Null, $iCapStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOImpress_ShapeStyleLineProperties($oDoc, $oShapeStyle, $vStyle, $iColor, $iWidth, $iTransparency, $iCornerStyle, $iCapStyle)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleLineProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleOrganizer
; Description ...: Set or retrieve the Organizer settings of a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleOrganizer(ByRef $oDoc, $oShapeStyle[, $sNewShapeStyleName = Null[, $sParentStyle = Null[, $bHidden = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $sNewShapeStyleName  - [optional] Default is Null. The new name to set the Shape style called in $oShapeStyle to.
;                  $sParentStyle        - [optional] Default is Null. Set an existing Shape style (or an Empty String ("") = - None -) to apply its settings to the current style. Use the other settings to modify the inherited style settings.
;                  $bHidden             - [optional] Default is Null. If True, this style is hidden in the L.O. UI. Libre 4.0 and up only.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters. If The current LibreOffice version is below 4.0, the $bHidden parameter will return a Null value.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 3 = $oShapeStyle not a Shape Style Object.
;                  @Error: 1, @Extended: 4 = $sNewShapeStyleName not a String.
;                  @Error: 1, @Extended: 5 = Shape Style name called in $sNewShapeStyleName already exists in document.
;                  @Error: 1, @Extended: 6 = Cannot rename built-in Cell Styles.
;                  @Error: 1, @Extended: 7 = $sParentStyle not a String.
;                  @Error: 1, @Extended: 8 = Shape Style called in $sParentStyle doesn't exist in this Document.
;                  @Error: 1, @Extended: 9 = $bHidden not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sNewShapeStyleName
;                  |                               2 = Error setting $sParentStyle
;                  |                               4 = Error setting $bHidden
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 4.0.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOImpress_ShapeStyleCreate, _LOImpress_ShapeStyleGetObjByName, _LOImpress_ShapeStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleOrganizer(ByRef $oDoc, ByRef $oShapeStyle, $sNewShapeStyleName = Null, $sParentStyle = Null, $bHidden = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOrganizer[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oShapeStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($sNewShapeStyleName, $sParentStyle, $bHidden) Then
		If __LO_VersionCheck(4.0) Then
			__LO_ArrayFill($avOrganizer, $oShapeStyle.Name(), $oShapeStyle.ParentStyle(), $oShapeStyle.Hidden())

		Else
			__LO_ArrayFill($avOrganizer, $oShapeStyle.Name(), $oShapeStyle.ParentStyle(), Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOrganizer)
	EndIf

	If ($sNewShapeStyleName <> Null) Then
		If Not IsString($sNewShapeStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If _LOImpress_ShapeStyleExists($oDoc, $sNewShapeStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not $oShapeStyle.isUserDefined() Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oShapeStyle.Name = $sNewShapeStyleName
		$iError = (__LOImpress_ShapeStyleCompare($oDoc, $oShapeStyle.Name(), $sNewShapeStyleName)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sParentStyle <> Null) Then
		If Not IsString($sParentStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If ($sParentStyle <> "") And Not _LOImpress_ShapeStyleExists($oDoc, $sParentStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oShapeStyle.ParentStyle = $sParentStyle
		$iError = (__LOImpress_ShapeStyleCompare($oDoc, $oShapeStyle.ParentStyle(), $sParentStyle)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not __LO_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oShapeStyle.Hidden = $bHidden
		$iError = ($oShapeStyle.Hidden() = $bHidden) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_ShapeStyleOrganizer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleParAlignment
; Description ...: Set and Retrieve Paragraph Alignment settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleParAlignment(ByRef $oShapeStyle[, $iHorAlign = Null[, $iLastLineAlign = Null[, $iTxtDirection = Null]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iHorAlign           - [optional] (0-3) Default is Null. The Horizontal alignment of the paragraph. See Constants, $LOI_PAR_ALIGN_HOR_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iLastLineAlign      - [optional] (0-3) Default is Null. Specify the alignment for the last line in the paragraph. See Constants, $LOI_PAR_LAST_LINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iTxtDirection       - [optional] (0-5) Default is Null. The Text Writing Direction. See Constants, $LOI_PAR_TXT_DIR_* as defined in LibreOfficeImpress_Constants.au3. [LibreOffice Default is 4]
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleParAlignment(ByRef $oShapeStyle, $iHorAlign = Null, $iLastLineAlign = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParAlignment($oShapeStyle, $iHorAlign, $iLastLineAlign, $iTxtDirection)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleParAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleParIndent
; Description ...: Set or Retrieve Paragraph Indent settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleParIndent(ByRef $oShapeStyle[, $iBeforeTxt = Null[, $iAfterTxt = Null[, $iFirstLine = Null]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iBeforeTxt          - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM).
;                  $iAfterTxt           - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM)
;                  $iFirstLine          - [optional] (0-1162202) Default is Null. Indentation distance of the first line of a paragraph. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleParIndent(ByRef $oShapeStyle, $iBeforeTxt = Null, $iAfterTxt = Null, $iFirstLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParIndent($oShapeStyle, $iBeforeTxt, $iAfterTxt, $iFirstLine)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleParIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleParSpacing
; Description ...: Set and Retrieve Line Spacing settings for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleParSpacing(ByRef $oShapeStyle[, $iAbovePar = Null[, $iBelowPar = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iAbovePar           - [optional] (0-100000) Default is Null. The Space above a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iBelowPar           - [optional] (0-100000) Default is Null. The Space Below a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iLineSpcMode        - [optional] (0-3) Default is Null. The line spacing type of the paragraph. See Constants, $LOI_PAR_LINE_SPC_MODE_* as defined in LibreOfficeImpress_Constants.au3, also notice min and max values for each.
;                  $iLineSpcHeight      - [optional] Default is Null. This value specifies the height in regard to Mode. See Remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleParSpacing(ByRef $oShapeStyle, $iAbovePar = Null, $iBelowPar = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParSpacing($oShapeStyle, $iAbovePar, $iBelowPar, $iLineSpcMode, $iLineSpcHeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleParSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleParTabStopCreate
; Description ...: Create a new TabStop for a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleParTabStopCreate(ByRef $oShapeStyle, $iPosition[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iPosition           - The TabStop position to set the new TabStop to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Settings were successfully set. New TabStop position is returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleParTabStopCreate(ByRef $oShapeStyle, $iPosition, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopCreate($oShapeStyle, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleParTabStopCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleParTabStopDelete
; Description ...: Delete a TabStop from a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleParTabStopDelete(ByRef $oShapeStyle, $iTabStop)
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Returning True if TabStop was successfully deleted, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleParTabStopDelete(ByRef $oShapeStyle, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopDelete($oShapeStyle, $iTabStop)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleParTabStopDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop in a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleParTabStopMod(ByRef $oShapeStyle, $iTabStop[, $iPosition = Null[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - [optional] Default is Null. The New position to set the input position to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: ?, Return: 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
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
Func _LOImpress_ShapeStyleParTabStopMod(ByRef $oShapeStyle, $iTabStop, $iPosition = Null, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopMod($oShapeStyle, $iTabStop, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleParTabStopMod

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleParTabStopsGetList
; Description ...: Retrieve an array of TabStops available in a Shape Style.
; Syntax ........: _LOImpress_ShapeStyleParTabStopsGetList(ByRef $oShapeStyle)
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array of TabStops. @Extended set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleParTabStopsGetList(ByRef $oShapeStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopsGetList($oShapeStyle)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleParTabStopsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStylesGetNames
; Description ...: Retrieve an array of all Drawing/Shape Style names available for a document.
; Syntax ........: _LOImpress_ShapeStylesGetNames(ByRef $oDoc[, $bUserOnly = False[, $bAppliedOnly = False[, $bDisplayName = False]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bUserOnly           - [optional] Default is False. If True, only User-Created Drawing/Shape Styles are returned.
;                  $bAppliedOnly        - [optional] Default is False. If True, only Applied Drawing/Shape Styles are returned.
;                  $bDisplayName        - [optional] Default is False. If True, the style name displayed in the UI (Display Name), instead of the programmatic style name, is returned. See remarks.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array containing all Drawing/Shape Styles matching the called parameters. See remarks. @Extended contains the count of results returned.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bUserOnly not a Boolean.
;                  @Error: 1, @Extended: 3 = $bAppliedOnly not a Boolean.
;                  @Error: 1, @Extended: 4 = $bDisplayName not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Array of Drawing/Shape Style names.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If Only a Document object is called, all available Drawing/Shape styles will be returned.
;                  If Both $bUserOnly and $bAppliedOnly are called with True, only User-Created styles that are applied are returned.
;                  Two Drawing/Shape styles have different internal names:
;                  - "Default Drawing Style" is internally called "Standard".
;                  - "Object without fill" is internally called "objectwithoutfill".
;                  Previous to LibreOffice 25.2 either name would work when setting a Style, however after 25.2 only the internal, or programmatic style names, will work.
;                  Calling $bDisplayName with True will return a list of Style names, as the user sees them in the UI, in the same order as they are returned if $bDisplayName is False. It is best not to use these when setting Styling.
; Related .......: _LOImpress_ShapeStyleGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStylesGetNames(ByRef $oDoc, $bUserOnly = False, $bAppliedOnly = False, $bDisplayName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asStyles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDisplayName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$asStyles = __LO_StylesGetNames($oDoc, "graphics", $bUserOnly, $bAppliedOnly, $bDisplayName)
	If Not IsArray($asStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asStyles), $asStyles)
EndFunc   ;==>_LOImpress_ShapeStylesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleTextAttrAnimation
; Description ...: Set or Retrieve Shape Style Text Attribute Animation settings.
; Syntax ........: _LOImpress_ShapeStyleTextAttrAnimation(ByRef $oShapeStyle[, $iEffect = Null[, $iDirection = Null[, $bStartInside = Null[, $bVisibleOnExit = Null[, $iCycles = Null[, $iInc = Null[, $bPixels = Null[, $iDelay = Null]]]]]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iEffect             - [optional] (0-4) Default is Null. The Animation type. See Constants, $LOI_ANIMATION_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDirection          - [optional] (0-3) Default is Null. The Direction of the text's movement, if applicable. See Constants, $LOI_ANIMATION_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bStartInside        - [optional] Default is Null. If True, Text is visible and inside the shape when the effect is applied.
;                  $bVisibleOnExit      - [optional] Default is Null. If True, Text remains visible after the effect is applied.
;                  $iCycles             - [optional] (0-100) Default is Null. The number of times to repeat the animation. 0 = Continuous.
;                  $iInc                - [optional] (1-100px/25-32766) Default is Null. the increment value for scrolling the text, in Hundredths of a Millimeter (HMM), or pixels.
;                  $bPixels             - [optional] Default is Null. If True, $iInc is set in pixels, else in Hundredths of a Millimeter (HMM).
;                  $iDelay              - [optional] (0-30000) Default is Null. The amount time (ms) to wait before repeating the effect.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iEffect not an Integer, less than 0 or greater than 4. See Constants, $LOI_ANIMATION_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iDirection not an Integer, less than 0 or greater than 3. See Constants, $LOI_ANIMATION_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $bStartInside not a Boolean.
;                  @Error: 1, @Extended: 5 = $bVisibleOnExit not a Boolean.
;                  @Error: 1, @Extended: 6 = $iCycles not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iInc not an Integer, less than 1 or greater than 100 pixels, less than 25 or greater than 32766 Hundredths of a Millimeter (HMM).
;                  @Error: 1, @Extended: 8 = $bPixels not a Boolean.
;                  @Error: 1, @Extended: 9 = $iDelay not an Integer, less than 0 or greater than 30000.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current TextAnimationAmount value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iEffect
;                  |                               2 = Error setting $iDirection
;                  |                               4 = Error setting $bStartInside
;                  |                               8 = Error setting $bVisibleOnExit
;                  |                               16 = Error setting $iCycles
;                  |                               32 = Error setting $iInc
;                  |                               64 = Error setting $bPixels
;                  |                               128 = Error setting $iDelay
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleTextAttrAnimation(ByRef $oShapeStyle, $iEffect = Null, $iDirection = Null, $bStartInside = Null, $bVisibleOnExit = Null, $iCycles = Null, $iInc = Null, $bPixels = Null, $iDelay = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeTextAttrAnimation($oShapeStyle, $iEffect, $iDirection, $bStartInside, $bVisibleOnExit, $iCycles, $iInc, $bPixels, $iDelay)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleTextAttrAnimation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleTextAttrFit
; Description ...: Set or Retrieve Shape Style Text Attribute Fit properties.
; Syntax ........: _LOImpress_ShapeStyleTextAttrFit(ByRef $oShapeStyle[, $bFitWidth = Null[, $bFitHeight = Null[, $bFitToFrame = Null[, $bAdjustContour = Null[, $bWordWrap = Null[, $bResizeShape = Null]]]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $bFitWidth           - [optional] Default is Null. If True, Expands the width of the object to the width of the text.
;                  $bFitHeight          - [optional] Default is Null. If True, Expands the height of the object to the height of the text.
;                  $bFitToFrame         - [optional] Default is Null. If True, Resizes the text to fit the entire area of the drawing object.
;                  $bAdjustContour      - [optional] Default is Null. If True, Adapts the text flow so that it matches the contours of the drawing object.
;                  $bWordWrap           - [optional] Default is Null. If True, Wraps the text to fit inside the shape.
;                  $bResizeShape        - [optional] Default is Null. If True, Resizes a custom shape to fit the text that you enter.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $bFitWidth not a Boolean.
;                  @Error: 1, @Extended: 3 = $bFitHeight not a Boolean.
;                  @Error: 1, @Extended: 4 = $bFitToFrame not a Boolean.
;                  @Error: 1, @Extended: 5 = $bAdjustContour not a Boolean.
;                  @Error: 1, @Extended: 6 = $bWordWrap not a Boolean.
;                  @Error: 1, @Extended: 7 = $bResizeShape not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFitWidth
;                  |                               2 = Error setting $bFitHeight
;                  |                               4 = Error setting $bFitToFrame
;                  |                               8 = Error setting $bAdjustContour
;                  |                               16 = Error setting $bWordWrap
;                  |                               32 = Error setting $bResizeShape
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Properties as found in the UI, and their equivalent: "Word Wrap Text in Shape" = $bWordWrap. "Resize Shape to Fit Text" = $bResizeShape. "Fit Width to Text" = $bFitWidth. "Fit Height to Text" = $bFitHeight. "Fit to Frame" = $bFitToFrame. "Adjust to Contour" = $bAdjustContour.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleTextAttrFit(ByRef $oShapeStyle, $bFitWidth = Null, $bFitHeight = Null, $bFitToFrame = Null, $bAdjustContour = Null, $bWordWrap = Null, $bResizeShape = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeTextAttrFit($oShapeStyle, $bFitWidth, $bFitHeight, $bFitToFrame, $bAdjustContour, $bWordWrap, $bResizeShape)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleTextAttrFit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeStyleTextAttrSettings
; Description ...: Set or Retrieve Shape Style text Attribute settings.
; Syntax ........: _LOImpress_ShapeStyleTextAttrSettings(ByRef $oShapeStyle[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null[, $iAnchor = Null[, $bFullWidth = Null]]]]]])
; Parameters ....: $oShapeStyle         - A Shape Style object returned by a previous _LOImpress_ShapeStyleCreate, or _LOImpress_ShapeStyleGetObjByName function.
;                  $iLeft               - [optional] (-100000-100000) Default is Null. The space between the left edge of the drawing object and the left border of the text, in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] (-100000-100000) Default is Null. The space between the right edge of the drawing object and the right border of the text, in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] (-100000-100000) Default is Null. The space between the top edge of the drawing object and the top border of the text, in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] (-100000-100000) Default is Null. The space between the bottom edge of the drawing object and the bottom border of the text, in Hundredths of a Millimeter (HMM).
;                  $iAnchor             - [optional] (0-8) Default is Null. The text anchor position. See Constants, $LOI_PAR_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bFullWidth          - [optional] Default is Null. If True, Anchors the text to the full width of the drawing object.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShapeStyle not an Object.
;                  @Error: 1, @Extended: 2 = $iLeft not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 3 = $iRight not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 4 = $iTop not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 5 = $iBottom not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 6 = $iAnchor  not an Integer, less than 0 or greater than 8. See Constants, $LOI_PAR_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 7 = $bFullWidth not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
;                  |                               16 = Error setting $iAnchor
;                  |                               32 = Error setting $bFullWidth
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeStyleTextAttrSettings(ByRef $oShapeStyle, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null, $iAnchor = Null, $bFullWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShapeStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeTextAttrSettings($oShapeStyle, $iLeft, $iRight, $iTop, $iBottom, $iAnchor, $bFullWidth)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeStyleTextAttrSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeTextAttrAnimation
; Description ...: Set or Retrieve Shape Text Attribute Animation settings.
; Syntax ........: _LOImpress_ShapeTextAttrAnimation(ByRef $oShape[, $iEffect = Null[, $iDirection = Null[, $bStartInside = Null[, $bVisibleOnExit = Null[, $iCycles = Null[, $iInc = Null[, $bPixels = Null[, $iDelay = Null]]]]]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iEffect             - [optional] (0-4) Default is Null. The Animation type. See Constants, $LOI_ANIMATION_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDirection          - [optional] (0-3) Default is Null. The Direction of the text's movement, if applicable. See Constants, $LOI_ANIMATION_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bStartInside        - [optional] Default is Null. If True, Text is visible and inside the shape when the effect is applied.
;                  $bVisibleOnExit      - [optional] Default is Null. If True, Text remains visible after the effect is applied.
;                  $iCycles             - [optional] (0-100) Default is Null. The number of times to repeat the animation. 0 = Continuous.
;                  $iInc                - [optional] (1-100px/25-32766) Default is Null. the increment value for scrolling the text, in Hundredths of a Millimeter (HMM), or pixels.
;                  $bPixels             - [optional] Default is Null. If True, $iInc is set in pixels, else in Hundredths of a Millimeter (HMM).
;                  $iDelay              - [optional] (0-30000) Default is Null. The amount time (ms) to wait before repeating the effect.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $iEffect not an Integer, less than 0 or greater than 4. See Constants, $LOI_ANIMATION_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iDirection not an Integer, less than 0 or greater than 3. See Constants, $LOI_ANIMATION_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $bStartInside not a Boolean.
;                  @Error: 1, @Extended: 5 = $bVisibleOnExit not a Boolean.
;                  @Error: 1, @Extended: 6 = $iCycles not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iInc not an Integer, less than 1 or greater than 100 pixels, less than 25 or greater than 32766 Hundredths of a Millimeter (HMM).
;                  @Error: 1, @Extended: 8 = $bPixels not a Boolean.
;                  @Error: 1, @Extended: 9 = $iDelay not an Integer, less than 0 or greater than 30000.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current TextAnimationAmount value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iEffect
;                  |                               2 = Error setting $iDirection
;                  |                               4 = Error setting $bStartInside
;                  |                               8 = Error setting $bVisibleOnExit
;                  |                               16 = Error setting $iCycles
;                  |                               32 = Error setting $iInc
;                  |                               64 = Error setting $bPixels
;                  |                               128 = Error setting $iDelay
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeTextAttrAnimation(ByRef $oShape, $iEffect = Null, $iDirection = Null, $bStartInside = Null, $bVisibleOnExit = Null, $iCycles = Null, $iInc = Null, $bPixels = Null, $iDelay = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeTextAttrAnimation($oShape, $iEffect, $iDirection, $bStartInside, $bVisibleOnExit, $iCycles, $iInc, $bPixels, $iDelay)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeTextAttrAnimation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeTextAttrColumns
; Description ...: Set or Retrieve Shape Text Attribute Column settings. (L.O. 7.2+)
; Syntax ........: _LOImpress_ShapeTextAttrColumns(ByRef $oShape[, $iColumns = Null[, $iSpacing = Null]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iColumns            - [optional] (1-16) Default is Null. The number of columns.
;                  $iSpacing            - [optional] Default is Null. The spacing between each column, in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $iColumns not an Integer, less than 1 or greater than 16.
;                  @Error: 1, @Extended: 3 = $iSpacing not an Integer, less than 0.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.text.TextColumns Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve parent Document Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColumns
;                  |                               2 = Error setting $iSpacing
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current version is less than 7.2.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
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
; Description ...: Set or Retrieve Shape Text Attribute Fit properties.
; Syntax ........: _LOImpress_ShapeTextAttrFit(ByRef $oShape[, $bFitWidth = Null[, $bFitHeight = Null[, $bFitToFrame = Null[, $bAdjustContour = Null[, $bWordWrap = Null[, $bResizeShape = Null]]]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $bFitWidth           - [optional] Default is Null. If True, Expands the width of the object to the width of the text.
;                  $bFitHeight          - [optional] Default is Null. If True, Expands the height of the object to the height of the text.
;                  $bFitToFrame         - [optional] Default is Null. If True, Resizes the text to fit the entire area of the drawing object.
;                  $bAdjustContour      - [optional] Default is Null. If True, Adapts the text flow so that it matches the contours of the drawing object.
;                  $bWordWrap           - [optional] Default is Null. If True, Wraps the text to fit inside the shape.
;                  $bResizeShape        - [optional] Default is Null. If True, Resizes a custom shape to fit the text that you enter.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $bFitWidth not a Boolean.
;                  @Error: 1, @Extended: 3 = $bFitHeight not a Boolean.
;                  @Error: 1, @Extended: 4 = $bFitToFrame not a Boolean.
;                  @Error: 1, @Extended: 5 = $bAdjustContour not a Boolean.
;                  @Error: 1, @Extended: 6 = $bWordWrap not a Boolean.
;                  @Error: 1, @Extended: 7 = $bResizeShape not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFitWidth
;                  |                               2 = Error setting $bFitHeight
;                  |                               4 = Error setting $bFitToFrame
;                  |                               8 = Error setting $bAdjustContour
;                  |                               16 = Error setting $bWordWrap
;                  |                               32 = Error setting $bResizeShape
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
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeTextAttrFit(ByRef $oShape, $bFitWidth = Null, $bFitHeight = Null, $bFitToFrame = Null, $bAdjustContour = Null, $bWordWrap = Null, $bResizeShape = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeTextAttrFit($oShape, $bFitWidth, $bFitHeight, $bFitToFrame, $bAdjustContour, $bWordWrap, $bResizeShape)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeTextAttrFit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeTextAttrSettings
; Description ...: Set or Retrieve Shape text Attribute settings.
; Syntax ........: _LOImpress_ShapeTextAttrSettings(ByRef $oShape[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null[, $iAnchor = Null[, $bFullWidth = Null]]]]]])
; Parameters ....: $oShape              - A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iLeft               - [optional] (-100000-100000) Default is Null. The space between the left edge of the drawing object and the left border of the text, in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] (-100000-100000) Default is Null. The space between the right edge of the drawing object and the right border of the text, in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] (-100000-100000) Default is Null. The space between the top edge of the drawing object and the top border of the text, in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] (-100000-100000) Default is Null. The space between the bottom edge of the drawing object and the bottom border of the text, in Hundredths of a Millimeter (HMM).
;                  $iAnchor             - [optional] (0-8) Default is Null. The text anchor position. See Constants, $LOI_PAR_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bFullWidth          - [optional] Default is Null. If True, Anchors the text to the full width of the drawing object.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oShape not an Object.
;                  @Error: 1, @Extended: 2 = $iLeft not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 3 = $iRight not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 4 = $iTop not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 5 = $iBottom not an Integer, less than -100000 or greater than 100000.
;                  @Error: 1, @Extended: 6 = $iAnchor  not an Integer, less than 0 or greater than 8. See Constants, $LOI_PAR_TEXT_ANCHOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 7 = $bFullWidth not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
;                  |                               16 = Error setting $iAnchor
;                  |                               32 = Error setting $bFullWidth
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function will work, where applicable, for all drawing shapes, as well as other shapes that are returned by _LOImpress_ShapesGetList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeTextAttrSettings(ByRef $oShape, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null, $iAnchor = Null, $bFullWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeTextAttrSettings($oShape, $iLeft, $iRight, $iTop, $iBottom, $iAnchor, $bFullWidth)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_ShapeTextAttrSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ShapeTextBoxInsert
; Description ...: Create and Insert a Text box into a Slide.
; Syntax ........: _LOImpress_ShapeTextBoxInsert(ByRef $oSlide, $iTextBoxType, $iWidth, $iHeight[, $iX = -1[, $iY = -1]])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iTextBoxType        - (0-3) The type of Text Box to create. See Constants, $LOI_SHAPE_TEXTBOX_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iWidth              - The Text Box's Width in Hundredths of a Millimeter (HMM).
;                  $iHeight             - The Text Box's Height in Hundredths of a Millimeter (HMM).
;                  $iX                  - [optional] Default is -1. The X position from the top-left of the page, in Hundredths of a Millimeter (HMM). Call with -1 to center the Text Box horizontally.
;                  $iY                  - [optional] Default is -1. The Y position from the top-left of the page, in Hundredths of a Millimeter (HMM). Call with -1 to center the Text Box vertically.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Inserted a new Text Box. Returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iTextBoxType not an Integer, less than 0 or greater than 3. See Constants, $LOI_SHAPE_TEXTBOX_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 4 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 5 = $iX not an Integer.
;                  @Error: 1, @Extended: 6 = $iY not an Integer.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create the requested Text Box type.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve parent Document Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Position Structure.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Size Structure.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ShapeTextBoxInsert(ByRef $oSlide, $iTextBoxType, $iWidth, $iHeight, $iX = -1, $iY = -1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tSize, $tPos

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iTextBoxType, $LOI_SHAPE_TEXTBOX_TYPE_TEXTBOX, $LOI_SHAPE_TEXTBOX_TYPE_TITLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iTextBoxType
		Case $LOI_SHAPE_TEXTBOX_TYPE_TEXTBOX
			$oShape = $oDoc.createInstance("com.sun.star.drawing.TextShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_SHAPE_TEXTBOX_TYPE_OUTLINE
			$oShape = $oDoc.createInstance("com.sun.star.presentation.OutlinerShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_SHAPE_TEXTBOX_TYPE_SUBTITLE
			$oShape = $oDoc.createInstance("com.sun.star.presentation.SubtitleShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_SHAPE_TEXTBOX_TYPE_TITLE
			$oShape = $oDoc.createInstance("com.sun.star.presentation.TitleTextShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndSwitch

	$oSlide.add($oShape)

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tPos.X = ($iX = -1) ? (Int(($oSlide.Width() - $iWidth) / 2)) : ($iX)
	$tPos.Y = ($iY = -1) ? (Int(($oSlide.Height() - $iHeight) / 2)) : ($iY)

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>_LOImpress_ShapeTextBoxInsert
