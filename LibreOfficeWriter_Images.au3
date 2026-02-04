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
#include "LibreOfficeWriter_Page.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Inserting and Modifying Images in L.O. Writer.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_ImageAreaColor
; _LOWriter_ImageAreaFillStyle
; _LOWriter_ImageAreaGradient
; _LOWriter_ImageAreaGradientMulticolor
; _LOWriter_ImageAreaTransparency
; _LOWriter_ImageAreaTransparencyGradient
; _LOWriter_ImageAreaTransparencyGradientMulti
; _LOWriter_ImageBorderColor
; _LOWriter_ImageBorderPadding
; _LOWriter_ImageBorderStyle
; _LOWriter_ImageBorderWidth
; _LOWriter_ImageColorAdjust
; _LOWriter_ImageCrop
; _LOWriter_ImageDelete
; _LOWriter_ImageExists
; _LOWriter_ImageGetAnchor
; _LOWriter_ImageGetObjByName
; _LOWriter_ImageHyperlink
; _LOWriter_ImageInsert
; _LOWriter_ImageModify
; _LOWriter_ImageOptions
; _LOWriter_ImageOptionsName
; _LOWriter_ImageReplace
; _LOWriter_ImagesGetNames
; _LOWriter_ImageShadow
; _LOWriter_ImageSize
; _LOWriter_ImageTransparency
; _LOWriter_ImageTypePosition
; _LOWriter_ImageTypeSize
; _LOWriter_ImageWrap
; _LOWriter_ImageWrapOptions
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageAreaColor
; Description ...: Set or Retrieve background color settings for an Image.
; Syntax ........: _LOWriter_ImageAreaColor(ByRef $oImage[, $iBackColor = Null])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for "None".
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageAreaColor(ByRef $oImage, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency
	Local $iColor

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iBackColor) Then
		$iColor = __LOWriter_ColorRemoveAlpha($oImage.BackColor())
		If Not IsInt($iColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iOldTransparency = $oImage.FillTransparence()
	If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oImage.BackColor = $iBackColor
	$iError = ($oImage.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))

	$oImage.FillTransparence = $iOldTransparency

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOWriter_ImageAreaFillStyle(ByRef $oImage)
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
Func _LOWriter_ImageAreaFillStyle(ByRef $oImage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oImage.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOWriter_ImageAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageAreaGradient
; Description ...: Modify or retrieve the settings for an Image Background color Gradient.
; Syntax ........: _LOWriter_ImageAreaGradient(ByRef $oDoc, ByRef $oImage[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See Constants, $LOW_GRAD_NAME_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient type to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iIncrement          - [optional] an integer value (0,3-256). Default is Null. Specifies the number of steps of change color. 0 = Automatic.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] an integer value (0-16777215). Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] an integer value (0-16777215). Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] an integer value (0-100). Default is Null . Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sGradientName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageAreaGradient(ByRef $oDoc, ByRef $oImage, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName
	Local $atColorStop

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tStyleGradient = $oImage.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oImage.FillGradientName(), $tStyleGradient.Style(), _
				$oImage.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oImage.FillStyle() <> $LOW_AREA_FILL_STYLE_GRADIENT) Then $oImage.FillStyle = $LOW_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		__LOWriter_GradientPresets($oDoc, $oImage, $tStyleGradient, $sGradientName)
		$iError = ($oImage.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oImage.FillStyle = $LOW_AREA_FILL_STYLE_OFF
			$oImage.FillGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImage.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oImage.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
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

	If ($oImage.FillGradientName() = "") Or __LOWriter_GradientIsModified($tStyleGradient, $oImage.FillGradientName()) Then
		$sGradName = __LOWriter_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oImage.FillGradientName = $sGradName
		If ($oImage.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oImage.FillGradient = $tStyleGradient

	; Error checking
	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oImage.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oImage.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oImage.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oImage.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oImage.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = (__LO_VarsAreNull($iFromColor)) ? ($iError) : (($oImage.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = (__LO_VarsAreNull($iToColor)) ? ($iError) : (($oImage.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = (__LO_VarsAreNull($iFromIntense)) ? ($iError) : (($oImage.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = (__LO_VarsAreNull($iToIntense)) ? ($iError) : (($oImage.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageAreaGradientMulticolor
; Description ...: Set or Retrieve an Image's Multicolor Gradient settings. See remarks.
; Syntax ........: _LOWriter_ImageAreaGradientMulticolor(ByRef $oImage[, $avColorStops = Null])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
; Related .......: _LOWriter_GradientMulticolorAdd, _LOWriter_GradientMulticolorDelete, _LOWriter_GradientMulticolorModify, _LOWriter_ImageAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageAreaGradientMulticolor(ByRef $oImage, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oImage.FillGradient()
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
	$oImage.FillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oImage.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageAreaTransparency
; Description ...: Modify or retrieve Transparency settings for an Image's background color.
; Syntax ........: _LOWriter_ImageAreaTransparency(ByRef $oImage[, $iTransparency = Null])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageAreaTransparency(ByRef $oImage, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oImage.FillTransparence())

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oImage.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oImage.FillTransparence = $iTransparency
	$iError = ($oImage.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageAreaTransparencyGradient
; Description ...: Modify or retrieve the Image's background transparency gradient settings.
; Syntax ........: _LOWriter_ImageAreaTransparencyGradient(ByRef $oDoc, ByRef $oImage[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient that you want to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Call with $LOW_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error 1 @Extended 2 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iType not an Integer, less than -1 or greater than 5. See constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageAreaTransparencyGradient(ByRef $oDoc, ByRef $oImage, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $sTGradName
	Local $aiTransparent[7]
	Local $atColorStop
	Local $fValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tStyleGradient = $oImage.FillTransparenceGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		__LO_ArrayFill($aiTransparent, $tStyleGradient.Style(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), _
				Int($tStyleGradient.Angle() / 10), $tStyleGradient.Border(), __LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.StartColor()), _
				__LOWriter_TransparencyGradientConvert(Null, $tStyleGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oImage.FillTransparenceGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iStart <> Null) Then
		If Not __LO_IntIsBetween($iStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

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
		If Not __LO_IntIsBetween($iEnd, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

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

	If ($oImage.FillTransparenceGradientName() = "") Then
		$sTGradName = __LOWriter_TransparencyGradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oImage.FillTransparenceGradientName = $sTGradName
		If ($oImage.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oImage.FillTransparenceGradient = $tStyleGradient

	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oImage.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oImage.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oImage.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oImage.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oImage.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iStart)) ? ($iError) : (($oImage.FillTransparenceGradient.StartColor() = __LOWriter_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iEnd)) ? ($iError) : (($oImage.FillTransparenceGradient.EndColor() = __LOWriter_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageAreaTransparencyGradientMulti
; Description ...: Set or Retrieve an Image's Multi Transparency Gradient settings. See remarks.
; Syntax ........: _LOWriter_ImageAreaTransparencyGradientMulti(ByRef $oImage[, $avColorStops = Null])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
; Related .......: _LOWriter_TransparencyGradientMultiModify, _LOWriter_TransparencyGradientMultiDelete, _LOWriter_TransparencyGradientMultiAdd, _LOWriter_ImageAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageAreaTransparencyGradientMulti(ByRef $oImage, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oImage.FillTransparenceGradient()
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
	$oImage.FillTransparenceGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oImage.FillTransparenceGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageBorderColor
; Description ...: Set or retrieve the Image Border Line Color. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_ImageBorderColor(ByRef $oImage[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. The Top Border Line Color of the Image, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. The Bottom Border Line Color of the Image, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. The Left Border Line Color of the Image, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. The Right Border Line Color of the Image, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOWriter_ImageBorderWidth, _LOWriter_ImageBorderStyle, _LOWriter_ImageBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageBorderColor(ByRef $oImage, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oImage, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ImageBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageBorderPadding
; Description ...: Set or retrieve the Image Border Padding settings.
; Syntax ........: _LOWriter_ImageBorderPadding(ByRef $oImage[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The Top Distance between the Border and Image in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Distance between the Border and Image in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Distance between the Border and Image in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The Right Distance between the Border and Image in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_UnitConvert, _LOWriter_ImageBorderWidth, _LOWriter_ImageBorderStyle, _LOWriter_ImageBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageBorderPadding(ByRef $oImage, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oImage.BorderDistance(), $oImage.TopBorderDistance(), _
				$oImage.BottomBorderDistance(), $oImage.LeftBorderDistance(), $oImage.RightBorderDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LO_IntIsBetween($iAll, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oImage.BorderDistance = $iAll
		$iError = (__LO_IntIsBetween($oImage.BorderDistance(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.TopBorderDistance = $iTop
		$iError = (__LO_IntIsBetween($oImage.TopBorderDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.BottomBorderDistance = $iBottom
		$iError = (__LO_IntIsBetween($oImage.BottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImage.LeftBorderDistance = $iLeft
		$iError = (__LO_IntIsBetween($oImage.LeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImage.RightBorderDistance = $iRight
		$iError = (__LO_IntIsBetween($oImage.RightBorderDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageBorderStyle
; Description ...: Set or Retrieve the Image Border Line style. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_ImageBorderStyle(ByRef $oImage[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. The Top Border Line Style of the Image. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. The Bottom Border Line Style of the Image. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. The Left Border Line Style of the Image. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. The Right Border Line Style of the Image. See Constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LOWriter_ImageBorderWidth, _LOWriter_ImageBorderColor, _LOWriter_ImageBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageBorderStyle(ByRef $oImage, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oImage, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ImageBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageBorderWidth
; Description ...: Set or Retrieve the Image Border Line Width. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_ImageBorderWidth(ByRef $oImage[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iTop                - [optional] an integer value. Default is Null. The Top Border Line width of the Image in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value. Default is Null. The Bottom Border Line Width of the Image in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The Left Border Line width of the Image in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value. Default is Null. The Right Border Line Width of the Image in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_UnitConvert, _LOWriter_ImageBorderStyle, _LOWriter_ImageBorderColor, _LOWriter_ImageBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageBorderWidth(ByRef $oImage, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oImage, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ImageBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageColorAdjust
; Description ...: Set or retrieve Image color adjustment settings.
; Syntax ........: _LOWriter_ImageColorAdjust(ByRef $oImage[, $iRed = Null[, $iGreen = Null[, $iBlue = Null[, $iBrightness = Null[, $iContrast = Null[, $nGamma = Null[, $iColorMode = Null[, $bInvert = Null]]]]]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iRed                - [optional] an integer value (-100-100). Default is Null. Changes the display of the Red color channel. As a percentage.
;                  $iGreen              - [optional] an integer value (-100-100). Default is Null. Changes the display of the Green color channel. As a percentage.
;                  $iBlue               - [optional] an integer value (-100-100). Default is Null. Changes the display of the Blue color channel. As a percentage.
;                  $iBrightness         - [optional] an integer value (-100-100). Default is Null. Adjust the brightness of the graphic.
;                  $iContrast           - [optional] an integer value (-100-100). Default is Null. Adjust the contrast of the graphic.
;                  $nGamma              - [optional] a general number value (0.1-10). Default is Null. The gamma value of the graphic.
;                  $iColorMode          - [optional] an integer value (0-3). Default is Null. The color mode of the graphic. See constants, $LOW_COLORMODE_* as defined in LibreOfficeWriter_Constants.au3
;                  $bInvert             - [optional] a boolean value. Default is Null. If True, the graphic is displayed in inverted colors. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iRed not an Integer, less than -100 or greater than 100.
;                  @Error 1 @Extended 3 Return 0 = $iGreen not an Integer, less than -100 or greater than 100.
;                  @Error 1 @Extended 4 Return 0 = $iBlue not an Integer, less than -100 or greater than 100.
;                  @Error 1 @Extended 5 Return 0 = $iBrightness not an Integer, less than -100 or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iContrast not an Integer, less than -100 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $nGamma not a number, less than 0.1, or greater than 10.
;                  @Error 1 @Extended 8 Return 0 = $iColorMode not an Integer, less than 0 or greater than 3. See constants, $LOW_COLORMODE_* as defined in LibreOfficeWriter_Constants.au3
;                  @Error 1 @Extended 9 Return 0 = $bInvert not a boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRed
;                  |                               2 = Error setting $iGreen
;                  |                               4 = Error setting $iBlue
;                  |                               8 = Error setting $iBrightness
;                  |                               16 = Error setting $iContrast
;                  |                               32 = Error setting $nGamma
;                  |                               64 = Error setting $iColorMode
;                  |                               128 = Error setting $bInvert
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bInvert is glitchy to set. The current setting will always be returned as False if set by the user. Setting inverted using this function can be difficult to remove by the user.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageColorAdjust(ByRef $oImage, $iRed = Null, $iGreen = Null, $iBlue = Null, $iBrightness = Null, $iContrast = Null, $nGamma = Null, $iColorMode = Null, $bInvert = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avImage[8]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iRed, $iGreen, $iBlue, $iBrightness, $iContrast, $nGamma, $iColorMode, $bInvert) Then
		__LO_ArrayFill($avImage, $oImage.AdjustRed(), $oImage.AdjustGreen(), $oImage.AdjustBlue(), $oImage.AdjustLuminance(), _
				$oImage.AdjustContrast(), $oImage.Gamma(), $oImage.GraphicColorMode(), $oImage.GraphicIsInverted())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avImage)
	EndIf

	If ($iRed <> Null) Then
		If Not __LO_IntIsBetween($iRed, -100, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oImage.AdjustRed = $iRed
		$iError = ($oImage.AdjustRed() = $iRed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iGreen <> Null) Then
		If Not __LO_IntIsBetween($iGreen, -100, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.AdjustGreen = $iGreen
		$iError = ($oImage.AdjustGreen() = $iGreen) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iBlue <> Null) Then
		If Not __LO_IntIsBetween($iBlue, -100, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.AdjustBlue = $iBlue
		$iError = ($oImage.AdjustBlue() = $iBlue) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iBrightness <> Null) Then
		If Not __LO_IntIsBetween($iBrightness, -100, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImage.AdjustLuminance = $iBrightness
		$iError = ($oImage.AdjustLuminance() = $iBrightness) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iContrast <> Null) Then
		If Not __LO_IntIsBetween($iContrast, -100, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImage.AdjustContrast = $iContrast
		$iError = ($oImage.AdjustContrast() = $iContrast) ? ($iError) : (BitOR($iError, 16))
	EndIf

	; Min 0.1, Max 10
	If ($nGamma <> Null) Then
		If Not __LO_NumIsBetween($nGamma, 0.1, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oImage.Gamma = $nGamma
		$iError = ($oImage.Gamma() = $nGamma) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iColorMode <> Null) Then
		If Not __LO_IntIsBetween($iColorMode, $LOW_COLORMODE_STANDARD, $LOW_COLORMODE_WATERMARK) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oImage.GraphicColorMode = $iColorMode
		$iError = ($oImage.GraphicColorMode() = $iColorMode) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bInvert <> Null) Then
		If Not IsBool($bInvert) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oImage.GraphicIsInverted = $bInvert
		$iError = ($oImage.GraphicIsInverted() = $bInvert) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageColorAdjust

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageCrop
; Description ...: Set or retrieve Image crop settings.
; Syntax ........: _LOWriter_ImageCrop(ByRef $oImage[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null[, $bKeepScale = Null]]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iLeft               - [optional] an integer value. Default is Null. The amount in Hundredths of a Millimeter (HMM) to either extend the background of the image, (negative numbers), or to crop, (positive numbers) from the Left side.
;                  $iRight              - [optional] an integer value. Default is Null. The amount in Hundredths of a Millimeter (HMM) to either extend the background of the image, (negative numbers), or to crop, (positive numbers) from the Right side.
;                  $iTop                - [optional] an integer value. Default is Null. The amount in Hundredths of a Millimeter (HMM) to either extend the background of the image, (negative numbers), or to crop, (positive numbers) from the Top side.
;                  $iBottom             - [optional] an integer value. Default is Null. The amount in Hundredths of a Millimeter (HMM) to either extend the background of the image, (negative numbers), or to crop, (positive numbers) from the Bottom side.
;                  $bKeepScale          - [optional] a boolean value. Default is Null. If True, crop amounts are removed or added to the image, while keeping the scaling. If False, crop values are removed or added while retaining the image size. See remarks. This setting is internally static, you do not need to set this each call for as long as the script life, unless you wish to change the value. Default static setting is True.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bKeepScale not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $iLeft not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iRight not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the image Crop structure.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the image Size structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: There is no literal setting for $bKeepScale in Libre Office's settings, so I have made an internal static setting in this function to behave the same as Libre Office. When you retrieve the current settings for an image, the return for $bKeepScale will be my internal static value, and NOT the current LibreOffice setting.
;                  Maximum crop values are based on page width. You cannot exceed the size of the page, nor crop too much of the image away.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageCrop(ByRef $oImage, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null, $bKeepScale = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avImage[5]
	Local $tCrop, $tSize
	Local Static $bKeepScaleInternal = True

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($bKeepScale <> Null) And Not IsBool($bKeepScale) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$bKeepScaleInternal = ($bKeepScale = Null) ? ($bKeepScaleInternal) : ($bKeepScale)

	$tCrop = $oImage.GraphicCrop()
	If Not IsObj($tCrop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tSize = $oImage.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iLeft, $iRight, $iTop, $iBottom, $bKeepScale) Then
		__LO_ArrayFill($avImage, $tCrop.Left(), $tCrop.Right(), $tCrop.Top(), $tCrop.Bottom(), $bKeepScaleInternal)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avImage)
	EndIf

	If ($iLeft <> Null) Then
		If Not IsInt($iLeft) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If ($bKeepScaleInternal = True) Then $tSize.Width = ($tSize.Width() + $tCrop.Left() - $iLeft)
		$tCrop.Left = $iLeft
	EndIf

	If ($iRight <> Null) Then
		If Not IsInt($iRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If ($bKeepScaleInternal = True) Then $tSize.Width = ($tSize.Width() + $tCrop.Right() - $iRight)
		$tCrop.Right = $iRight
	EndIf

	If ($iTop <> Null) Then
		If Not IsInt($iTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		If ($bKeepScaleInternal = True) Then $tSize.Height = ($tSize.Height() + $tCrop.Top() - $iTop)
		$tCrop.Top = $iTop
	EndIf

	If ($iBottom <> Null) Then
		If Not IsInt($iBottom) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		If ($bKeepScaleInternal = True) Then $tSize.Height = ($tSize.Height() + $tCrop.Bottom() - $iBottom)
		$tCrop.Bottom = $iBottom
	EndIf

	$oImage.GraphicCrop = $tCrop

	If ($bKeepScaleInternal = True) Then $oImage.Size = $tSize

	; Error checking
	$iError = (__LO_VarsAreNull($iLeft)) ? ($iError) : ((__LO_IntIsBetween($oImage.GraphicCrop.Left(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iRight)) ? ($iError) : ((__LO_IntIsBetween($oImage.GraphicCrop.Right(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iTop)) ? ($iError) : ((__LO_IntIsBetween($oImage.GraphicCrop.Top(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iBottom)) ? ($iError) : ((__LO_IntIsBetween($oImage.GraphicCrop.Bottom(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageCrop

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageDelete
; Description ...: Delete an Image from the document.
; Syntax ........: _LOWriter_ImageDelete(ByRef $oDoc, ByRef $oImage)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oImage not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to delete image.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Image was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageDelete(ByRef $oDoc, ByRef $oImage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sImageName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sImageName = $oImage.getName()
	$oImage.dispose()
	If ($oDoc.GraphicObjects().hasByName($sImageName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Document still contains Image named the same.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_ImageDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageExists
; Description ...: Check if a Document contains an Image with the specified name.
; Syntax ........: _LOWriter_ImageExists(ByRef $oDoc, $sImageName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sImageName          - a string value. The Image name to search for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sImageName not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Search was successful, If an Image was found matching $sImageName, True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ImageDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageExists(ByRef $oDoc, $sImageName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sImageName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($oDoc.GraphicObjects().hasByName($sImageName)) Then Return SetError($__LO_STATUS_SUCCESS, 1, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; No matches
EndFunc   ;==>_LOWriter_ImageExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageGetAnchor
; Description ...: Create a Text Cursor at the Image Anchor position.
; Syntax ........: _LOWriter_ImageGetAnchor(ByRef $oImage)
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to retrieve Image anchor Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully returned the Image Anchor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageGetAnchor(ByRef $oImage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnchor

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnchor = $oImage.Anchor.Text.createTextCursorByRange($oImage.Anchor())
	If Not IsObj($oAnchor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oAnchor)
EndFunc   ;==>_LOWriter_ImageGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageGetObjByName
; Description ...: Retrieve an Image's Object by name from a document.
; Syntax ........: _LOWriter_ImageGetObjByName(ByRef $oDoc, $sImage)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sImage              - a string value. The Image name to retrieve the Object for.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sImage not a string.
;                  @Error 1 @Extended 3 Return 0 = Image name called in $sImage not found in document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve GraphicObjects object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve requested Image object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully found requested Image by name, returning Image Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ImagesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageGetObjByName(ByRef $oDoc, $sImage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oImage, $oImages

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oImages = $oDoc.GraphicObjects()
	If Not IsObj($oImages) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oImages.hasByName($sImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oImage = $oImages.getByName($sImage)
	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oImage)
EndFunc   ;==>_LOWriter_ImageGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageHyperlink
; Description ...: Set or Retrieve Image Hyperlink settings.
; Syntax ........: _LOWriter_ImageHyperlink(ByRef $oImage[, $sURL = Null[, $sName = Null[, $sFrameTarget = Null[, $bServerSideMap = Null]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $sURL                - [optional] a string value. Default is Null. The complete path to the file that you want to open.
;                  $sName               - [optional] a string value. Default is Null. Name for the hyperlink.
;                  $sFrameTarget        - [optional] a string value. Default is Null. Specify the name of the frame where you want to open the targeted file. See Constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bServerSideMap      - [optional] a boolean value. Default is Null. If True, Uses a server-side image map.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sURL not a String
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sFrameTarget not a String.
;                  @Error 1 @Extended 5 Return 0 = $sFrameTarget not equal to one of the Constants. See constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageHyperlink(ByRef $oImage, $sURL = Null, $sName = Null, $sFrameTarget = Null, $bServerSideMap = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHyperlink[4]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sURL, $sName, $sFrameTarget, $bServerSideMap) Then
		__LO_ArrayFill($avHyperlink, $oImage.HyperLinkURL(), $oImage.HyperLinkName(), $oImage.HyperLinkTarget(), $oImage.ServerMap())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avHyperlink)
	EndIf

	If ($sURL <> Null) Then
		If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oImage.HyperLinkURL = $sURL
		$iError = ($oImage.HyperLinkURL() = $sURL) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.HyperLinkName = $sName
		$iError = ($oImage.HyperLinkName = $sName) ? ($iError) : (BitOR($iError, 2))
	EndIf
	;"" ; "_top" ; "_parent" ; "_blank" ; "_self"

	If ($sFrameTarget <> Null) Then
		If Not IsString($sFrameTarget) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If ($sFrameTarget <> "") Then
			If ($sFrameTarget <> $LOW_FRAME_TARGET_TOP) And _
					($sFrameTarget <> $LOW_FRAME_TARGET_PARENT) And _
					($sFrameTarget <> $LOW_FRAME_TARGET_BLANK) And _
					($sFrameTarget <> $LOW_FRAME_TARGET_SELF) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		EndIf
		$oImage.HyperLinkTarget = $sFrameTarget
		$iError = ($oImage.HyperLinkTarget() = $sFrameTarget) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bServerSideMap <> Null) Then
		If Not IsBool($bServerSideMap) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImage.ServerMap = $bServerSideMap
		$iError = ($oImage.ServerMap() = $bServerSideMap) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageHyperlink

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageInsert
; Description ...: Insert an image into a document.
; Syntax ........: _LOWriter_ImageInsert(ByRef $oDoc, $sImage, ByRef $oCursor[, $iAnchorType = $LOW_ANCHOR_AT_CHARACTER[, $bOverwrite = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sImage              - a string value. The file path to the image to insert.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions. See Remarks.
;                  $iAnchorType         - [optional] an integer value (0-2,4). Default is $LOW_ANCHOR_AT_CHARACTER. Specify the anchoring options for the Image. See Constants, $LOW_ANCHOR_* as defined in LibreOfficeWriter_Constants.au3
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any data selected by the cursor is overwritten.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sImage not a String.
;                  @Error 1 @Extended 3 Return 0 = $oCursor not an Object. And not called with Default
;                  @Error 1 @Extended 4 Return 0 = $iAnchorType not an Integer, less than 0 or greater than 2, but not equal to 4. See Constants, $LOW_ANCHOR_* as defined in LibreOfficeWriter_Constants.au3
;                  @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $oCursor is a Table Cursor and is not supported.
;                  @Error 1 @Extended 7 Return 0 = Image called in $sImage doesn't exist at given path.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failure creating "com.sun.star.text.TextGraphicObject" Object.
;                  @Error 2 @Extended 2 Return 0 = Failure creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 3 Return 0 = Failure Creating "com.sun.star.graphic.GraphicProvider" Object.
;                  @Error 2 @Extended 4 Return 0 = Error creating a property value for retrieving the Image's size.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error getting Cursor type.
;                  @Error 3 @Extended 2 Return 0 = Error converting Image Path to Libre Office URL.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving current Page Style name at insertion point.
;                  @Error 3 @Extended 4 Return 0 = Error retrieving Page Style Object.
;                  @Error 3 @Extended 5 Return 0 = Error calculating suggested image size.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Image was successfully inserted, returning image Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Unfortunately, I am unable to find a way to insert an image "linked", images can only be inserted as embedded.
; Related .......: _LOWriter_ImageDelete, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_FrameCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageInsert(ByRef $oDoc, $sImage, ByRef $oCursor, $iAnchorType = $LOW_ANCHOR_AT_CHARACTER, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $oImage
	Local $oServiceManager, $oProvider, $oSize, $oPageStyle
	Local $sPageStyle
	Local $atProp[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iAnchorType, $LOW_ANCHOR_AT_PARAGRAPH, $LOW_ANCHOR_AT_CHARACTER) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not FileExists($sImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$sImage = _LO_PathConvert($sImage, $LO_PATHCONV_OFFICE_RETURN)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oImage = $oDoc.createInstance("com.sun.star.text.TextGraphicObject")
	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oProvider = $oServiceManager.createInstance("com.sun.star.graphic.GraphicProvider")
	If Not IsObj($oProvider) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	$atProp[0] = __LO_SetPropertyValue("URL", $sImage)
	If (@error > 0) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$sPageStyle = $oCursor.PageStyleName()
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, $sPageStyle)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$oSize = __LOWriter_ImageGetSuggestedSize(($oProvider.queryGraphicDescriptor($atProp)), $oPageStyle)
	If Not IsObj($oSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	With $oImage
		.GraphicURL = $sImage
		.AnchorType = $iAnchorType
		.Width = $oSize.Width()
		.Height = $oSize.Height()
	EndWith

	$oCursor.Text.insertTextContent($oCursor, $oImage, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oImage)
EndFunc   ;==>_LOWriter_ImageInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageModify
; Description ...: Set or retrieve Image modification settings.
; Syntax ........: _LOWriter_ImageModify(ByRef $oImage[, $bFlipVert = Null[, $bFlipHoriOnRight = Null[, $bFlipHoriOnLeft = Null[, $nAngle = Null]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $bFlipVert           - [optional] a boolean value. Default is Null. If True, the image is flipped vertically.
;                  $bFlipHoriOnRight    - [optional] a boolean value. Default is Null. If True, the image is flipped horizontally on right (odd) pages. Set both this and $bFlipHoriOnLeft to True to flip on all pages.
;                  $bFlipHoriOnLeft     - [optional] a boolean value. Default is Null. If True, the image is flipped horizontally on left (even) pages. Set both this and $bFlipHoriOnRight to True to flip on all pages.
;                  $nAngle              - [optional] a floating point value (0-360). Default is Null. The angle to rotate the image.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFlipVert not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bFlipHoriOnRight not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bFlipHoriOnLeft not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $nAngle not a number, less than 0 or greater than 360.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFlipVert
;                  |                               2 = Error setting $bFlipHoriOnRight
;                  |                               4 = Error setting $bFlipHoriOnLeft
;                  |                               8 = Error setting $nAngle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Unfortunately I cannot find a way to replace an image as a linked image. Thus I have skipped "Link" setting.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageModify(ByRef $oImage, $bFlipVert = Null, $bFlipHoriOnRight = Null, $bFlipHoriOnLeft = Null, $nAngle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avImage[4]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bFlipVert, $bFlipHoriOnRight, $bFlipHoriOnLeft, $nAngle) Then
		__LO_ArrayFill($avImage, $oImage.VertMirrored(), $oImage.HoriMirroredOnEvenPages(), $oImage.HoriMirroredOnOddPages(), _
				($oImage.GraphicRotation()) / 10) ; /10 to match L.O. values.

		Return SetError($__LO_STATUS_SUCCESS, 1, $avImage)
	EndIf

	If ($bFlipVert <> Null) Then
		If Not IsBool($bFlipVert) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oImage.VertMirrored = $bFlipVert
		$iError = ($oImage.VertMirrored() = $bFlipVert) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bFlipHoriOnRight <> Null) Then
		If Not IsBool($bFlipHoriOnRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.HoriMirroredOnEvenPages = $bFlipHoriOnRight
		$iError = ($oImage.HoriMirroredOnEvenPages() = $bFlipHoriOnRight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bFlipHoriOnLeft <> Null) Then
		If Not IsBool($bFlipHoriOnLeft) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.HoriMirroredOnOddPages = $bFlipHoriOnLeft
		$iError = ($oImage.HoriMirroredOnOddPages() = $bFlipHoriOnLeft) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($nAngle <> Null) Then
		If Not __LO_NumIsBetween($nAngle, 0, 360) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImage.GraphicRotation = ($nAngle * 10) ; x10 to match L.O. Values
		$iError = ($oImage.GraphicRotation() = ($nAngle * 10)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageOptions
; Description ...: Set or Retrieve Image Options.
; Syntax ........: _LOWriter_ImageOptions(ByRef $oImage[, $bProtectContent = Null[, $bProtectPos = Null[, $bProtectSize = Null[, $bPrint = Null]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $bProtectContent     - [optional] a boolean value. Default is Null. If True, Prevents changes to the contents of the Image.
;                  $bProtectPos         - [optional] a boolean value. Default is Null. If True, Locks the position of the Image in the current document.
;                  $bProtectSize        - [optional] a boolean value. Default is Null. If True, Locks the size of the Image.
;                  $bPrint              - [optional] a boolean value. Default is Null. If True, Includes the image when you print the document.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bProtectContent not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bProtectPos not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bProtectSize not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bPrint not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bProtectContent
;                  |                               2 = Error setting $bProtectPos
;                  |                               4 = Error setting $bProtectSize
;                  |                               8 = Error setting $bPrint
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LOWriter_ImageOptionsName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageOptions(ByRef $oImage, $bProtectContent = Null, $bProtectPos = Null, $bProtectSize = Null, $bPrint = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abOptions[4]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bProtectContent, $bProtectPos, $bProtectSize, $bPrint) Then
		__LO_ArrayFill($abOptions, $oImage.ContentProtected(), $oImage.PositionProtected(), $oImage.SizeProtected(), $oImage.Print())

		Return SetError($__LO_STATUS_SUCCESS, 1, $abOptions)
	EndIf

	If ($bProtectContent <> Null) Then
		If Not IsBool($bProtectContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oImage.ContentProtected = $bProtectContent
		$iError = ($oImage.ContentProtected() = $bProtectContent) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bProtectPos <> Null) Then
		If Not IsBool($bProtectPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.PositionProtected = $bProtectPos
		$iError = ($oImage.PositionProtected() = $bProtectPos) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bProtectSize <> Null) Then
		If Not IsBool($bProtectSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.SizeProtected = $bProtectSize
		$iError = ($oImage.SizeProtected() = $bProtectSize) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bPrint <> Null) Then
		If Not IsBool($bPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImage.Print = $bPrint
		$iError = ($oImage.Print() = $bPrint) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageOptions

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageOptionsName
; Description ...: Set or Retrieve Image Name settings.
; Syntax ........: _LOWriter_ImageOptionsName(ByRef $oDoc, ByRef $oImage[, $sName = Null[, $sAltText = Null[, $sDesc = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $sName               - [optional] a string value. Default is Null. The new name for the Image.
;                  $sAltText            - [optional] a string value. Default is Null. Enter alternative text to display when the image isn't available.
;                  $sDesc               - [optional] a string value. Default is Null. Description of the Image.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = Document already contains Image with same name as called in $sName.
;                  @Error 1 @Extended 5 Return 0 = $sAltText not a string.
;                  @Error 1 @Extended 6 Return 0 = $sDesc not a string.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sAltText
;                  |                               4 = Error setting $sDesc
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Previous and Next link are omitted as they seem to have no use for images.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LOWriter_ImageOptions
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageOptionsName(ByRef $oDoc, ByRef $oImage, $sName = Null, $sAltText = Null, $sDesc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asName[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sName, $sAltText, $sDesc) Then
		__LO_ArrayFill($asName, $oImage.Name(), $oImage.Title(), $oImage.Description())

		Return SetError($__LO_STATUS_SUCCESS, 1, $asName)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If _LOWriter_ImageExists($oDoc, $sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.Name = $sName
		$iError = ($oImage.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sAltText <> Null) Then
		If Not IsString($sAltText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImage.Title = $sAltText
		$iError = ($oImage.Title() = $sAltText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sDesc <> Null) Then
		If Not IsString($sDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImage.Description = $sDesc
		$iError = ($oImage.Description() = $sDesc) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageOptionsName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageReplace
; Description ...: Replace an image with another image.
; Syntax ........: _LOWriter_ImageReplace(ByRef $oImage, $sNewImage)
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $sNewImage           - a string value. The file path to the new image.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sNewImage not a string.
;                  @Error 1 @Extended 3 Return 0 = File called in $sNewImage doesn't exist.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to convert $sNewImage Path to Libre Office URL.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Image was successfully replaced.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Unfortunately I am unable to find a way to convert or insert an image as a linked image instead of an embedded image. All linked images will remain as such, all embedded images will stay as embedded.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageReplace(ByRef $oImage, $sNewImage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sNewImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not FileExists($sNewImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$sNewImage = _LO_PathConvert($sNewImage, $LO_PATHCONV_OFFICE_RETURN)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oImage.GraphicURL = $sNewImage

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_ImageReplace

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImagesGetNames
; Description ...: Retrieve an array of image names contained in a document.
; Syntax ........: _LOWriter_ImagesGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve GraphicObjects object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Image Names. @Extended set to number of Names returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ImageGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImagesGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asImages[0]
	Local $oImages

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oImages = $oDoc.GraphicObjects()
	If Not IsObj($oImages) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asImages[$oImages.getCount()]
	For $i = 0 To $oImages.getCount() - 1
		$asImages[$i] = ($oImages.getByIndex($i).Name)
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))     ; Sleep every x cycles.
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($asImages), $asImages)
EndFunc   ;==>_LOWriter_ImagesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageShadow
; Description ...: Set or Retrieve the shadow settings for an Image.
; Syntax ........: _LOWriter_ImageShadow(ByRef $oImage[, $iWidth = Null[, $iColor = Null[, $iLocation = Null]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iWidth              - [optional] an integer value. Default is Null. The Width of the Image Shadow set in Hundredths of a Millimeter (HMM).
;                  $iColor              - [optional] an integer value (-1-16777215). Default is Null. The Color of the Image shadow, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The Location of the Image Shadow. See constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageShadow(ByRef $oImage, $iWidth = Null, $iColor = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tShdwFrmt
	Local $iError = 0
	Local $avShadow[3]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tShdwFrmt = $oImage.ShadowFormat()
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
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tShdwFrmt.Color = $iColor
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LOW_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oImage.ShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oImage.ShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($tShdwFrmt.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iColor)) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iLocation)) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 4)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageSize
; Description ...: Set or retrieve Image size settings.
; Syntax ........: _LOWriter_ImageSize(ByRef $oImage[, $iScaleWidth = Null[, $iScaleHeight = Null[, $iWidth = Null[, $iHeight = Null[, $bOriginalSize = Null]]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iScaleWidth         - [optional] an integer value (Min. 1%). Default is Null. The Scale Width percentage of the image.
;                  $iScaleHeight        - [optional] an integer value (Min. 1%). Default is Null. The Scale Height percentage of the image.
;                  $iWidth              - [optional] an integer value. Default is Null. The Width of the image, set in Hundredths of a Millimeter (HMM).
;                  $iHeight             - [optional] an integer value. Default is Null. The Height of the image, set in Hundredths of a Millimeter (HMM).
;                  $bOriginalSize       - [optional] a boolean value. Default is Null. Only accepts True. If True, the image is returned to its original size, or the maximum size allowed for the current page size.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iScaleWidth not an Integer, or less than 1%.
;                  @Error 1 @Extended 3 Return 0 = $iScaleHeight not an Integer, or less than 1%.
;                  @Error 1 @Extended 4 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $bOriginalSize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the image's Actual Size structure.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the image's Size structure.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the image's Size structure again after setting scale sizing.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iScaleWidth
;                  |                               2 = Error setting $iScaleHeight
;                  |                               4 = Error setting $iWidth
;                  |                               8 = Error setting $iHeight
;                  |                               16 = Error setting Image to Original Size, possibly the page size is smaller than the image size.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The return for $bOriginalSize is a Boolean whether the image is currently set to its original size (True) or not.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageSize(ByRef $oImage, $iScaleWidth = Null, $iScaleHeight = Null, $iWidth = Null, $iHeight = Null, $bOriginalSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avImage[5]
	Local $tSize, $tOrigSize

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tOrigSize = $oImage.ActualSize()
	If Not IsObj($tOrigSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tSize = $oImage.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iScaleWidth, $iScaleHeight, $iWidth, $iHeight, $bOriginalSize) Then
		__LO_ArrayFill($avImage, Round(($tSize.Width() / $tOrigSize.Width()) * 100), _
				Round(($tSize.Height() / $tOrigSize.Height()) * 100), $tSize.Width(), $tSize.Height(), _
				((($oImage.Size.Width() = $tOrigSize.Width()) And $oImage.Size.Height() = $tOrigSize.Height()) ? (True) : (False))) ; If image is set to its original size, return true.

		Return SetError($__LO_STATUS_SUCCESS, 1, $avImage)
	EndIf

	If ($iScaleWidth <> Null) Or ($iScaleHeight <> Null) Then
		If ($iScaleWidth <> Null) Then
			If Not __LO_IntIsBetween($iScaleWidth, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Min is 1%, no max

			$tSize.Width = Int(($tOrigSize.Width() * ($iScaleWidth / 100))) ; Times original Width by scale percentage
		EndIf

		If ($iScaleHeight <> Null) Then
			If Not __LO_IntIsBetween($iScaleHeight, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Min 1%, no max

			$tSize.Height = Int(($tOrigSize.Height() * ($iScaleHeight / 100))) ; Times original Height by scale percentage
		EndIf

		$oImage.Size = $tSize

		$tSize = $oImage.Size() ; Retrieve the size Struct again.
		If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		; Error checking
		$iError = (__LO_VarsAreNull($iScaleWidth)) ? ($iError) : ((Round(($tSize.Width() / $tOrigSize.Width()) * 100) = $iScaleWidth) ? ($iError) : (BitOR($iError, 1)))
		$iError = (__LO_VarsAreNull($iScaleHeight)) ? ($iError) : ((Round(($tSize.Height() / $tOrigSize.Height()) * 100) = $iScaleHeight) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($iWidth <> Null) Or ($iHeight <> Null) Then
		If ($iWidth <> Null) Then
			If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$tSize.Width = $iWidth
		EndIf

		If ($iHeight <> Null) Then
			If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$tSize.Height = $iHeight
		EndIf

		$oImage.Size = $tSize

		; Error checking
		$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($oImage.Size.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 4)))
		$iError = (__LO_VarsAreNull($iHeight)) ? ($iError) : ((__LO_IntIsBetween($oImage.Size.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 8)))
	EndIf

	If ($bOriginalSize = True) Then
		$tSize.Width = $tOrigSize.Width()
		$tSize.Height = $tOrigSize.Height()

		$oImage.Size = $tSize

		$iError = (($oImage.Size.Width() = $tOrigSize.Width()) And $oImage.Size.Height() = $tOrigSize.Height()) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageTransparency
; Description ...: Set or retrieve Image transparency settings.
; Syntax ........: _LOWriter_ImageTransparency(ByRef $oImage[, $iTransparency = Null])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The percentage of transparency. 0% = visible, 100% = transparent.
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTransparency
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning the current Transparency setting as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LOWriter_ImageAreaTransparency
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageTransparency(ByRef $oImage, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oImage.Transparency())

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oImage.Transparency = $iTransparency

	Return ($oImage.Transparency() = $iTransparency) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOWriter_ImageTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageTypePosition
; Description ...: Set or Retrieve Image Position Settings.
; Syntax ........: _LOWriter_ImageTypePosition(ByRef $oImage[, $iHorAlign = Null[, $iHorPos = Null[, $iHorRelation = Null[, $bMirror = Null[, $iVertAlign = Null[, $iVertPos = Null[, $iVertRelation = Null[, $bKeepInside = Null[, $iAnchorPos = Null]]]]]]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iHorAlign           - [optional] an integer value (0-3). Default is Null. The horizontal orientation of the Image. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3. Can't be set if Anchor position is set to "As Character".
;                  $iHorPos             - [optional] an integer value. Default is Null. The horizontal position of the Image. set in Hundredths of a Millimeter (HMM). Only valid if $iHorAlign is set to $LOW_ORIENT_HORI_NONE().
;                  $iHorRelation        - [optional] an integer value (0-8). Default is Null. The reference point for the selected horizontal alignment option. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3, and Remarks for acceptable values.
;                  $bMirror             - [optional] a boolean value. Default is Null. If True, Reverses the current horizontal alignment settings on even pages.
;                  $iVertAlign          - [optional] an integer value (0-9). Default is Null. The vertical orientation of the Image. See Constants, $LOW_ORIENT_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertPos            - [optional] an integer value. Default is Null. The vertical position of the Image. set in Hundredths of a Millimeter (HMM). Only valid if $iVertAlign is set to $LOW_ORIENT_VERT_NONE().
;                  $iVertRelation       - [optional] an integer value (-1-9). Default is Null. The reference point for the selected vertical alignment option. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3, and Remarks for acceptable values.
;                  $bKeepInside         - [optional] a boolean value. Default is Null. If True, Keeps the Image within the layout boundaries of the text that the Image is anchored to.
;                  $iAnchorPos          - [optional] an integer value (0-2,4). Default is Null. Specify the anchoring options for the Image. See Constants, $LOW_ANCHOR_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
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
;                  $iHorRelation has varying acceptable values, depending on the current Anchor position and also the current $iHorAlign setting. The Following is a list of acceptable values per anchor position.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageTypePosition(ByRef $oImage, $iHorAlign = Null, $iHorPos = Null, $iHorRelation = Null, $bMirror = Null, $iVertAlign = Null, $iVertPos = Null, $iVertRelation = Null, $bKeepInside = Null, $iAnchorPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurrentAnchor
	Local $avPosition[9]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iHorAlign, $iHorPos, $iHorRelation, $bMirror, $iVertAlign, $iVertPos, $iVertRelation, $bKeepInside, $iAnchorPos) Then
		__LO_ArrayFill($avPosition, $oImage.HoriOrient(), $oImage.HoriOrientPosition(), $oImage.HoriOrientRelation(), _
				$oImage.PageToggle(), $oImage.VertOrient(), $oImage.VertOrientPosition(), $oImage.VertOrientRelation(), _
				$oImage.IsFollowingTextFlow(), $oImage.AnchorType())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosition)
	EndIf
	; Accepts HoriOrient Left, Right, Center, and "None" = "From Left"
	If ($iHorAlign <> Null) Then ; Cant be set if Anchor is set to "As Char"
		If Not __LO_IntIsBetween($iHorAlign, $LOW_ORIENT_HORI_NONE, $LOW_ORIENT_HORI_LEFT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oImage.HoriOrient = $iHorAlign
		$iError = ($oImage.HoriOrient() = $iHorAlign) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iHorPos <> Null) Then
		If Not IsInt($iHorPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.HoriOrientPosition = $iHorPos
		$iError = (__LO_IntIsBetween($oImage.HoriOrientPosition(), $iHorPos - 1, $iHorPos + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iHorRelation <> Null) Then
		If Not __LO_IntIsBetween($iHorRelation, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PAGE_PRINT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.HoriOrientRelation = $iHorRelation
		$iError = ($oImage.HoriOrientRelation() = $iHorRelation) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bMirror <> Null) Then
		If Not IsBool($bMirror) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImage.PageToggle = $bMirror
		$iError = ($oImage.PageToggle() = $bMirror) ? ($iError) : (BitOR($iError, 8))
	EndIf

	; Accepts Orient Top, Bottom, Center, and "None" = "From Top"/From Bottom, plus Row and Char.
	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ORIENT_VERT_NONE, $LOW_ORIENT_VERT_LINE_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImage.VertOrient = $iVertAlign
		$iError = ($oImage.VertOrient() = $iVertAlign) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iVertPos <> Null) Then
		If Not IsInt($iVertPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oImage.VertOrientPosition = $iVertPos
		$iError = (__LO_IntIsBetween($oImage.VertOrientPosition(), $iVertPos - 1, $iVertPos + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iVertRelation <> Null) Then
		If Not __LO_IntIsBetween($iVertRelation, $LOW_RELATIVE_ROW, $LOW_RELATIVE_TEXT_LINE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$iCurrentAnchor = (($iAnchorPos <> Null) ? ($iAnchorPos) : ($oImage.AnchorType()))

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
				Switch $oImage.VertOrient()
					Case $LOW_ORIENT_VERT_NONE ; None = "From Bottom or From Top in L.O. UI
						$iError = BitOR($iError, 64) ; -- Row not accepted with this VertOrient Setting.

					Case $LOW_ORIENT_VERT_TOP, $LOW_ORIENT_VERT_CHAR_TOP, $LOW_ORIENT_VERT_LINE_TOP
						$oImage.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oImage.VertOrient = $LOW_ORIENT_VERT_LINE_TOP
						$iError = (($oImage.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oImage.VertOrient() = $LOW_ORIENT_VERT_LINE_TOP)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_CENTER, $LOW_ORIENT_VERT_CHAR_CENTER, $LOW_ORIENT_VERT_LINE_CENTER
						$oImage.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oImage.VertOrient = $LOW_ORIENT_VERT_LINE_CENTER
						$iError = (($oImage.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oImage.VertOrient() = $LOW_ORIENT_VERT_LINE_CENTER)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_BOTTOM, $LOW_ORIENT_VERT_CHAR_BOTTOM, $LOW_ORIENT_VERT_LINE_BOTTOM
						$oImage.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oImage.VertOrient = $LOW_ORIENT_VERT_LINE_BOTTOM
						$iError = (($oImage.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oImage.VertOrient() = $LOW_ORIENT_VERT_LINE_BOTTOM)) ? ($iError) : (BitOR($iError, 64))
				EndSwitch

			ElseIf ($iVertRelation = $LOW_RELATIVE_PARAGRAPH) Then ; Paragraph = Baseline setting in L.O. UI
				$oImage.VertOrientRelation = $iVertRelation ; Paragraph = Baseline in this case
				$iError = (($oImage.VertOrientRelation() = $iVertRelation)) ? ($iError) : (BitOR($iError, 64))

			ElseIf ($iVertRelation = $LOW_RELATIVE_CHARACTER) Then
				Switch $oImage.VertOrient()
					Case $LOW_ORIENT_VERT_NONE ; None = "From Bottom or From Top in L.O. UI
						$iError = BitOR($iError, 64) ; -- Character not accepted with this VertOrient Setting.

					Case $LOW_ORIENT_VERT_TOP, $LOW_ORIENT_VERT_CHAR_TOP, $LOW_ORIENT_VERT_LINE_TOP
						$oImage.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oImage.VertOrient = $LOW_ORIENT_VERT_CHAR_TOP
						$iError = (($oImage.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oImage.VertOrient() = $LOW_ORIENT_VERT_CHAR_TOP)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_CENTER, $LOW_ORIENT_VERT_CHAR_CENTER, $LOW_ORIENT_VERT_LINE_CENTER
						$oImage.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oImage.VertOrient = $LOW_ORIENT_VERT_CHAR_CENTER
						$iError = (($oImage.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oImage.VertOrient() = $LOW_ORIENT_VERT_CHAR_CENTER)) ? ($iError) : (BitOR($iError, 64))

					Case $LOW_ORIENT_VERT_BOTTOM, $LOW_ORIENT_VERT_CHAR_BOTTOM, $LOW_ORIENT_VERT_LINE_BOTTOM
						$oImage.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oImage.VertOrient = $LOW_ORIENT_VERT_CHAR_BOTTOM
						$iError = (($oImage.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oImage.VertOrient() = $LOW_ORIENT_VERT_CHAR_BOTTOM)) ? ($iError) : (BitOR($iError, 64))
				EndSwitch
			EndIf

		Else
			$oImage.VertOrientRelation = $iVertRelation
			$iError = ($oImage.VertOrientRelation() = $iVertRelation) ? ($iError) : (BitOR($iError, 64))
		EndIf
	EndIf

	If ($bKeepInside <> Null) Then
		If Not IsBool($bKeepInside) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oImage.IsFollowingTextFlow = $bKeepInside
		$iError = ($oImage.IsFollowingTextFlow() = $bKeepInside) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iAnchorPos <> Null) Then
		If Not __LO_IntIsBetween($iAnchorPos, $LOW_ANCHOR_AT_PARAGRAPH, $LOW_ANCHOR_AT_CHARACTER, $LOW_ANCHOR_AT_FRAME) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oImage.AnchorType = $iAnchorPos
		$iError = ($oImage.AnchorType() = $iAnchorPos) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageTypePosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageTypeSize
; Description ...: Set or Retrieve Image Size related settings.
; Syntax ........: _LOWriter_ImageTypeSize(ByRef $oDoc, ByRef $oImage[, $iWidth = Null[, $iRelativeWidth = Null[, $iWidthRelativeTo = Null[, $iHeight = Null[, $iRelativeHeight = Null[, $iHeightRelativeTo = Null[, $bKeepRatio = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Image, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iRelativeWidth      - [optional] an integer value (0-254). Default is Null. Calculates the width of the Image as a percentage of the width of the page text area. 0 = off.
;                  $iWidthRelativeTo    - [optional] an integer value (0,7). Default is Null. Decides what 100% width means: either text area (excluding margins) or the entire page (including margins). See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3. Libre Office 4.3 and Up.
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the Image, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iRelativeHeight     - [optional] an integer value (0-254). Default is Null. Calculates the Height of the Image as a percentage of the Height of the page text area. 0 = off.
;                  $iHeightRelativeTo   - [optional] an integer value (0,7). Default is Null. Decides what 100% Height means: either text area (excluding margins) or the entire page (including margins). See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3. Libre Office 4.3 and Up.
;                  $bKeepRatio          - [optional] a boolean value. Default is Null. Maintains the height and width ratio when you change the width or the height setting.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iWidth Not an Integer, or less than 51.
;                  @Error 1 @Extended 4 Return 0 = $iRelativeWidth not an Integer, less than 0 or greater than 254.
;                  @Error 1 @Extended 5 Return 0 = $iWidthRelativeTo not an Integer, not equal to 0, and not equal to 7. See Constants.
;                  @Error 1 @Extended 6 Return 0 = $iHeight Not an Integer, or less than 51.
;                  @Error 1 @Extended 7 Return 0 = $iRelativeHeight not an Integer, less than 0 or greater than 254.
;                  @Error 1 @Extended 8 Return 0 = $iHeightRelativeTo not an Integer, not equal to 0 and not equal to 7. See Constants.
;                  @Error 1 @Extended 9 Return 0 = $bKeepRatio not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iRelativeWidth
;                  |                               4 = Error setting $iWidthRelativeTo
;                  |                               8 = Error setting $iHeight
;                  |                               16 = Error setting $iRelativeHeight
;                  |                               32 = Error setting $iHeightRelativeTo
;                  |                               64 = Error setting $bKeepRatio
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.3.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 or 7 Element Array depending on current Libre Office Version, If the current Libre Office version is greater than or equal to 4.3, then a 7 element Array is returned, else 5 element array with both $iWidthRelativeTo and $iHeightRelativeTo skipped. Array Element values will be in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function can successfully set "Keep Ratio" however when the user changes this setting in the UI, for some reason the applicable setting values are not updated, so this function may return incorrect values for "Keep Ratio".
;                  When Keep Ratio is set to True, setting Width/Height values via this function will not be kept in ratio.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageTypeSize(ByRef $oDoc, ByRef $oImage, $iWidth = Null, $iRelativeWidth = Null, $iWidthRelativeTo = Null, $iHeight = Null, $iRelativeHeight = Null, $iHeightRelativeTo = Null, $bKeepRatio = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSize[7]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iWidth, $iRelativeWidth, $iWidthRelativeTo, $iHeight, $iRelativeHeight, $iHeightRelativeTo, $bKeepRatio) Then
		If __LO_VersionCheck(4.3) Then
			__LO_ArrayFill($avSize, $oImage.Width(), $oImage.RelativeWidth(), $oImage.RelativeWidthRelation(), _
					$oImage.Height(), $oImage.RelativeHeight(), $oImage.RelativeHeightRelation(), _
					(($oImage.IsSyncHeightToWidth() And $oImage.IsSyncWidthToHeight()) ? (True) : (False)))

		Else
			__LO_ArrayFill($avSize, $oImage.Width(), $oImage.RelativeWidth(), $oImage.Height(), $oImage.RelativeHeight(), _
					(($oImage.IsSyncHeightToWidth() And $oImage.IsSyncWidthToHeight()) ? (True) : (False)))
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSize)
	EndIf

	If ($iWidth <> Null) Then ; Min 51
		If Not __LO_IntIsBetween($iWidth, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.Width = $iWidth
		$iError = (__LO_IntIsBetween($oImage.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRelativeWidth <> Null) Then
		If Not __LO_IntIsBetween($iRelativeWidth, 0, 254) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.RelativeWidth = $iRelativeWidth
		$iError = ($oImage.RelativeWidth() = $iRelativeWidth) ? ($iError) : (BitOR($iError, 2))

		If ($iRelativeWidth <> 0) Then __LOWriter_ObjRelativeSize($oDoc, $oImage, True) ; If Relative Width isn't being turned off, then set Width Value.
	EndIf

	If ($iWidthRelativeTo <> Null) Then
		If Not __LO_IntIsBetween($iWidthRelativeTo, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PARAGRAPH, "", $LOW_RELATIVE_PAGE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not __LO_VersionCheck(4.3) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oImage.RelativeWidthRelation = $iWidthRelativeTo
		$iError = ($oImage.RelativeWidthRelation() = $iWidthRelativeTo) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iHeight <> Null) Then
		If Not __LO_IntIsBetween($iHeight, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImage.Height = $iHeight
		$iError = ($oImage.Height() = $iHeight) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRelativeHeight <> Null) Then
		If Not __LO_IntIsBetween($iRelativeHeight, 0, 254) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oImage.RelativeHeight = $iRelativeHeight
		$iError = ($oImage.RelativeHeight() = $iRelativeHeight) ? ($iError) : (BitOR($iError, 16))

		If ($iRelativeHeight <> 0) Then __LOWriter_ObjRelativeSize($oDoc, $oImage, False, True) ; If Relative Height isn't being turned off, then set Height Value.
	EndIf

	If ($iHeightRelativeTo <> Null) Then
		If Not __LO_IntIsBetween($iHeightRelativeTo, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PARAGRAPH, "", $LOW_RELATIVE_PAGE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		If Not __LO_VersionCheck(4.3) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oImage.RelativeHeightRelation = $iHeightRelativeTo
		$iError = ($oImage.RelativeHeightRelation() = $iHeightRelativeTo) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bKeepRatio <> Null) Then
		If Not IsBool($bKeepRatio) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oImage.IsSyncHeightToWidth = $bKeepRatio
		$oImage.IsSyncWidthToHeight = $bKeepRatio
		$iError = (($oImage.IsSyncHeightToWidth() = $bKeepRatio) And ($oImage.IsSyncWidthToHeight() = $bKeepRatio)) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageTypeSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageWrap
; Description ...: Set or Retrieve Image Wrap and Spacing settings.
; Syntax ........: _LOWriter_ImageWrap(ByRef $oImage[, $iWrapType = Null[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $iWrapType           - [optional] an integer value (0-5). Default is Null. The way to wrap text around the Image. See Constants, $LOW_WRAP_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The amount of space between the left edge of the Image and the text. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] an integer value. Default is Null. The amount of space between the Right edge of the Image and the text. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] an integer value. Default is Null. The amount of space between the Top edge of the Image and the text. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] an integer value. Default is Null. The amount of space between the Bottom edge of the Image and the text. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWrapType not an Integer, less than 0 or greater than 5. See Constants.
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
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LO_UnitConvert, _LOWriter_ImageWrapOptions
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageWrap(ByRef $oImage, $iWrapType = Null, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPropInfo
	Local $iError = 0
	Local $avWrap[5]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPropInfo = $oImage.getPropertySetInfo()
	If Not IsObj($oPropInfo) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWrapType, $iLeft, $iRight, $iTop, $iBottom) Then
		If $oPropInfo.hasPropertyByName("Surround") Then ; Surround is marked as deprecated, but there is no indication of what version of L.O. this occurred. So Test for its existence.
			__LO_ArrayFill($avWrap, $oImage.Surround(), $oImage.LeftMargin(), $oImage.RightMargin(), $oImage.TopMargin(), _
					$oImage.BottomMargin())

		Else
			__LO_ArrayFill($avWrap, $oImage.TextWrap(), $oImage.LeftMargin(), $oImage.RightMargin(), $oImage.TopMargin(), _
					$oImage.BottomMargin())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avWrap)
	EndIf

	If ($iWrapType <> Null) Then
		If Not __LO_IntIsBetween($iWrapType, $LOW_WRAP_MODE_NONE, $LOW_WRAP_MODE_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If $oPropInfo.hasPropertyByName("Surround") Then $oImage.Surround = $iWrapType
		If $oPropInfo.hasPropertyByName("TextWrap") Then $oImage.TextWrap = $iWrapType
		If $oPropInfo.hasPropertyByName("Surround") Then
			$iError = ($oImage.Surround() = $iWrapType) ? ($iError) : (BitOR($iError, 1))

		Else
			$iError = ($oImage.TextWrap() = $iWrapType) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($iLeft <> Null) Then
		If Not IsInt($iLeft) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.LeftMargin = $iLeft
		$iError = (__LO_IntIsBetween($oImage.LeftMargin(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iRight <> Null) Then
		If Not IsInt($iRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.RightMargin = $iRight
		$iError = (__LO_IntIsBetween($oImage.RightMargin(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iTop <> Null) Then
		If Not IsInt($iTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImage.TopMargin = $iTop
		$iError = (__LO_IntIsBetween($oImage.TopMargin(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iBottom <> Null) Then
		If Not IsInt($iBottom) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImage.BottomMargin = $iBottom
		$iError = (__LO_IntIsBetween($oImage.BottomMargin(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageWrap

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ImageWrapOptions
; Description ...: Set or Retrieve Image Wrap Options.
; Syntax ........: _LOWriter_ImageWrapOptions(ByRef $oImage[, $bFirstPar = Null[, $bContour = Null[, $bOutsideOnly = Null[, $bInBackground = Null[, $bAllowOverlap = Null]]]]])
; Parameters ....: $oImage              - [in/out] an object. A Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $bFirstPar           - [optional] a boolean value. Default is Null. If True, starts a new paragraph below the Image.
;                  $bContour            - [optional] a boolean value. Default is Null. If True, text is wrapped around the shape of the Image. This option is not available for the Through wrap type.
;                  $bOutsideOnly        - [optional] a boolean value. Default is Null. If True, text is wrapped only around the contour of the Image, but not in open areas within the Image shape. $bContour must be True before this can be set.
;                  $bInBackground       - [optional] a boolean value. Default is Null. If True, moves the selected Image to the background. This option is only available with the "Through" wrap type.
;                  $bAllowOverlap       - [optional] a boolean value. Default is Null. If True, the Image is allowed to overlap another Image. This option has no effect on wrap through Images, which can always overlap.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImage not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFirstPar not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bContour not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bOutsideOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bInBackground not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bAllowOverlap not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bFirstPar
;                  |                               2 = Error setting $bContour
;                  |                               4 = Error setting $bOutsideOnly
;                  |                               8 = Error setting $bInBackground
;                  |                               16 = Error setting $bAllowOverlap
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  This function may indicate the settings were set successfully when they haven't been if the appropriate wrap type, anchor type etc. hasn't been set before hand.
; Related .......: _LOWriter_ImageInsert, _LOWriter_ImageGetObjByName, _LOWriter_ImageWrap
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ImageWrapOptions(ByRef $oImage, $bFirstPar = Null, $bContour = Null, $bOutsideOnly = Null, $bInBackground = Null, $bAllowOverlap = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abWrapOptions[5]

	If Not IsObj($oImage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bFirstPar, $bContour, $bOutsideOnly, $bInBackground, $bAllowOverlap) Then
		__LO_ArrayFill($abWrapOptions, $oImage.SurroundAnchorOnly(), $oImage.SurroundContour(), $oImage.ContourOutside(), _
				(($oImage.Opaque()) ? (False) : (True)), $oImage.AllowOverlap()) ; Opaque/Background is False when InBackground is checked, so switch Boolean values around.

		Return SetError($__LO_STATUS_SUCCESS, 1, $abWrapOptions)
	EndIf

	If ($bFirstPar <> Null) Then
		If Not IsBool($bFirstPar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oImage.SurroundAnchorOnly = $bFirstPar
		$iError = ($oImage.SurroundAnchorOnly() = $bFirstPar) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bContour <> Null) Then
		If Not IsBool($bContour) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImage.SurroundContour = $bContour
		$iError = ($oImage.SurroundContour() = $bContour) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bOutsideOnly <> Null) Then
		If Not IsBool($bOutsideOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImage.ContourOutside = $bOutsideOnly
		$iError = ($oImage.ContourOutside() = $bOutsideOnly) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bInBackground <> Null) Then
		If Not IsBool($bInBackground) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImage.Opaque = (($bInBackground) ? False : True)
		$iError = ($oImage.Opaque() = (($bInBackground) ? False : True)) ? ($iError) : (BitOR($iError, 8)) ; Opaque/Background is False when InBackground is checked, so switch Boolean values around.
	EndIf

	If ($bAllowOverlap <> Null) Then
		If Not IsBool($bAllowOverlap) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImage.AllowOverlap = $bAllowOverlap
		$iError = ($oImage.AllowOverlap() = $bAllowOverlap) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ImageWrapOptions
