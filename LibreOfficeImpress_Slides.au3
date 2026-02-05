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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Deleting, etc. L.O. Impress Slides.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_SlideAdd
; _LOImpress_SlideBackColor
; _LOImpress_SlideBackFillStyle
; _LOImpress_SlideBackGradient
; _LOImpress_SlideBackTransparency
; _LOImpress_SlideBackTransparencyGradient
; _LOImpress_SlideCopy
; _LOImpress_SlideCurrent
; _LOImpress_SlideDeleteByIndex
; _LOImpress_SlideDeleteByObj
; _LOImpress_SlideGetByIndex
; _LOImpress_SlideMove
; _LOImpress_SlidesGetCount
; _LOImpress_SlideShapesGetList
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideAdd
; Description ...: Add a slide to a presentation.
; Syntax ........: _LOImpress_SlideAdd(ByRef $oDoc[, $iPos = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iPos                - [optional] an integer value. Default is Null. The position to insert the new slide in the collection of slides. 0 Based. See remarks.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iPos not an Integer, less than 0 or greater than number of slides.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to create a slide.
;                  @Error 3 @Extended 2 Return 0 = Failed to backup currently active slide.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning new slide's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $iPos is called with Null, the new slide is inserted at the end.
;                  Call $iPos with the last slide index to insert the slide at the end.
;                  Due to limitations in the API, I have made a small workaround for inserting a slide at the beginning. A dispatch is executed to move the slide to the beginning. The current slide will temporarily be set to the new slide in order to move it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideAdd(ByRef $oDoc, $iPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSlide, $oServiceManager, $oDispatcher, $oCurrSlide
	Local $bMoveToFirst = False
	Local $aArray[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($iPos = Null) Then $iPos = $oDoc.DrawPages.getCount()
	If Not __LO_IntIsBetween($iPos, 0, $oDoc.DrawPages.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iPos -= 1 ; -1 because when 0 is called in insertNewByIndex, it inserts it in position 1, etc. Also there is no way to insert a new slide at position 0, so I made a workaround.

	If ($iPos = -1) Then
		$iPos = 0
		$bMoveToFirst = True
	EndIf

	$oSlide = $oDoc.DrawPages.insertNewByIndex($iPos)
	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $bMoveToFirst Then
		$oCurrSlide = $oDoc.getCurrentController.CurrentPage() ; Backup current slide
		If Not IsObj($oCurrSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oDoc.getCurrentController.setCurrentPage($oSlide)

		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
		If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:MovePageFirst", "", 0, $aArray)

		$oDoc.getCurrentController.setCurrentPage($oCurrSlide) ; Restore current slide.
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSlide)
EndFunc   ;==>_LOImpress_SlideAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideBackColor
; Description ...: Set or Retrieve the Slide's background color.
; Syntax ........: _LOImpress_SlideBackColor(ByRef $oSlide[, $iColor = Null])
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Slide background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve parent Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting as an Integer value. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current setting.
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the slide, the Constant $LO_COLOR_OFF is returned.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideBackColor(ByRef $oSlide, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBackground, $oDoc

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oSlide.Background()

	If __LO_VarsAreNull($iColor) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, $LO_COLOR_OFF) ; If no background is set, this will be void, instead of an Object.

		Return SetError($__LO_STATUS_SUCCESS, 1, $oBackground.FillColor())
	EndIf

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oDoc = $oSlide.MasterPage.Forms.Parent()
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	$oBackground.FillStyle = $LOI_AREA_FILL_STYLE_SOLID
	$oBackground.FillColor = $iColor

	$oSlide.Background = $oBackground

	If ($oSlide.Background.FillColor() <> $iColor) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideBackFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOImpress_SlideBackFillStyle(ByRef $oSlide[, $bFillOff = False])
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $bFillOff            - [optional] a boolean value. Default is False. If True, the Fill style will be set to Off. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFillOff not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Fill Style.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning current background fill style. Return will be one of the constants $LOI_AREA_FILL_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 0 @Extended 1 Return 0 = Success. Fill style was successfully turned off.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
;                  When the Fill style is disabled for a Slide, the Fill properties are completely removed. This is how Impress works normally.
;                  $bFillOff will do nothing if it is called with False, and is not, of course, returned when retrieving the FillStyle value.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideBackFillStyle(ByRef $oSlide, $bFillOff = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle
	Local $oBackground

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bFillOff) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bFillOff Then
		If IsObj($oSlide.Background()) Then
			$oBackground = $oSlide.Background
			If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, 0) ; If no Background Object, no Fillstyle is active.

			$oBackground.FillStyle = $LOI_AREA_FILL_STYLE_OFF
			$oSlide.Background = $oBackground
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, 0)
	EndIf

	$oBackground = $oSlide.Background
	If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_AREA_FILL_STYLE_OFF) ; If no Background Object, no Fillstyle is active.

	$iFillStyle = $oBackground.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOImpress_SlideBackFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideBackGradient
; Description ...: Modify or retrieve the settings for Slide Background color Gradient.
; Syntax ........: _LOImpress_SlideBackGradient(ByRef $oSlide[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
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
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
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
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving "FillGradient" Struct.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Parent Document.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving Background Object.
;                  @Error 3 @Extended 4 Return 0 = Error retrieving Color Stop Array for "From" color
;                  @Error 3 @Extended 5 Return 0 = Error retrieving Color Stop Array for "To" color
;                  @Error 3 @Extended 6 Return 0 = Error creating Gradient Name.
;                  @Error 3 @Extended 7 Return 0 = Error setting Gradient Name.
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
;                  @Error 0 @Extended 2 Return -1 = Success. All optional parameters were called with Null, no background is currently active for the slide. Returning -1.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideBackGradient(ByRef $oSlide, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBackground, $oDoc
	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $nRed, $nGreen, $nBlue
	Local $atColorStop
	Local $avGradient[11]
	Local $sGradName

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oSlide.Background()

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 2, -1) ; No background active.

		$tStyleGradient = $oBackground.FillGradient()
		If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LO_ArrayFill($avGradient, $oBackground.FillGradientName(), $tStyleGradient.Style(), _
				$oBackground.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), ($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	$tStyleGradient = $oBackground.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oBackground.FillStyle() <> $LOI_AREA_FILL_STYLE_GRADIENT) Then $oBackground.FillStyle = $LOI_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		__LOImpress_GradientPresets($oDoc, $oBackground, $tStyleGradient, $sGradientName)

		$oSlide.Background = $oBackground

		$oBackground = $oSlide.Background()
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$tStyleGradient = $oBackground.FillGradient()
		If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$iError = ($oBackground.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oBackground.FillStyle = $LOI_AREA_FILL_STYLE_OFF
			$oBackground.FillGradientName = ""
			$oSlide.Background = $oBackground

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOI_GRAD_TYPE_LINEAR, $LOI_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oBackground.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oBackground.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
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

		$tStyleGradient.Angle = ($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iFromColor <> Null) Then
		If Not __LO_IntIsBetween($iFromColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tStyleGradient.StartColor = $iFromColor

		If __LO_VersionCheck(7.6) Then
			$nRed = (BitAND(BitShift($iFromColor, 16), 0xff) / 255)
			$nGreen = (BitAND(BitShift($iFromColor, 8), 0xff) / 255)
			$nBlue = (BitAND($iFromColor, 0xff) / 255)

			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			$tColorStop = $atColorStop[0] ; StopOffset 0 is the "From Color" Value.

			$tStopColor = $tColorStop.StopColor()

			$tStopColor.Red = $nRed
			$tStopColor.Green = $nGreen
			$tStopColor.Blue = $nBlue

			$tColorStop.StopColor = $tStopColor

			$atColorStop[0] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($iToColor <> Null) Then
		If Not __LO_IntIsBetween($iToColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tStyleGradient.EndColor = $iToColor

		If __LO_VersionCheck(7.6) Then
			$nRed = (BitAND(BitShift($iToColor, 16), 0xff) / 255)
			$nGreen = (BitAND(BitShift($iToColor, 8), 0xff) / 255)
			$nBlue = (BitAND($iToColor, 0xff) / 255)

			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

			$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; Last StopOffset is the "To Color" Value.

			$tStopColor = $tColorStop.StopColor()

			$tStopColor.Red = $nRed
			$tStopColor.Green = $nGreen
			$tStopColor.Blue = $nBlue

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

	If ($oBackground.FillGradientName() = "") Then
		$sGradName = __LOImpress_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

		$oBackground.FillGradientName = $sGradName
		If ($oBackground.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)
	EndIf

	$oBackground.FillGradient = $tStyleGradient
	$oSlide.Background = $oBackground

	; Error checking
	$iError = (__LO_VarsAreNull($iType)) ? $iError : ($oSlide.Background.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2))
	$iError = (__LO_VarsAreNull($iXCenter)) ? $iError : ($oSlide.Background.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8))
	$iError = (__LO_VarsAreNull($iYCenter)) ? $iError : ($oSlide.Background.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16))
	$iError = (__LO_VarsAreNull($iAngle)) ? $iError : (($oSlide.Background.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? $iError : ($oSlide.Background.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64))
	$iError = (__LO_VarsAreNull($iFromColor)) ? $iError : ($oSlide.Background.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128))
	$iError = (__LO_VarsAreNull($iToColor)) ? $iError : ($oSlide.Background.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256))
	$iError = (__LO_VarsAreNull($iFromIntense)) ? $iError : ($oSlide.Background.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512))
	$iError = (__LO_VarsAreNull($iToIntense)) ? $iError : ($oSlide.Background.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideBackGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideBackTransparency
; Description ...: Set or retrieve Transparency settings for a Slide.
; Syntax ........: _LOImpress_SlideBackTransparency(ByRef $oSlide[, $iTransparency = Null])
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve parent Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the slide, -1 is returned.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideBackTransparency(ByRef $oSlide, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oBackground, $oDoc

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oSlide.Background()

	If __LO_VarsAreNull($iTransparency) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, -1) ; No background present.

		Return SetError($__LO_STATUS_SUCCESS, 1, $oBackground.FillTransparence())
	EndIf

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oDoc = $oSlide.MasterPage.Forms.Parent()
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	$oBackground.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oBackground.FillTransparence = $iTransparency

	$oSlide.Background = $oBackground

	$iError = ($oSlide.Background.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideBackTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideBackTransparencyGradient
; Description ...: Set or retrieve the Slide's transparency gradient settings.
; Syntax ........: _LOImpress_SlideBackTransparencyGradient(ByRef $oSlide[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3. Call with $LOI_GRAD_TYPE_OFF to turn Transparency Gradient off.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] an integer value (0-100). Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iStart              - [optional] an integer value (0-100). Default is Null. The transparency value for the beginning point of the gradient, where 0% is fully opaque and 100% is fully transparent.
;                  $iEnd                - [optional] an integer value (0-100). Default is Null. The transparency value for the endpoint of the gradient, where 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iType Not an Integer, less than -1 or greater than 5. See constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iXCenter Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 4 Return 0 = $iYCenter Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 5 Return 0 = $iAngle Not an Integer, less than 0 or greater than 359.
;                  @Error 1 @Extended 6 Return 0 = $iTransitionStart Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iStart Not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iEnd Not an Integer, less than 0 or greater than 100.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 =
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving "FillTransparenceGradient" Struct.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve parent Document.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving Color Stop Array for "From" color
;                  @Error 3 @Extended 4 Return 0 = Error retrieving Color Stop Array for "To" color
;                  @Error 3 @Extended 5 Return 0 = Error creating Transparency Gradient name.
;                  @Error 3 @Extended 6 Return 0 = Error setting Transparency Gradient name.
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
;                  @Error 0 @Extended 1 Return -1 = Success. All optional parameters were called with Null no background is currently active for the slide. Returning -1.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideBackTransparencyGradient(ByRef $oSlide, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop
	Local $oBackground, $oDoc
	Local $fValue

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oSlide.Background()

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 2, -1)

		$tGradient = $oBackground.FillTransparenceGradient()
		If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LO_ArrayFill($aiTransparent, $tGradient.Style(), $tGradient.XOffset(), $tGradient.YOffset(), _
				($tGradient.Angle() / 10), $tGradient.Border(), __LOImpress_TransparencyGradientConvert(Null, $tGradient.StartColor()), _
				__LOImpress_TransparencyGradientConvert(Null, $tGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	$tGradient = $oBackground.FillTransparenceGradient()
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iType <> Null) Then
		If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oBackground.FillTransparenceGradientName = ""
			$oSlide.Background = $oBackground

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

		$tGradient.Angle = ($iAngle * 10) ; Angle is set in thousands
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
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

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

	If ($oBackground.FillTransparenceGradientName() = "") Then
		$sTGradName = __LOImpress_TransparencyGradientNameInsert($oDoc, $tGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		$oBackground.FillTransparenceGradientName = $sTGradName
		If ($oBackground.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
	EndIf

	$oBackground.FillTransparenceGradient = $tGradient
	$oSlide.Background = $oBackground

	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oSlide.Background.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oSlide.Background.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oSlide.Background.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((($oSlide.Background.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oSlide.Background.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iStart)) ? ($iError) : (($oSlide.Background.FillTransparenceGradient.StartColor() = __LOImpress_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iEnd)) ? ($iError) : (($oSlide.Background.FillTransparenceGradient.EndColor() = __LOImpress_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideBackTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideCopy
; Description ...: Create a copy of a slide.
; Syntax ........: _LOImpress_SlideCopy(ByRef $oSlide[, $iPos = Null])
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iPos                - [optional] an integer value. Default is Null. The position to insert the new slide in the collection of slides. 0 Based. See remarks.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iPos not an Integer, less than 0 or greater than number of slides.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document.
;                  @Error 3 @Extended 2 Return 0 = Failed to copy slide.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify copied slide's position.
;                  @Error 3 @Extended 4 Return 0 = Failed to backup currently active slide.
;                  @Error 3 @Extended 5 Return 0 = Failed to move copied slide.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully copied the slide, returning the new slide's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The copied slide is inserted after the slide to be copied.
;                  If $iPos is called with Null, the slide is left in the position described above. Otherwise, due to limitations in the API, some dispatches are executed to move the slide. The current slide will temporarily be set to the new slide in order to move it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideCopy(ByRef $oSlide, $iPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc, $oNewSlide, $oCurrSlide, $oServiceManager, $oDispatcher
	Local $iNewPos, $iMove
	Local $sDispatch
	Local $aArray[0]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iPos <> Null) And Not __LO_IntIsBetween($iPos, 0, $oDoc.DrawPages.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oNewSlide = $oDoc.Duplicate($oSlide)
	If Not IsObj($oNewSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Failed to copy Slide.

	If ($iPos <> Null) Then
		For $i = 0 To $oDoc.DrawPages.getCount() - 1
			If ($oDoc.DrawPages.getByIndex($i) = $oNewSlide) Then
				$iNewPos = $i
				ExitLoop
			EndIf
			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		If Not IsInt($iNewPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$iMove = ($iNewPos > $iPos) ? ($iNewPos - $iPos) : (($iNewPos < $iPos) ? ($iPos - $iNewPos) : (0)) ; 0 = NewPos and current Pos are the same.
		$sDispatch = ($iNewPos > $iPos) ? (".uno:MovePageUp") : (".uno:MovePageDown")
		If ($iPos = 0) Then ; Move Slide to beginning.
			$iMove = 1 ; Set to 1 so it will be called once.
			$sDispatch = ".uno:MovePageFirst"

		ElseIf ($iPos = $oDoc.DrawPages.getCount() - 1) Then ; Move slide to end.
			$iMove = 1 ; Set to 1 so it will be called once.
			$sDispatch = ".uno:MovePageLast"
		EndIf

		$oCurrSlide = $oDoc.getCurrentController.CurrentPage() ; Backup current slide
		If Not IsObj($oCurrSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oDoc.getCurrentController.setCurrentPage($oNewSlide)

		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
		If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		For $i = 0 To $iMove - 1
			$oDispatcher.executeDispatch($oDoc.CurrentController(), $sDispatch, "", 0, $aArray)
			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		For $i = 0 To $oDoc.DrawPages.getCount() - 1
			If ($oDoc.DrawPages.getByIndex($i) = $oNewSlide) Then
				$iNewPos = $i
				ExitLoop
			EndIf
			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		$oDoc.getCurrentController.setCurrentPage($oCurrSlide) ; Restore current slide.

		If ($iNewPos <> $iPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNewSlide)
EndFunc   ;==>_LOImpress_SlideCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideCurrent
; Description ...: Set or Retrieve the currently active slide.
; Syntax ........: _LOImpress_SlideCurrent(ByRef $oDoc[, $oSlide = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oSlide              - [optional] an object. Default is Null. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSlide not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current slide's Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oSlide
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Object = Success. All optional parameters were called with Null, returning currently active slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current slide.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideCurrent(ByRef $oDoc, $oSlide = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCurrSlide
	Local $iError

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($oSlide) Then
		$oCurrSlide = $oDoc.getCurrentController.CurrentPage()
		If Not IsObj($oCurrSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $oCurrSlide)
	EndIf

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.getCurrentController.setCurrentPage($oSlide)
	$iError = ($oDoc.getCurrentController.CurrentPage() = $oSlide) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideDeleteByIndex
; Description ...: Delete a slide by index.
; Syntax ........: _LOImpress_SlideDeleteByIndex(ByRef $oDoc, $iSlide)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iSlide              - an integer value. The slide to delete. 0 based.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iSlide not an Integer, less than 0 or greater than number of slides minus one.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve count of slides.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve slide's Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete slide.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Slide was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideDeleteByIndex(ByRef $oDoc, $iSlide)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSlide
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iSlide, 0, $oDoc.DrawPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iCount = $oDoc.DrawPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSlide = $oDoc.DrawPages.getByIndex($iSlide)
	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.DrawPages.remove($oSlide)
	If ($iCount = $oDoc.DrawPages.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to delete because the count is the same.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideDeleteByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideDeleteByObj
; Description ...: Delete a slide using its Object.
; Syntax ........: _LOImpress_SlideDeleteByObj(ByRef $oSlide)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of slides.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete slide.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Slide was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideDeleteByObj(ByRef $oSlide)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc
	Local $iCount

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oDoc.DrawPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.DrawPages.Remove($oSlide)
	If ($iCount = $oDoc.DrawPages.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to delete because the count is the same.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideDeleteByObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideGetByIndex
; Description ...: Retrieve a Slide's Object by index.
; Syntax ........: _LOImpress_SlideGetByIndex(ByRef $oDoc, $iSlide)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iSlide              - an integer value. The slide to retrieve. 0 based.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iSlide not an Integer, less than 0 or greater than number of slides minus one.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve requested slide.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested slide's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideGetByIndex(ByRef $oDoc, $iSlide)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSlide

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iSlide, 0, $oDoc.DrawPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSlide = $oDoc.DrawPages.getByIndex($iSlide)
	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSlide)
EndFunc   ;==>_LOImpress_SlideGetByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMove
; Description ...: Move a slide in the collection of slides.
; Syntax ........: _LOImpress_SlideMove(ByRef $oSlide, $iPos)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iPos                - an integer value. The position to move the slide to in the collection of slides. 0 Based. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iPos not an Integer, less than 0 or greater than number of slides minus 1.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify slide's current position.
;                  @Error 3 @Extended 3 Return 0 = Failed to backup currently active slide.
;                  @Error 3 @Extended 4 Return 0 = Failed to move slide.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Slide was successfully moved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Due to limitations in the API, some dispatches are executed to move the slide. The current slide will temporarily be set to the new slide in order to move it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMove(ByRef $oSlide, $iPos)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc, $oCurrSlide, $oServiceManager, $oDispatcher
	Local $iCurrPos, $iMove
	Local $sDispatch
	Local $aArray[0]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iPos <> Null) And Not __LO_IntIsBetween($iPos, 0, $oDoc.DrawPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	For $i = 0 To $oDoc.DrawPages.getCount() - 1
		If ($oDoc.DrawPages.getByIndex($i) = $oSlide) Then
			$iCurrPos = $i
			ExitLoop
		EndIf
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If Not IsInt($iCurrPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iMove = ($iCurrPos > $iPos) ? ($iCurrPos - $iPos) : (($iCurrPos < $iPos) ? ($iPos - $iCurrPos) : (0))     ; 0 = CurrPos and New Pos are the same.
	$sDispatch = ($iCurrPos > $iPos) ? (".uno:MovePageUp") : (".uno:MovePageDown")
	If ($iPos = 0) Then     ; Move Slide to beginning.
		$iMove = 1    ; Set to 1 so it will be called once.
		$sDispatch = ".uno:MovePageFirst"

	ElseIf ($iPos = $oDoc.DrawPages.getCount() - 1) Then     ; Move slide to end.
		$iMove = 1    ; Set to 1 so it will be called once.
		$sDispatch = ".uno:MovePageLast"
	EndIf

	$oCurrSlide = $oDoc.getCurrentController.CurrentPage()     ; Backup current slide
	If Not IsObj($oCurrSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oDoc.getCurrentController.setCurrentPage($oSlide)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	For $i = 0 To $iMove - 1
		$oDispatcher.executeDispatch($oDoc.CurrentController(), $sDispatch, "", 0, $aArray)
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	For $i = 0 To $oDoc.DrawPages.getCount() - 1
		If ($oDoc.DrawPages.getByIndex($i) = $oSlide) Then
			$iCurrPos = $i
			ExitLoop
		EndIf
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$oDoc.getCurrentController.setCurrentPage($oCurrSlide)     ; Restore current slide.

	If ($iCurrPos <> $iPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlidesGetCount
; Description ...: Retrieve a count of slides.
; Syntax ........: _LOImpress_SlidesGetCount(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of slides.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of slides contained in the document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlidesGetCount(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oDoc.DrawPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOImpress_SlidesGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideShapesGetList
; Description ...: Retrieve an array of Shapes (Text Boxes, smileys, images etc) contained in a Slide.
; Syntax ........: _LOImpress_SlideShapesGetList(ByRef $oSlide[, $iTypes = $LOI_SHAPE_TYPE_ALL])
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iTypes              - [optional] an integer value (0-511). Default is $LOI_SHAPE_TYPE_ALL. The type of Shapes to return in the Array. Can be BitOR'd. See Constants, $LOI_SHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTypes not an Integer, less than 1 or greater than 511. See Constants, $LOI_SHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify Shape Type.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. A two columned Array containing the Shape Objects contained in the Slide. See Remarks. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array returned has two columns. The first column is the shape Object. The second column is the Shape Type, corresponding to one of the Constants $LOI_SHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideShapesGetList(ByRef $oSlide, $iTypes = $LOI_SHAPE_TYPE_ALL)
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
EndFunc   ;==>_LOImpress_SlideShapesGetList
