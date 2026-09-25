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
; _LOImpress_SlideExists
; _LOImpress_SlideFooter
; _LOImpress_SlideFormat
; _LOImpress_SlideGetObjByIndex
; _LOImpress_SlideGetObjByName
; _LOImpress_SlideHandoutFooter
; _LOImpress_SlideHandoutFormat
; _LOImpress_SlideHandoutGetObj
; _LOImpress_SlideHandoutHeader
; _LOImpress_SlideHandoutLayout
; _LOImpress_SlideHandoutMargins
; _LOImpress_SlideLayout
; _LOImpress_SlideMargins
; _LOImpress_SlideMasterAdd
; _LOImpress_SlideMasterBackColor
; _LOImpress_SlideMasterBackFillStyle
; _LOImpress_SlideMasterBackGradient
; _LOImpress_SlideMasterBackTransparency
; _LOImpress_SlideMasterBackTransparencyGradient
; _LOImpress_SlideMasterCurrent
; _LOImpress_SlideMasterDeleteByIndex
; _LOImpress_SlideMasterDeleteByObj
; _LOImpress_SlideMasterExists
; _LOImpress_SlideMasterFormat
; _LOImpress_SlideMasterGetObjByIndex
; _LOImpress_SlideMasterGetObjByName
; _LOImpress_SlideMasterMargins
; _LOImpress_SlideMasterName
; _LOImpress_SlideMasterNotesGetObj
; _LOImpress_SlideMastersGetCount
; _LOImpress_SlideMastersGetNames
; _LOImpress_SlideMove
; _LOImpress_SlideName
; _LOImpress_SlideNotesFooter
; _LOImpress_SlideNotesFormat
; _LOImpress_SlideNotesGetObj
; _LOImpress_SlideNotesHeader
; _LOImpress_SlideNotesMargins
; _LOImpress_SlidesGetCount
; _LOImpress_SlidesGetNames
; _LOImpress_SlideshowActiveSettings
; _LOImpress_SlideshowCustomCreate
; _LOImpress_SlideshowCustomDelete
; _LOImpress_SlideshowCustomModify
; _LOImpress_SlideshowCustomSetName
; _LOImpress_SlideshowIsRunning
; _LOImpress_SlideshowPresentationControl
; _LOImpress_SlideshowsCustomGetNames
; _LOImpress_SlideshowSettingsMode
; _LOImpress_SlideshowSettingsOptions
; _LOImpress_SlideshowSettingsRange
; _LOImpress_SlideshowStart
; _LOImpress_SlideshowStop
; _LOImpress_SlideSoundsGetNames
; _LOImpress_SlideTransition
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideAdd
; Description ...: Add a slide to a presentation.
; Syntax ........: _LOImpress_SlideAdd(ByRef $oDoc[, $iPos = Null[, $sName = ""]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iPos                - [optional] Default is Null. The position to insert the new slide in the collection of slides. 0 Based. See remarks.
;                  $sName               - [optional] Default is "". The unique name of the Slide. If called with an empty string, LibreOffice automatically names it.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning new slide's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPos not an Integer, less than 0 or greater than number of slides.
;                  @Error: 1, @Extended: 3 = $sName not a String.
;                  @Error: 1, @Extended: 4 = Name called in $sName already exists.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to create a slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $iPos is called with Null, the new slide is inserted at the end.
;                  Call $iPos with the last slide index to insert the slide at the end. Call $iPos with 0 to insert the new slide in the first slide position.
;                  Due to limitations in the API, I have made a small workaround for inserting a slide at the beginning. A dispatch is executed to move the slide to the beginning. The current slide will temporarily be set to the new slide in order to move it.
; Related .......: _LOImpress_SlideDeleteByIndex, _LOImpress_SlideDeleteByObj, _LOImpress_SlideMasterAdd, _LOImpress_SlideExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideAdd(ByRef $oDoc, $iPos = Null, $sName = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSlide, $oServiceManager, $oDispatcher, $oCurrSlide
	Local $bMoveToFirst = False
	Local $aArray[0]
	Local $iCurrView

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($iPos = Null) Then $iPos = $oDoc.DrawPages.getCount()
	If Not __LO_IntIsBetween($iPos, 0, $oDoc.DrawPages.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sName <> "") And _LOImpress_SlideExists($oDoc, $sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iPos -= 1 ; -1 because when 0 is called in insertNewByIndex, it inserts it in position 1, etc. Also there is no way to insert a new slide at position 0, so I made a workaround.

	If ($iPos = -1) Then
		$iPos = 0
		$bMoveToFirst = True
	EndIf

	$oSlide = $oDoc.DrawPages.insertNewByIndex($iPos)
	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $bMoveToFirst Then
		$oCurrSlide = $oDoc.getCurrentController.CurrentPage() ; Backup current slide and view mode

		$iCurrView = __LOImpress_DocCurrView($oDoc)

		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
		If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oDoc.getCurrentController.setCurrentPage($oSlide)

		$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:MovePageFirst", "", 0, $aArray)

		If IsObj($oCurrSlide) Then $oDoc.getCurrentController.setCurrentPage($oCurrSlide) ; Restore current slide and view mode.

		If IsInt($iCurrView) Then __LOImpress_DocCurrView($oDoc, $iCurrView)
	EndIf

	If ($sName <> "") Then
		$oSlide.Name = $sName
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSlide)
EndFunc   ;==>_LOImpress_SlideAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideBackColor
; Description ...: Set or Retrieve the Slide's background color.
; Syntax ........: _LOImpress_SlideBackColor(ByRef $oSlide[, $iColor = Null])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Slide background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Integer
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting as an Integer value. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current color value.
;                  @Error: 3, @Extended: 2 = Failed to retrieve parent Document.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColor
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the slide, the Constant $LO_COLOR_OFF is returned.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_SlideBackFillStyle, _LOImpress_SlideBackGradient, _LOImpress_SlideMasterBackColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideBackColor(ByRef $oSlide, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBackground, $oDoc
	Local $iError = 0, $iCurColor

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oSlide.Background()

	If __LO_VarsAreNull($iColor) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, $LO_COLOR_OFF) ; If no background is set, this will be void, instead of an Object.

		$iCurColor = __LOImpress_ColorRemoveAlpha($oBackground.FillColor())
		If Not IsInt($iCurColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurColor)
	EndIf

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oDoc = __LOImpress_GetParentDoc($oSlide)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	$oBackground.FillStyle = $LOI_AREA_FILL_STYLE_SOLID
	$oBackground.FillColor = $iColor

	$oSlide.Background = $oBackground
	$iError = ($oSlide.Background.FillColor() = $iColor) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideBackFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOImpress_SlideBackFillStyle(ByRef $oSlide[, $bFillOff = False])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $bFillOff            - [optional] Default is False. If True, the Fill style will be set to Off. See remarks.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOI_AREA_FILL_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 0, @Extended: 1, Return: 0 = Success. Fill style was successfully turned off.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $bFillOff not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Fill Style.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
;                  When the Fill style is disabled for a Slide, the Fill properties are completely removed. This is how Impress works normally.
;                  $bFillOff will do nothing if it is called with False, and is not, of course, returned when retrieving the FillStyle value.
; Related .......: _LOImpress_SlideBackColor, _LOImpress_SlideBackGradient, _LOImpress_SlideMasterBackFillStyle
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
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
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
;                  @Error: 0, @Extended: 2, Return: -1 = Success. All optional parameters were called with Null, no background is currently active for the slide. Returning -1.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
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
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillGradient" Struct.
;                  @Error: 3, @Extended: 2 = Error retrieving Parent Document.
;                  @Error: 3, @Extended: 3 = Error retrieving Background Object.
;                  @Error: 3, @Extended: 4 = Error retrieving Color Stop Array for "From" color
;                  @Error: 3, @Extended: 5 = Error retrieving Color Stop Array for "To" color
;                  @Error: 3, @Extended: 6 = Error creating Gradient Name.
;                  @Error: 3, @Extended: 7 = Error setting Gradient Name.
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
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_SlideBackColor, _LOImpress_SlideBackFillStyle, _LOImpress_SlideMasterBackGradient
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

	$oDoc = __LOImpress_GetParentDoc($oSlide)
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

	If ($oBackground.FillGradientName() = "") Or __LOImpress_GradientIsModified($tStyleGradient, $oBackground.FillGradientName()) Then
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
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iTransparency       - [optional] (0-100) Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Transparency value.
;                  @Error: 3, @Extended: 2 = Failed to retrieve parent Document.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the slide, -1 is returned.
; Related .......: _LOImpress_SlideBackTransparencyGradient, _LOImpress_SlideMasterBackTransparency
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideBackTransparency(ByRef $oSlide, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurTransp
	Local $oBackground, $oDoc

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oSlide.Background()

	If __LO_VarsAreNull($iTransparency) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, -1) ; No background present.

		$iCurTransp = $oBackground.FillTransparence()
		If Not IsInt($iCurTransp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurTransp)
	EndIf

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oDoc = __LOImpress_GetParentDoc($oSlide)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

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
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iType               - [optional] (-1-5) Default is Null. The type of transparency gradient to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3. Call with $LOI_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error: 0, @Extended: 1, Return: -1 = Success. All optional parameters were called with Null no background is currently active for the slide. Returning -1.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iType Not an Integer, less than -1 or greater than 5. See constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iXCenter Not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 4 = $iYCenter Not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 5 = $iAngle Not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 6 = $iTransitionStart Not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iStart Not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iEnd Not an Integer, less than 0 or greater than 100.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillTransparenceGradient" Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve parent Document.
;                  @Error: 3, @Extended: 3 = Error retrieving Color Stop Array for "From" color
;                  @Error: 3, @Extended: 4 = Error retrieving Color Stop Array for "To" color
;                  @Error: 3, @Extended: 5 = Error creating Transparency Gradient name.
;                  @Error: 3, @Extended: 6 = Error setting Transparency Gradient name.
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
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_SlideBackTransparency, _LOImpress_SlideMasterBackTransparencyGradient
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

	$oDoc = __LOImpress_GetParentDoc($oSlide)
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
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iPos                - [optional] Default is Null. The position to insert the new slide in the collection of slides. 0 Based. See remarks.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully copied the slide, returning the new slide's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iPos not an Integer, less than 0 or greater than number of slides.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Parent Document.
;                  @Error: 3, @Extended: 2 = Failed to copy slide.
;                  @Error: 3, @Extended: 3 = Failed to identify copied slide's position.
;                  @Error: 3, @Extended: 4 = Failed to move copied slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The copied slide is inserted after the slide to be copied.
;                  If $iPos is called with Null, the slide is left in the position described above. Otherwise, due to limitations in the API, some dispatches are executed to move the slide. The current slide will temporarily be set to the new slide in order to move it.
; Related .......: _LOImpress_SlideAdd, _LOImpress_SlideDeleteByIndex, _LOImpress_SlideDeleteByObj, _LOImpress_SlideMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideCopy(ByRef $oSlide, $iPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc, $oNewSlide, $oCurrSlide, $oServiceManager, $oDispatcher
	Local $iNewPos, $iMove, $iCurrView
	Local $sDispatch
	Local $aArray[0]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = __LOImpress_GetParentDoc($oSlide)
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

		$iCurrView = __LOImpress_DocCurrView($oDoc)

		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
		If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oDoc.getCurrentController.setCurrentPage($oNewSlide)

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

		If IsObj($oCurrSlide) Then $oDoc.getCurrentController.setCurrentPage($oCurrSlide) ; Restore current slide and view mode.

		If IsInt($iCurrView) Then __LOImpress_DocCurrView($oDoc, $iCurrView)

		If ($iNewPos <> $iPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNewSlide)
EndFunc   ;==>_LOImpress_SlideCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideCurrent
; Description ...: Set or Retrieve the currently active slide or master slide.
; Syntax ........: _LOImpress_SlideCurrent(ByRef $oDoc[, $oObj = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oObj                - [optional] Default is Null. A Slide or Master Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, _LOImpress_SlideCopy, _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
; Return values .: Success: 1 or Object
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Object = Success. All optional parameters were called with Null, returning currently active slide. @Extended page's type, see remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oObj not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current slide's Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oObj
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current slide.
;                  If this function fails to return an Object with processing error 1, it is possible the current view mode is set to Slide sorter.
;                  You can only set the current slide to either a Master slide or a normal slide. To change views to Notes, Handouts etc., see _LOImpress_DocView.
;                  If the current view mode is set to Slide outline or Slide Notes, the current slide Object is returned. If the current view mode is set to Master Slide Notes or Master Slide Handout, the current Master slide Object is returned.
;                  When retrieving the current page, @Extended will be set to either $LOI_PAGE_VIEW_SLIDE or $LOI_PAGE_VIEW_MASTER. See Constants, $LOI_PAGE_VIEW_* as defined in LibreOfficeImpress_Constants.au3. Use _LOImpress_DocView to determine the current view mode active.
; Related .......: _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, _LOImpress_SlideMasterCurrent, _LOImpress_DocView
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideCurrent(ByRef $oDoc, $oObj = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCurrSlide
	Local $iError, $iPageType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($oObj) Then
		$oCurrSlide = $oDoc.getCurrentController.CurrentPage()
		If Not IsObj($oCurrSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Could be because the current mode is set to Slide Sorter.

		$iPageType = ($oDoc.getCurrentController.IsMasterPageMode()) ? ($LOI_PAGE_VIEW_MASTER) : ($LOI_PAGE_VIEW_SLIDE)

		Return SetError($__LO_STATUS_SUCCESS, $iPageType, $oCurrSlide)
	EndIf

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.getCurrentController.setCurrentPage($oObj)
	$iError = ($oDoc.getCurrentController.CurrentPage() = $oObj) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideDeleteByIndex
; Description ...: Delete a slide by index.
; Syntax ........: _LOImpress_SlideDeleteByIndex(ByRef $oDoc, $iSlide)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iSlide              - The slide to delete. 0 based.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Slide was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iSlide not an Integer, less than 0 or greater than number of slides minus one.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve count of slides.
;                  @Error: 3, @Extended: 2 = Failed to retrieve slide's Object.
;                  @Error: 3, @Extended: 3 = Failed to delete slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideDeleteByObj, _LOImpress_SlidesGetCount, _LOImpress_SlideMasterDeleteByIndex
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
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Slide was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Parent Document.
;                  @Error: 3, @Extended: 2 = Failed to retrieve count of slides.
;                  @Error: 3, @Extended: 3 = Failed to delete slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideDeleteByIndex, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, _LOImpress_SlideMasterDeleteByObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideDeleteByObj(ByRef $oSlide)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc
	Local $iCount

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = __LOImpress_GetParentDoc($oSlide)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oDoc.DrawPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.DrawPages.Remove($oSlide)
	If ($oDoc.DrawPages.getCount() = $iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to delete because the count is the same.

	$oSlide = Null

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideDeleteByObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideExists
; Description ...: Check whether a slide with a certain name exists in a document.
; Syntax ........: _LOImpress_SlideExists(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sName               - The slide name to check for.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Document contains a Slide with the called name, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query for Slide name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideAdd, _LOImpress_SlideName, _LOImpress_SlidesGetNames, _LOImpress_SlideMasterExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideExists(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bExists

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$bExists = $oDoc.Links.getByName("Slide").Links.hasByName($sName)
	If Not IsBool($bExists) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bExists)
EndFunc   ;==>_LOImpress_SlideExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideFooter
; Description ...: Set or Retrieve Slide Footer settings.
; Syntax ........: _LOImpress_SlideFooter(ByRef $oSlide[, $bDateTime = Null[, $bDateTimeIsFixed = Null[, $sDateTimeValue = Null[, $iDateTimeFormat = Null[, $bFooter = Null[, $sFooterText = Null[, $bSlideNum = Null]]]]]]])
; Parameters ....: $oSlide              -  A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $bDateTime           - [optional] Default is Null. If True, a Date or Time entry is added to the footer of the slide.
;                  $bDateTimeIsFixed    - [optional] Default is Null. If True, the Date or Time entry is fixed.
;                  $sDateTimeValue      - [optional] Default is Null. If $bDateTimeIsFixed is True, this is the custom date or time value to display.
;                  $iDateTimeFormat     - [optional] (4-112) Default is Null. If $bDateTimeIsFixed is False, the format to display the Date or Time in. See Constants, $LOI_SLIDE_DT_FMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bFooter             - [optional] Default is Null. If True, a Footer entry is added to the footer of the slide.
;                  $sFooterText         - [optional] Default is Null. If $bFooter is True, the text to display in the footer of the Slide.
;                  $bSlideNum           - [optional] Default is Null. If True, a current Slide number is added to the footer of the slide.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $bDateTime not a Boolean.
;                  @Error: 1, @Extended: 3 = $bDateTimeIsFixed not a Boolean.
;                  @Error: 1, @Extended: 4 = $sDateTimeValue not a String.
;                  @Error: 1, @Extended: 5 = $iDateTimeFormat not an Integer, less than 4 or greater than 9 but not equal to one of the constant values. See Constants, $LOI_SLIDE_DT_FMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 6 = $bFooter not a Boolean.
;                  @Error: 1, @Extended: 7 = $sFooterText not a String.
;                  @Error: 1, @Extended: 8 = $bSlideNum not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bDateTime
;                  |                               2 = Error setting $bDateTimeIsFixed
;                  |                               4 = Error setting $sDateTimeValue
;                  |                               8 = Error setting $iDateTimeFormat
;                  |                               16 = Error setting $bFooter
;                  |                               32 = Error setting $sFooterText
;                  |                               64 = Error setting $bSlideNum
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When retrieving current setting values, both $sDateTimeValue and $iDateTimeFormat may return a value. To determine which is currently valid, check $bDateTimeIsFixed. If $bDateTimeIsFixed is True, $sDateTimeValue is valid, else $iDateTimeFormat. If $bDateTime is false, neither will be valid.
;                  Skip first slide, and Apply to all are not added to this function as they are not actual settings. The user can simulate these easily by making a loop to apply it to all slides, and skip the first slide if required.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, _LOImpress_SlideHandoutFooter, _LOImpress_SlideNotesFooter
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideFooter(ByRef $oSlide, $bDateTime = Null, $bDateTimeIsFixed = Null, $sDateTimeValue = Null, $iDateTimeFormat = Null, $bFooter = Null, $sFooterText = Null, $bSlideNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFooter[7]
	Local $sAllowed = $LOI_SLIDE_DT_FMT_24H_HM & ":" & $LOI_SLIDE_DT_FMT_MMDDYY_24H_HM & ":" & $LOI_SLIDE_DT_FMT_24H_HMS & ":" & $LOI_SLIDE_DT_FMT_12H_HM_AMPM & ":" & $LOI_SLIDE_DT_FMT_MMDDYY_12H_HM_AMPM & ":" & $LOI_SLIDE_DT_FMT_12H_HMS_AMPM

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bDateTime, $bDateTimeIsFixed, $sDateTimeValue, $iDateTimeFormat, $bFooter, $sFooterText, $bSlideNum) Then
		__LO_ArrayFill($avFooter, $oSlide.IsDateTimeVisible(), $oSlide.IsDateTimeFixed(), $oSlide.DateTimeText(), $oSlide.DateTimeFormat(), _
				$oSlide.IsFooterVisible(), $oSlide.FooterText(), $oSlide.IsPageNumberVisible())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFooter)
	EndIf

	If ($bDateTime <> Null) Then
		If Not IsBool($bDateTime) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oSlide.IsDateTimeVisible = $bDateTime

		$iError = ($oSlide.IsDateTimeVisible() = $bDateTime) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bDateTimeIsFixed <> Null) Then
		If Not IsBool($bDateTimeIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oSlide.IsDateTimeFixed = $bDateTimeIsFixed

		$iError = ($oSlide.IsDateTimeFixed() = $bDateTimeIsFixed) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sDateTimeValue <> Null) Then
		If Not IsString($sDateTimeValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSlide.DateTimeText = $sDateTimeValue

		$iError = ($oSlide.DateTimeText() = $sDateTimeValue) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iDateTimeFormat <> Null) Then
		If Not __LO_IntIsBetween($iDateTimeFormat, $LOI_SLIDE_DT_FMT_MMDDYY, $LOI_SLIDE_DT_FMT_DOW_MMMM_DD_YYYY, "", $sAllowed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oSlide.DateTimeFormat = $iDateTimeFormat

		$iError = ($oSlide.DateTimeFormat() = $iDateTimeFormat) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bFooter <> Null) Then
		If Not IsBool($bFooter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oSlide.IsFooterVisible = $bFooter

		$iError = ($oSlide.IsFooterVisible() = $bFooter) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($sFooterText <> Null) Then
		If Not IsString($sFooterText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oSlide.FooterText = $sFooterText

		$iError = ($oSlide.FooterText() = $sFooterText) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bSlideNum <> Null) Then
		If Not IsBool($bSlideNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oSlide.IsPageNumberVisible = $bSlideNum

		$iError = ($oSlide.IsPageNumberVisible() = $bSlideNum) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideFormat
; Description ...: Set or Retrieve the slide format settings.
; Syntax ........: _LOImpress_SlideFormat(ByRef $oSlide[, $iWidth = Null[, $iHeight = Null[, $iOrientation = Null]]])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iWidth              - [optional] Default is Null. The Width of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOI_PAGE_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iHeight             - [optional] Default is Null. The Height of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOI_PAGE_HEIGHT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iOrientation        - [optional] (0-1) Default is Null. The page orientation. See Constants, $LOI_PAGE_ORIENT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 4 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants, $LOI_PAGE_ORIENT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current slide width.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iHeight
;                  |                               4 = Error setting $iOrientation
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  When modifying the page format, the shapes etc., aren't readjusted as they are in LibreOffice UI.
;                  I am unable to find the properties to set for "FitObject to Paper Format", "Background covers margins", "Slide numbers", and "Paper tray".
; Related .......: _LO_UnitConvert, _LOImpress_SlideLayout, _LOImpress_SlideMargins, _LOImpress_SlideHandoutFormat, _LOImpress_SlideMasterFormat, _LOImpress_SlideNotesFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideFormat(ByRef $oSlide, $iWidth = Null, $iHeight = Null, $iOrientation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_Format($oSlide, $iWidth, $iHeight, $iOrientation)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_SlideFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideGetObjByIndex
; Description ...: Retrieve a Slide's Object by index.
; Syntax ........: _LOImpress_SlideGetObjByIndex(ByRef $oDoc, $iSlide)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iSlide              - The slide to retrieve. 0 based.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested slide's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iSlide not an Integer, less than 0 or greater than number of slides minus one.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideGetObjByName, _LOImpress_SlidesGetCount, _LOImpress_SlideMasterGetObjByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideGetObjByIndex(ByRef $oDoc, $iSlide)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSlide

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iSlide, 0, $oDoc.DrawPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSlide = $oDoc.DrawPages.getByIndex($iSlide)
	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSlide)
EndFunc   ;==>_LOImpress_SlideGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideGetObjByName
; Description ...: Retrieve a Slide's Object by name.
; Syntax ........: _LOImpress_SlideGetObjByName(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sName               - The Slide's name to retrieve the Object for.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested Slide's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Slide name called in $sName not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested Slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideGetObjByIndex, _LOImpress_SlidesGetNames, _LOImpress_SlideMasterGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideGetObjByName(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSlide

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.Links.getByName("Slide").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSlide = $oDoc.Links.getByName("Slide").Links.getByName($sName)
	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSlide)
EndFunc   ;==>_LOImpress_SlideGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideFooter
; Description ...: Set or Retrieve handout page Footer settings.
; Syntax ........: _LOImpress_SlideFooter(ByRef $oHandout[, $bFooter = Null[, $sFooterText = Null[, $bSlideNum = Null]]])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LOImpress_SlideHandoutGetObj function.
;                  $bFooter             - [optional] Default is Null. If True, a Footer entry is added to the footer of the page.
;                  $sFooterText         - [optional] Default is Null. If $bFooter is True, the text to display in the footer of the page.
;                  $bSlideNum           - [optional] Default is Null. If True, a current Slide number is added to the footer of the page.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oHandout not an Object.
;                  @Error: 1, @Extended: 2 = $bFooter not a Boolean.
;                  @Error: 1, @Extended: 3 = $sFooterText not a String.
;                  @Error: 1, @Extended: 4 = $bSlideNum not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFooter
;                  |                               2 = Error setting $sFooterText
;                  |                               4 = Error setting $bSlideNum
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Apply to all is not added to this function as they it is not an actual setting. The user can simulate this easily by making a loop to apply it to all slides.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  During basic testing, while the settings were successfully set, LibreOffice seems to ignore footer values set for handout pages.
; Related .......: _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, _LOImpress_SlideFooter, _LOImpress_SlideNotesFooter
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideHandoutFooter(ByRef $oHandout, $bFooter = Null, $sFooterText = Null, $bSlideNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFooter[3]

	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bFooter, $sFooterText, $bSlideNum) Then
		__LO_ArrayFill($avFooter, $oHandout.IsFooterVisible(), $oHandout.FooterText(), $oHandout.IsPageNumberVisible())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFooter)
	EndIf

	If ($bFooter <> Null) Then
		If Not IsBool($bFooter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oHandout.IsFooterVisible = $bFooter

		$iError = ($oHandout.IsFooterVisible() = $bFooter) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sFooterText <> Null) Then
		If Not IsString($sFooterText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oHandout.FooterText = $sFooterText

		$iError = ($oHandout.FooterText() = $sFooterText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bSlideNum <> Null) Then
		If Not IsBool($bSlideNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oHandout.IsPageNumberVisible = $bSlideNum

		$iError = ($oHandout.IsPageNumberVisible() = $bSlideNum) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideHandoutFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideHandoutFormat
; Description ...: Set or Retrieve the handout page format settings.
; Syntax ........: _LOImpress_SlideHandoutFormat(ByRef $oHandout[, $iWidth = Null[, $iHeight = Null[, $iOrientation = Null]]])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LOImpress_SlideHandoutGetObj function.
;                  $iWidth              - [optional] Default is Null. The Width of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOI_PAGE_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iHeight             - [optional] Default is Null. The Height of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOI_PAGE_HEIGHT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iOrientation        - [optional] (0-1) Default is Null. The page orientation. See Constants, $LOI_PAGE_ORIENT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oHandout not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 4 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants, $LOI_PAGE_ORIENT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current slide width.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iHeight
;                  |                               4 = Error setting $iOrientation
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  When modifying the page format, the shapes etc., aren't readjusted as they are in LibreOffice UI.
; Related .......: _LO_UnitConvert, _LOImpress_SlideLayout, _LOImpress_SlideMargins, _LOImpress_SlideMasterFormat, _LOImpress_SlideNotesFormat, _LOImpress_SlideFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideHandoutFormat(ByRef $oHandout, $iWidth = Null, $iHeight = Null, $iOrientation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_Format($oHandout, $iWidth, $iHeight, $iOrientation)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_SlideHandoutFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideHandoutGetObj
; Description ...: Retrieve the Handout page Object for an Impress document.
; Syntax ........: _LOImpress_SlideHandoutGetObj(ByRef $oDoc)
; Parameters ....: $oDoc                -  A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning Handouts page Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Handouts Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: There seems to be only one handouts page per document.
; Related .......: _LOImpress_SlideNotesGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideHandoutGetObj(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oHandout

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oHandout = $oDoc.HandoutMasterPage()
	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oHandout)
EndFunc   ;==>_LOImpress_SlideHandoutGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideHandoutHeader
; Description ...: Set or Retrieve handout page header settings.
; Syntax ........: _LOImpress_SlideHandoutHeader(ByRef $oHandout[, $bHeader = Null[, $sHeaderText = Null[, $bDateTime = Null[, $bDateTimeIsFixed = Null[, $sDateTimeValue = Null[, $iDateTimeFormat = Null]]]]]])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LOImpress_SlideHandoutGetObj function.
;                  $bHeader             - [optional] Default is Null. If True, a Header entry is added to the Header of the page.
;                  $sHeaderText         - [optional] Default is Null. If $bHeader is True, the text to display in the Header of the page.
;                  $bDateTime           - [optional] Default is Null. If True, a Date or Time entry is added to the header of the page.
;                  $bDateTimeIsFixed    - [optional] Default is Null. If True, the Date or Time entry is fixed.
;                  $sDateTimeValue      - [optional] Default is Null. If $bDateTimeIsFixed is True, this is the custom date or time value to display.
;                  $iDateTimeFormat     - [optional] (4-112) Default is Null. If $bDateTimeIsFixed is False, the format to display the Date or Time in. See Constants, $LOI_SLIDE_DT_FMT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oHandout not an Object.
;                  @Error: 1, @Extended: 2 = $bHeader not a Boolean.
;                  @Error: 1, @Extended: 3 = $sHeaderText not a String.
;                  @Error: 1, @Extended: 4 = $bDateTime not a Boolean.
;                  @Error: 1, @Extended: 5 = $bDateTimeIsFixed not a Boolean.
;                  @Error: 1, @Extended: 6 = $sDateTimeValue not a String.
;                  @Error: 1, @Extended: 7 = $iDateTimeFormat not an Integer, less than 4 or greater than 9 but not equal to one of the constant values. See Constants, $LOI_SLIDE_DT_FMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bHeader
;                  |                               2 = Error setting $sHeaderText
;                  |                               4 = Error setting $bDateTime
;                  |                               8 = Error setting $bDateTimeIsFixed
;                  |                               16 = Error setting $sDateTimeValue
;                  |                               32 = Error setting $iDateTimeFormat
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When retrieving current setting values, both $sDateTimeValue and $iDateTimeFormat may return a value. To determine which is currently valid, check $bDateTimeIsFixed. If $bDateTimeIsFixed is True, $sDateTimeValue is valid, else $iDateTimeFormat. If $bDateTime is false, neither will be valid.
;                  Apply to all is not added to this function as it is not an actual setting. The user can simulate this easily by making a loop to apply it to all slides.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, _LOImpress_SlideNotesHeader
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideHandoutHeader(ByRef $oHandout, $bHeader = Null, $sHeaderText = Null, $bDateTime = Null, $bDateTimeIsFixed = Null, $sDateTimeValue = Null, $iDateTimeFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHeader[6]
	Local $sAllowed = $LOI_SLIDE_DT_FMT_24H_HM & ":" & $LOI_SLIDE_DT_FMT_MMDDYY_24H_HM & ":" & $LOI_SLIDE_DT_FMT_24H_HMS & ":" & $LOI_SLIDE_DT_FMT_12H_HM_AMPM & ":" & $LOI_SLIDE_DT_FMT_MMDDYY_12H_HM_AMPM & ":" & $LOI_SLIDE_DT_FMT_12H_HMS_AMPM

	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bHeader, $sHeaderText, $bDateTime, $bDateTimeIsFixed, $sDateTimeValue, $iDateTimeFormat) Then
		__LO_ArrayFill($avHeader, $oHandout.IsHeaderVisible(), $oHandout.HeaderText(), $oHandout.IsDateTimeVisible(), $oHandout.IsDateTimeFixed(), _
				$oHandout.DateTimeText(), $oHandout.DateTimeFormat())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avHeader)
	EndIf

	If ($bHeader <> Null) Then
		If Not IsBool($bHeader) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oHandout.IsHeaderVisible = $bHeader

		$iError = ($oHandout.IsHeaderVisible() = $bHeader) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sHeaderText <> Null) Then
		If Not IsString($sHeaderText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oHandout.HeaderText = $sHeaderText

		$iError = ($oHandout.HeaderText() = $sHeaderText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bDateTime <> Null) Then
		If Not IsBool($bDateTime) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oHandout.IsDateTimeVisible = $bDateTime

		$iError = ($oHandout.IsDateTimeVisible() = $bDateTime) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bDateTimeIsFixed <> Null) Then
		If Not IsBool($bDateTimeIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oHandout.IsDateTimeFixed = $bDateTimeIsFixed

		$iError = ($oHandout.IsDateTimeFixed() = $bDateTimeIsFixed) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sDateTimeValue <> Null) Then
		If Not IsString($sDateTimeValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oHandout.DateTimeText = $sDateTimeValue

		$iError = ($oHandout.DateTimeText() = $sDateTimeValue) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iDateTimeFormat <> Null) Then
		If Not __LO_IntIsBetween($iDateTimeFormat, $LOI_SLIDE_DT_FMT_MMDDYY, $LOI_SLIDE_DT_FMT_DOW_MMMM_DD_YYYY, "", $sAllowed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oHandout.DateTimeFormat = $iDateTimeFormat

		$iError = ($oHandout.DateTimeFormat() = $iDateTimeFormat) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideHandoutHeader

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideHandoutLayout
; Description ...: Set or Retrieve the current Handout page's layout.
; Syntax ........: _LOImpress_SlideHandoutLayout(ByRef $oHandout[, $iLayout = Null])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LOImpress_SlideHandoutGetObj function.
;                  $iLayout             - [optional] (22-31) Default is Null. The layout format of the Handout page. See Constants, $LOI_HANDOUT_LAYOUT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current layout setting as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oHandout not an Object.
;                  @Error: 1, @Extended: 2 = $iLayout not an Integer, less than 22 or greater than 31. See Constants, $LOI_HANDOUT_LAYOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Page's current layout.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLayout
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOImpress_SlideHandoutFormat, _LOImpress_SlideHandoutMargins, _LOImpress_SlideLayout
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideHandoutLayout(ByRef $oHandout, $iLayout = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iCurrLayout

	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLayout) Then
		$iCurrLayout = $oHandout.Layout()
		If Not IsInt($iCurrLayout) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurrLayout)
	EndIf

	If Not __LO_IntIsBetween($iLayout, $LOI_HANDOUT_LAYOUT_ONE_SLIDE, $LOI_HANDOUT_LAYOUT_NINE_SLIDES) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oHandout.Layout = $iLayout
	$iError = ($oHandout.Layout() = $iLayout) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideHandoutLayout

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideHandoutMargins
; Description ...: Set or Retrieve the handout page margin settings.
; Syntax ........: _LOImpress_SlideHandoutMargins(ByRef $oHandout[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LOImpress_SlideHandoutGetObj function.
;                  $iLeft               - [optional] Default is Null. The amount of space to leave between the left edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] Default is Null. The amount of space to leave between the right edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] Default is Null. The amount of space to leave between the upper edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] Default is Null. The amount of space to leave between the lower edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oHandout not an Object.
;                  @Error: 1, @Extended: 2 = $iLeft not an Integer.
;                  @Error: 1, @Extended: 3 = $iRight not an Integer.
;                  @Error: 1, @Extended: 4 = $iTop not an Integer.
;                  @Error: 1, @Extended: 5 = $iBottom not an Integer.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOImpress_SlideLayout, _LOImpress_SlideFormat, _LOImpress_SlideMasterMargins, _LOImpress_SlideNotesMargins, _LOImpress_SlideMargins
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideHandoutMargins(ByRef $oHandout, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_Margins($oHandout, $iLeft, $iRight, $iTop, $iBottom)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_SlideHandoutMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideLayout
; Description ...: Set or Retrieve the current Slide's layout.
; Syntax ........: _LOImpress_SlideLayout(ByRef $oSlide[, $iLayout = Null])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iLayout             - [optional] (0-34) Default is Null. The layout format of the Slide. See Constants, $LOI_SLIDE_LAYOUT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current layout setting as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iLayout not an Integer, less than 0 or greater than 34. See Constants, $LOI_SLIDE_LAYOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Slide's current layout.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLayout
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOImpress_SlideName, _LOImpress_SlideTransition, _LOImpress_SlideFormat, _LOImpress_SlideMargins, _LOImpress_SlideHandoutLayout
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideLayout(ByRef $oSlide, $iLayout = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iCurrLayout

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLayout) Then
		$iCurrLayout = $oSlide.Layout()
		If Not IsInt($iCurrLayout) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurrLayout)
	EndIf

	If Not __LO_IntIsBetween($iLayout, $LOI_SLIDE_LAYOUT_TITLE, $LOI_SLIDE_LAYOUT_TITLE_6_CONTENT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSlide.Layout = $iLayout
	$iError = ($oSlide.Layout() = $iLayout) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideLayout

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMargins
; Description ...: Set or Retrieve the slide page margin settings.
; Syntax ........: _LOImpress_SlideMargins(ByRef $oSlide[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iLeft               - [optional] Default is Null. The amount of space to leave between the left edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] Default is Null. The amount of space to leave between the right edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] Default is Null. The amount of space to leave between the upper edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] Default is Null. The amount of space to leave between the lower edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iLeft not an Integer.
;                  @Error: 1, @Extended: 3 = $iRight not an Integer.
;                  @Error: 1, @Extended: 4 = $iTop not an Integer.
;                  @Error: 1, @Extended: 5 = $iBottom not an Integer.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOImpress_SlideLayout, _LOImpress_SlideFormat, _LOImpress_SlideHandoutMargins, _LOImpress_SlideMasterMargins, _LOImpress_SlideNotesMargins
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMargins(ByRef $oSlide, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_Margins($oSlide, $iLeft, $iRight, $iTop, $iBottom)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_SlideMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterAdd
; Description ...: Add a master slide to a presentation.
; Syntax ........: _LOImpress_SlideMasterAdd(ByRef $oDoc[, $iPos = Null[, $sName = ""[, $bBlank = True]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iPos                - [optional] Default is Null. The position to insert the new master slide in the collection of slides. 0 Based. This is ignored if $bBlank is False.
;                  $sName               - [optional] Default is "". The unique name of the Master Slide. If called with an empty string, LibreOffice automatically names it.
;                  $bBlank              - [optional] Default is True. If True, the new Master Slide is blank. If False a Master Slide with a preset layout is inserted. See remarks.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning new slide's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPos not an Integer, less than 0 or greater than number of master slides.
;                  @Error: 1, @Extended: 3 = $sName not a String.
;                  @Error: 1, @Extended: 4 = Name called in $sName already exists.
;                  @Error: 1, @Extended: 5 = $bBlank not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to create a master slide.
;                  @Error: 3, @Extended: 2 = Failed to retrieve master pages Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve count of master pages.
;                  @Error: 3, @Extended: 4 = Failed to retrieve master page Object
;                  @Error: 3, @Extended: 5 = Failed to identify new master slide Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $iPos is called with Null, the new master slide is inserted at the end.
;                  Call $iPos with the last master slide index to insert the master slide at the end. Call $iPos with 0 to insert the new master slide at the beginning.
;                  When inserting a non-blank master slide ($bBlank called with false), the new slide will be inserted AFTER the last slide, $iPos is ignored.
;                  This function uses two methods to insert a Master slide. Using the API, the resulting new master slide is blank, without text boxes etc., the second method uses a document dispatch command, which results in a normally formatted master slide, like when you add a master slide manually.
;                  When inserting a new slide with $bBlank set to False, I use the dispatch command to accomplish the insertion, this method seems to only ever insert the new slide at the end of all the slides.
;                  I have not found a way to import Master slide from the LibreOffice templates yet.
; Related .......: _LOImpress_SlideMasterDeleteByIndex, _LOImpress_SlideMasterDeleteByObj, _LOImpress_SlideAdd, _LOImpress_SlideMasterExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterAdd(ByRef $oDoc, $iPos = Null, $sName = "", $bBlank = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMSlide, $oServiceManager, $oDispatcher, $oMasters, $oMaster
	Local $aoMasters[0]
	Local $iMasters
	Local $aArray[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($iPos = Null) Then $iPos = ($bBlank) ? ($oDoc.MasterPages.getCount()) : ($oDoc.MasterPages.getCount() - 1) ; If I am inserting a Master using the dispatch, I have make position be 1 less than the count so I can retrieve the Object for the last master slide.
	If ($iPos = $oDoc.MasterPages.getCount()) Then $iPos = $iPos - 1 ; If I am inserting a Master using the dispatch command, and the user called the last slide position plus 1, I need to change it to be 1 less so I can retrieve the Object for the last master slide.
	If Not __LO_IntIsBetween($iPos, 0, $oDoc.MasterPages.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sName <> "") And _LOImpress_SlideMasterExists($oDoc, $sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bBlank) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If $bBlank Then
		$oMSlide = $oDoc.MasterPages.insertNewByIndex($iPos)
		If Not IsObj($oMSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else
		; When inserting a new Master using a Dispatch, I have to backup a copy of all current master slide objects so I can identify the new slide.
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
		If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oMasters = $oDoc.MasterPages()
		If Not IsObj($oMasters) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$iMasters = $oMasters.getCount()
		If Not IsInt($iMasters) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		ReDim $aoMasters[$iMasters]

		For $i = 0 To $iMasters - 1
			$aoMasters[$i] = $oMasters.getByIndex($i)
			If Not IsObj($aoMasters[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:InsertMasterPage", "", 0, $aArray)

		; Identify the new Master Slide.
		For $i = 0 To $oMasters.getCount() - 1
			$oMaster = $oMasters.getByIndex($i)
			If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			For $j = 0 To $iMasters - 1
				; If the Object is a match, exit this loop and continue the top-level loop, bypassing the Objext assignment.
				If $aoMasters[$j] = $oMaster Then ContinueLoop 2

				Sleep((IsInt($j / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
			Next

			$oMSlide = $oMaster
		Next

		If Not IsObj($oMSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	If ($sName <> "") Then $oMSlide.Name = $sName

	Return SetError($__LO_STATUS_SUCCESS, 0, $oMSlide)
EndFunc   ;==>_LOImpress_SlideMasterAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterBackColor
; Description ...: Set or Retrieve the Master Slide's background color.
; Syntax ........: _LOImpress_SlideMasterBackColor(ByRef $oMaster[, $iColor = Null])
; Parameters ....: $oMaster             - A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Master Slide background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Integer
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting as an Integer value. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current color value.
;                  @Error: 3, @Extended: 2 = Failed to retrieve parent Document.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColor
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the slide, the Constant $LO_COLOR_OFF is returned.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_SlideMasterBackFillStyle, _LOImpress_SlideMasterBackGradient, _LOImpress_SlideBackColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterBackColor(ByRef $oMaster, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBackground, $oDoc
	Local $iError = 0, $iCurColor

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oMaster.Background()

	If __LO_VarsAreNull($iColor) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, $LO_COLOR_OFF) ; If no background is set, this will be void, instead of an Object.

		$iCurColor = __LOImpress_ColorRemoveAlpha($oBackground.FillColor())
		If Not IsInt($iCurColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurColor)
	EndIf

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oDoc = __LOImpress_GetParentDoc($oMaster)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oMaster.Background = $oBackground
	EndIf

	$oBackground.FillStyle = $LOI_AREA_FILL_STYLE_SOLID
	$oBackground.FillColor = $iColor
	$iError = ($oMaster.Background.FillColor() = $iColor) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideMasterBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterBackFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOImpress_SlideMasterBackFillStyle(ByRef $oMaster[, $bFillOff = False])
; Parameters ....: $oMaster             - A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
;                  $bFillOff            - [optional] Default is False. If True, the Fill style will be set to Off. See remarks.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOI_AREA_FILL_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 0, @Extended: 1, Return: 0 = Success. Fill style was successfully turned off.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $bFillOff not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Fill Style.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
;                  When the Fill style is disabled for a Master Slide, the Fill properties are completely removed. This is how Impress works normally.
;                  $bFillOff will do nothing if it is called with False, and is not, of course, returned when retrieving the FillStyle value.
; Related .......: _LOImpress_SlideMasterBackColor, _LOImpress_SlideMasterBackGradient, _LOImpress_SlideBackFillStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterBackFillStyle(ByRef $oMaster, $bFillOff = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle
	Local $oBackground

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bFillOff) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bFillOff Then
		If IsObj($oMaster.Background()) Then
			$oBackground = $oMaster.Background
			If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, 0) ; If no Background Object, no Fillstyle is active.

			$oBackground.FillStyle = $LOI_AREA_FILL_STYLE_OFF
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, 0)
	EndIf

	$oBackground = $oMaster.Background()
	If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_AREA_FILL_STYLE_OFF) ; If no Background Object, no Fillstyle is active.

	$iFillStyle = $oBackground.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOImpress_SlideMasterBackFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterBackGradient
; Description ...: Modify or retrieve the settings for Master Slide Background color Gradient.
; Syntax ........: _LOImpress_SlideMasterBackGradient(ByRef $oMaster[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oMaster             - A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
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
;                  @Error: 0, @Extended: 2, Return: -1 = Success. All optional parameters were called with Null, no background is currently active for the master slide. Returning -1.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
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
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillGradient" Struct.
;                  @Error: 3, @Extended: 2 = Error retrieving Parent Document.
;                  @Error: 3, @Extended: 3 = Error retrieving Color Stop Array for "From" color
;                  @Error: 3, @Extended: 4 = Error retrieving Color Stop Array for "To" color
;                  @Error: 3, @Extended: 5 = Error creating Gradient Name.
;                  @Error: 3, @Extended: 6 = Error setting Gradient Name.
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
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_SlideMasterBackColor, _LOImpress_SlideMasterBackFillStyle, _LOImpress_SlideBackGradient
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterBackGradient(ByRef $oMaster, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBackground, $oDoc
	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $nRed, $nGreen, $nBlue
	Local $atColorStop
	Local $avGradient[11]
	Local $sGradName

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oMaster.Background()

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

	$oDoc = __LOImpress_GetParentDoc($oMaster)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oMaster.Background = $oBackground
	EndIf

	$tStyleGradient = $oBackground.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oBackground.FillStyle() <> $LOI_AREA_FILL_STYLE_GRADIENT) Then $oBackground.FillStyle = $LOI_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		__LOImpress_GradientPresets($oDoc, $oBackground, $tStyleGradient, $sGradientName)

		$tStyleGradient = $oBackground.FillGradient()
		If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$iError = ($oBackground.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oBackground.FillStyle = $LOI_AREA_FILL_STYLE_OFF
			$oBackground.FillGradientName = ""

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
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

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
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

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

	If ($oBackground.FillGradientName() = "") Or __LOImpress_GradientIsModified($tStyleGradient, $oBackground.FillGradientName()) Then
		$sGradName = __LOImpress_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		$oBackground.FillGradientName = $sGradName
		If ($oBackground.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
	EndIf

	$oBackground.FillGradient = $tStyleGradient

	; Error checking
	$iError = (__LO_VarsAreNull($iType)) ? $iError : ($oMaster.Background.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2))
	$iError = (__LO_VarsAreNull($iXCenter)) ? $iError : ($oMaster.Background.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8))
	$iError = (__LO_VarsAreNull($iYCenter)) ? $iError : ($oMaster.Background.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16))
	$iError = (__LO_VarsAreNull($iAngle)) ? $iError : (($oMaster.Background.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? $iError : ($oMaster.Background.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64))
	$iError = (__LO_VarsAreNull($iFromColor)) ? $iError : ($oMaster.Background.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128))
	$iError = (__LO_VarsAreNull($iToColor)) ? $iError : ($oMaster.Background.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256))
	$iError = (__LO_VarsAreNull($iFromIntense)) ? $iError : ($oMaster.Background.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512))
	$iError = (__LO_VarsAreNull($iToIntense)) ? $iError : ($oMaster.Background.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideMasterBackGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterBackTransparency
; Description ...: Set or retrieve Transparency settings for a Master Slide.
; Syntax ........: _LOImpress_SlideMasterBackTransparency(ByRef $oMaster[, $iTransparency = Null])
; Parameters ....: $oMaster             - A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
;                  $iTransparency       - [optional] (0-100) Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Transparency value.
;                  @Error: 3, @Extended: 2 = Failed to retrieve parent Document.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the Master slide, -1 is returned.
; Related .......: _LOImpress_SlideMasterBackTransparencyGradient, _LOImpress_SlideBackTransparency
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterBackTransparency(ByRef $oMaster, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurTransp
	Local $oBackground, $oDoc

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oMaster.Background()

	If __LO_VarsAreNull($iTransparency) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, -1) ; No background present.

		$iCurTransp = $oBackground.FillTransparence()
		If Not IsInt($iCurTransp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurTransp)
	EndIf

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oDoc = __LOImpress_GetParentDoc($oMaster)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oMaster.Background = $oBackground
	EndIf

	$oBackground.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oBackground.FillTransparence = $iTransparency

	$iError = ($oMaster.Background.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideMasterBackTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterBackTransparencyGradient
; Description ...: Set or retrieve the Master Slide's transparency gradient settings.
; Syntax ........: _LOImpress_SlideMasterBackTransparencyGradient(ByRef $oMaster[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oMaster             - A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
;                  $iType               - [optional] (-1-5) Default is Null. The type of transparency gradient to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3. Call with $LOI_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error: 0, @Extended: 1, Return: -1 = Success. All optional parameters were called with Null no background is currently active for the master slide. Returning -1.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $iType Not an Integer, less than -1 or greater than 5. See constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iXCenter Not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 4 = $iYCenter Not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 5 = $iAngle Not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 6 = $iTransitionStart Not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iStart Not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iEnd Not an Integer, less than 0 or greater than 100.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.drawing.Background" service.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillTransparenceGradient" Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve parent Document.
;                  @Error: 3, @Extended: 3 = Error retrieving Color Stop Array for "From" color
;                  @Error: 3, @Extended: 4 = Error retrieving Color Stop Array for "To" color
;                  @Error: 3, @Extended: 5 = Error creating Transparency Gradient name.
;                  @Error: 3, @Extended: 6 = Error setting Transparency Gradient name.
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
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  While these properties can be set successfully, LibreOffice doesn't seem to apply it to the master slide, even when done using the UI.
; Related .......: _LOImpress_SlideMasterBackTransparency, _LOImpress_SlideBackTransparencyGradient
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterBackTransparencyGradient(ByRef $oMaster, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop
	Local $oBackground, $oDoc
	Local $fValue

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oMaster.Background()

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 2, -1)

		$tGradient = $oBackground.FillTransparenceGradient()
		If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LO_ArrayFill($aiTransparent, $tGradient.Style(), $tGradient.XOffset(), $tGradient.YOffset(), _
				($tGradient.Angle() / 10), $tGradient.Border(), __LOImpress_TransparencyGradientConvert(Null, $tGradient.StartColor()), _
				__LOImpress_TransparencyGradientConvert(Null, $tGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	$oDoc = __LOImpress_GetParentDoc($oMaster)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oMaster.Background = $oBackground
	EndIf

	$tGradient = $oBackground.FillTransparenceGradient()
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iType <> Null) Then
		If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oBackground.FillTransparenceGradientName = ""

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

	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oMaster.Background.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oMaster.Background.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oMaster.Background.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((($oMaster.Background.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oMaster.Background.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iStart)) ? ($iError) : (($oMaster.Background.FillTransparenceGradient.StartColor() = __LOImpress_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iEnd)) ? ($iError) : (($oMaster.Background.FillTransparenceGradient.EndColor() = __LOImpress_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideMasterBackTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterCurrent
; Description ...: Set or Retrieve the currently applied Master slide to a slide.
; Syntax ........: _LOImpress_SlideMasterCurrent(ByRef $oSlide[, $oMaster = Null])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $oMaster             - [optional] Default is Null. A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
; Return values .: Success: 1 or Object.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Object = Success. All optional parameters were called with Null, returning currently applied Master Slide as an Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $oMaster not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve currently applied Master slide.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oMaster
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOImpress_SlideMasterGetObjByIndex, _LOImpress_SlideMasterGetObjByName, _LOImpress_SlideCurrent, _LOImpress_SlideCurrent
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterCurrent(ByRef $oSlide, $oMaster = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCurrMaster
	Local $iError

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($oMaster) Then
		$oCurrMaster = $oSlide.MasterPage()
		If Not IsObj($oCurrMaster) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $oCurrMaster)
	EndIf

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSlide.MasterPage = $oMaster
	$iError = ($oSlide.MasterPage() = $oMaster) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideMasterCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterDeleteByIndex
; Description ...: Delete a master slide by index.
; Syntax ........: _LOImpress_SlideMasterDeleteByIndex(ByRef $oDoc, $iMaster)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iMaster             - The index of the master slide to delete. 0 based.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Master slide was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iMaster not an Integer, less than 0 or greater than number of Master slides minus one.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve count of master slides.
;                  @Error: 3, @Extended: 2 = Failed to retrieve master slide's Object.
;                  @Error: 3, @Extended: 3 = Failed to delete master slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Trying to delete a Master Slide that is used by a slide will result in a processing error. I currently have no way of checking if a master slide is free to be deleted.
; Related .......: _LOImpress_SlideMasterDeleteByObj, _LOImpress_SlideMastersGetCount, _LOImpress_SlideDeleteByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterDeleteByIndex(ByRef $oDoc, $iMaster)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMSlide
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iMaster, 0, $oDoc.MasterPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iCount = $oDoc.MasterPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oMSlide = $oDoc.MasterPages.getByIndex($iMaster)
	If Not IsObj($oMSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.MasterPages.remove($oMSlide)
	If ($iCount = $oDoc.MasterPages.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to delete because the count is the same.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideMasterDeleteByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterDeleteByObj
; Description ...: Delete a master slide using its Object.
; Syntax ........: _LOImpress_SlideMasterDeleteByObj(ByRef $oMaster)
; Parameters ....: $oMaster             -  A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Slide was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Parent Document.
;                  @Error: 3, @Extended: 2 = Failed to retrieve count of master slides.
;                  @Error: 3, @Extended: 3 = Failed to delete master slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Trying to delete a Master Slide that is used by a slide will result in a processing error. I currently have no way of checking if a master slide is free to be deleted.
; Related .......: _LOImpress_SlideMasterDeleteByIndex, _LOImpress_SlideMasterGetObjByIndex, _LOImpress_SlideMasterGetObjByName, _LOImpress_SlideDeleteByObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterDeleteByObj(ByRef $oMaster)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc
	Local $iCount

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = __LOImpress_GetParentDoc($oMaster)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oDoc.MasterPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.MasterPages.Remove($oMaster)
	If ($oDoc.MasterPages.getCount() = $iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to delete because the count is the same.

	$oMaster = Null

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideMasterDeleteByObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterExists
; Description ...: Check whether a master slide with a certain name exists in a document.
; Syntax ........: _LOImpress_SlideMasterExists(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sName               - The master slide name to check for.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Document contains a master Slide with the called name, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query for master Slide name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterName, _LOImpress_SlideMastersGetNames, _LOImpress_SlideExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterExists(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bExists

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$bExists = $oDoc.Links.getByName("Master Page").Links.hasByName($sName)
	If Not IsBool($bExists) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bExists)
EndFunc   ;==>_LOImpress_SlideMasterExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterFormat
; Description ...: Set or Retrieve the master slide format settings.
; Syntax ........: _LOImpress_SlideMasterFormat(ByRef $oMaster[, $iWidth = Null[, $iHeight = Null[, $iOrientation = Null]]])
; Parameters ....: $oMaster             - A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
;                  $iWidth              - [optional] Default is Null. The Width of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOI_PAGE_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iHeight             - [optional] Default is Null. The Height of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOI_PAGE_HEIGHT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iOrientation        - [optional] (0-1) Default is Null. The page orientation. See Constants, $LOI_PAGE_ORIENT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 4 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants, $LOI_PAGE_ORIENT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current slide width.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iHeight
;                  |                               4 = Error setting $iOrientation
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  When modifying the page format, the shapes etc., aren't readjusted as they are in LibreOffice UI.
;                  I am unable to find the properties to set for "FitObject to Paper Format", "Background covers margins", "Slide numbers", and "Paper tray".
; Related .......: _LO_UnitConvert, _LOImpress_SlideMasterMargins, _LOImpress_SlideHandoutFormat, _LOImpress_SlideNotesFormat, _LOImpress_SlideFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterFormat(ByRef $oMaster, $iWidth = Null, $iHeight = Null, $iOrientation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_Format($oMaster, $iWidth, $iHeight, $iOrientation)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_SlideMasterFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterGetObjByIndex
; Description ...: Retrieve a Master Slide's Object by index.
; Syntax ........: _LOImpress_SlideMasterGetObjByIndex(ByRef $oDoc, $iMaster)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iMaster             - The index of the master slide to retrieve. 0 based.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested master slide's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iMaster not an Integer, less than 0 or greater than number of master slides minus one.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested master slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideMasterGetObjByName, _LOImpress_SlideMastersGetCount, _LOImpress_SlideGetObjByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterGetObjByIndex(ByRef $oDoc, $iMaster)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMSlide

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iMaster, 0, $oDoc.MasterPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oMSlide = $oDoc.MasterPages.getByIndex($iMaster)
	If Not IsObj($oMSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oMSlide)
EndFunc   ;==>_LOImpress_SlideMasterGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterGetObjByName
; Description ...: Retrieve a Master Slide's Object by name.
; Syntax ........: _LOImpress_SlideMasterGetObjByName(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sName               - The Master Slide's name to retrieve the Object for.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested Master Slide's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Master Slide name called in $sName not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested Master Slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I have not found a way to import Master slide from the LibreOffice templates yet.
; Related .......: _LOImpress_SlideMasterGetObjByIndex, _LOImpress_SlideMastersGetNames, _LOImpress_SlideGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterGetObjByName(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMSlide

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.Links.getByName("Master Page").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oMSlide = $oDoc.Links.getByName("Master Page").Links.getByName($sName)
	If Not IsObj($oMSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oMSlide)
EndFunc   ;==>_LOImpress_SlideMasterGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterMargins
; Description ...: Set or Retrieve the master slide page margin settings.
; Syntax ........: _LOImpress_SlideMasterMargins(ByRef $oMaster[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oMaster             - A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
;                  $iLeft               - [optional] Default is Null. The amount of space to leave between the left edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] Default is Null. The amount of space to leave between the right edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] Default is Null. The amount of space to leave between the upper edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] Default is Null. The amount of space to leave between the lower edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $iLeft not an Integer.
;                  @Error: 1, @Extended: 3 = $iRight not an Integer.
;                  @Error: 1, @Extended: 4 = $iTop not an Integer.
;                  @Error: 1, @Extended: 5 = $iBottom not an Integer.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOImpress_SlideMasterFormat, _LOImpress_SlideHandoutMargins, _LOImpress_SlideNotesMargins, _LOImpress_SlideMargins
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterMargins(ByRef $oMaster, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_Margins($oMaster, $iLeft, $iRight, $iTop, $iBottom)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_SlideMasterMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterName
; Description ...: Set or Retrieve a Master Slide's name.
; Syntax ........: _LOImpress_SlideMasterName(ByRef $oMaster[, $sName = Null])
; Parameters ....: $oMaster             - A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
;                  $sName               - [optional] Default is Null. The new name to set the Master slide to. See Remarks.
; Return values .: Success: 1 or String.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. All optional parameters were called with Null, returning current Master Slide name as a String.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Master Slide name called in $sName already exists in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Master Slide name.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Document Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If setting the Master slide name to a name and a number, there is a good chance the name won't stay applied, as LibreOffice will assume it is an auto-numbered slide.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOImpress_SlideMasterExists, _LOImpress_SlideMastersGetNames, _LOImpress_SlideName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterName(ByRef $oMaster, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurrName
	Local $oDoc

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sName) Then
		$sCurrName = $oMaster.LinkDisplayName()
		If Not IsString($sCurrName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrName)
	EndIf

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc = __LOImpress_GetParentDoc($oMaster)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If $oDoc.Links.getByName("Master Page").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oMaster.Name = $sName
	$iError = ($oMaster.LinkDisplayName() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideMasterName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMasterNotesGetObj
; Description ...: Retrieve the Notes Object for a Master Slide.
; Syntax ........: _LOImpress_SlideMasterNotesGetObj(ByRef $oMaster)
; Parameters ....: $oMaster             - A Master Slide object returned by a previous _LOImpress_SlideMasterAdd, _LOImpress_SlideMasterGetObjByIndex, or _LOImpress_SlideMasterGetObjByName function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning Notes page Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Notes Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideNotesGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMasterNotesGetObj(ByRef $oMaster)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNotes

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oNotes = $oMaster.NotesPage()
	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNotes)
EndFunc   ;==>_LOImpress_SlideMasterNotesGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMastersGetCount
; Description ...: Retrieve a count of master slides.
; Syntax ........: _LOImpress_SlideMastersGetCount(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning count of master slides contained in the document.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve a count of master slides.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This only returns a count of master slides already loaded into the document.
; Related .......: _LOImpress_SlideMasterDeleteByIndex, _LOImpress_SlideMasterGetObjByIndex, _LOImpress_SlidesGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMastersGetCount(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oDoc.MasterPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOImpress_SlideMastersGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMastersGetNames
; Description ...: Retrieve an array of names for all Master Slides contained in the document.
; Syntax ........: _LOImpress_SlideMastersGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array containing all Master Slide names. @Extended is set to the number of slide names returned.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Master Slides Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve count of Master Slides.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Master Slide name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This only returns a list of master slide names already loaded into the document.
;                  I have not found a way to import Master slide from the LibreOffice templates yet.
; Related .......: _LOImpress_SlideMasterExists, _LOImpress_SlideMasterGetObjByName, _LOImpress_SlidesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMastersGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asMasters[0]
	Local $oMasters
	Local $iMasters = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oMasters = $oDoc.MasterPages()
	If Not IsObj($oMasters) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iMasters = $oMasters.getCount()
	If Not IsInt($iMasters) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	ReDim $asMasters[$iMasters]

	For $i = 0 To $iMasters - 1
		$asMasters[$i] = $oMasters.getByIndex($i).Name()
		If Not IsString($asMasters[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iMasters, $asMasters)
EndFunc   ;==>_LOImpress_SlideMastersGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideMove
; Description ...: Move a slide in the collection of slides.
; Syntax ........: _LOImpress_SlideMove(ByRef $oSlide, $iPos)
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iPos                - The position to move the slide to in the collection of slides. 0 Based. See remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Slide was successfully moved.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iPos not an Integer, less than 0 or greater than number of slides minus 1.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Parent Document.
;                  @Error: 3, @Extended: 2 = Failed to identify slide's current position.
;                  @Error: 3, @Extended: 3 = Failed to move slide.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Due to limitations in the API, some dispatches are executed to move the slide. The current slide will temporarily be set to the new slide in order to move it.
; Related .......: _LOImpress_SlideCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideMove(ByRef $oSlide, $iPos)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc, $oCurrSlide, $oServiceManager, $oDispatcher
	Local $iCurrPos, $iMove, $iCurrView
	Local $sDispatch
	Local $aArray[0]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = __LOImpress_GetParentDoc($oSlide)
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

	$iCurrView = __LOImpress_DocCurrView($oDoc)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oDoc.getCurrentController.setCurrentPage($oSlide)

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

	If IsObj($oCurrSlide) Then $oDoc.getCurrentController.setCurrentPage($oCurrSlide)     ; Restore current slide and view mode.

	If IsInt($iCurrView) Then __LOImpress_DocCurrView($oDoc, $iCurrView)

	If ($iCurrPos <> $iPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideName
; Description ...: Set or Retrieve a Slide's name.
; Syntax ........: _LOImpress_SlideName(ByRef $oSlide[, $sName = Null])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $sName               - [optional] Default is Null. The new name to set the slide to. See Remarks.
; Return values .: Success: 1 or String.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. All optional parameters were called with Null, returning current Slide name as a String.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Slide name called in $sName already exists in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Slide name.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Document Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If setting the slide name to a name and a number, there is a good chance the name won't stay applied, as LibreOffice will assume it is an auto-numbered slide.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOImpress_SlideExists, _LOImpress_SlidesGetNames, _LOImpress_SlideMasterName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideName(ByRef $oSlide, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurrName
	Local $oDoc

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sName) Then
		$sCurrName = $oSlide.LinkDisplayName()
		If Not IsString($sCurrName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrName)
	EndIf

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc = __LOImpress_GetParentDoc($oSlide)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If $oDoc.Links.getByName("Slide").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSlide.Name = $sName
	$iError = ($oSlide.LinkDisplayName() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideFooter
; Description ...: Set or Retrieve notes page Footer settings.
; Syntax ........: _LOImpress_SlideFooter(ByRef $oNotes[, $bFooter = Null[, $sFooterText = Null[, $bSlideNum = Null]]])
; Parameters ....: $oNotes              - A Notes page object returned by a previous _LOImpress_SlideNotesGetObj or _LOImpress_SlideMasterNotesGetObj function.
;                  $bFooter             - [optional] Default is Null. If True, a Footer entry is added to the footer of the page.
;                  $sFooterText         - [optional] Default is Null. If $bFooter is True, the text to display in the footer of the page.
;                  $bSlideNum           - [optional] Default is Null. If True, a current Slide number is added to the footer of the page.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oNotes not an Object.
;                  @Error: 1, @Extended: 2 = Object passed in $oNotes is a Master Notes Object.
;                  @Error: 1, @Extended: 3 = $bFooter not a Boolean.
;                  @Error: 1, @Extended: 4 = $sFooterText not a String.
;                  @Error: 1, @Extended: 5 = $bSlideNum not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFooter
;                  |                               2 = Error setting $sFooterText
;                  |                               4 = Error setting $bSlideNum
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Apply to all is not added to this function as they it is not an actual setting. The user can simulate this easily by making a loop to apply it to all slides.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  You can only set or retrieve footer property values for a slide notes page, not a master notes page.
; Related .......: _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, _LOImpress_SlideFooter, _LOImpress_SlideHandoutFooter
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideNotesFooter(ByRef $oNotes, $bFooter = Null, $sFooterText = Null, $bSlideNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFooter[3]

	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oNotes.supportsService("com.sun.star.drawing.MasterPage") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bFooter, $sFooterText, $bSlideNum) Then
		__LO_ArrayFill($avFooter, $oNotes.IsFooterVisible(), $oNotes.FooterText(), $oNotes.IsPageNumberVisible())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFooter)
	EndIf

	If ($bFooter <> Null) Then
		If Not IsBool($bFooter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oNotes.IsFooterVisible = $bFooter

		$iError = ($oNotes.IsFooterVisible() = $bFooter) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sFooterText <> Null) Then
		If Not IsString($sFooterText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oNotes.FooterText = $sFooterText

		$iError = ($oNotes.FooterText() = $sFooterText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bSlideNum <> Null) Then
		If Not IsBool($bSlideNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oNotes.IsPageNumberVisible = $bSlideNum

		$iError = ($oNotes.IsPageNumberVisible() = $bSlideNum) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideNotesFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideNotesFormat
; Description ...: Set or Retrieve the notes page format settings.
; Syntax ........: _LOImpress_SlideNotesFormat(ByRef $oNotes[, $iWidth = Null[, $iHeight = Null[, $iOrientation = Null]]])
; Parameters ....: $oNotes              - A Notes page object returned by a previous _LOImpress_SlideNotesGetObj or _LOImpress_SlideMasterNotesGetObj function.
;                  $iWidth              - [optional] Default is Null. The Width of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOI_PAGE_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iHeight             - [optional] Default is Null. The Height of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOI_PAGE_HEIGHT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iOrientation        - [optional] (0-1) Default is Null. The page orientation. See Constants, $LOI_PAGE_ORIENT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oNotes not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 4 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants, $LOI_PAGE_ORIENT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current slide width.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iHeight
;                  |                               4 = Error setting $iOrientation
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  When modifying the page format, the shapes etc., aren't readjusted as they are in LibreOffice UI.
; Related .......: _LO_UnitConvert, _LOImpress_SlideLayout, _LOImpress_SlideMargins, _LOImpress_SlideHandoutFormat, _LOImpress_SlideMasterFormat, _LOImpress_SlideFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideNotesFormat(ByRef $oNotes, $iWidth = Null, $iHeight = Null, $iOrientation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_Format($oNotes, $iWidth, $iHeight, $iOrientation)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_SlideNotesFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideNotesGetObj
; Description ...: Retrieve the Notes Object for a Slide.
; Syntax ........: _LOImpress_SlideNotesGetObj(ByRef $oSlide)
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning Notes page Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Notes Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideMasterNotesGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideNotesGetObj(ByRef $oSlide)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNotes

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oNotes = $oSlide.NotesPage()
	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNotes)
EndFunc   ;==>_LOImpress_SlideNotesGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideNotesHeader
; Description ...: Set or Retrieve notes page header settings.
; Syntax ........: _LOImpress_SlideNotesHeader(ByRef $oNotes[, $bHeader = Null[, $sHeaderText = Null[, $bDateTime = Null[, $bDateTimeIsFixed = Null[, $sDateTimeValue = Null[, $iDateTimeFormat = Null]]]]]])
; Parameters ....: $oNotes              - A Notes page object returned by a previous _LOImpress_SlideNotesGetObj function.
;                  $bHeader             - [optional] Default is Null. If True, a Header entry is added to the Header of the page.
;                  $sHeaderText         - [optional] Default is Null. If $bHeader is True, the text to display in the Header of the page.
;                  $bDateTime           - [optional] Default is Null. If True, a Date or Time entry is added to the header of the page.
;                  $bDateTimeIsFixed    - [optional] Default is Null. If True, the Date or Time entry is fixed.
;                  $sDateTimeValue      - [optional] Default is Null. If $bDateTimeIsFixed is True, this is the custom date or time value to display.
;                  $iDateTimeFormat     - [optional] (4-112) Default is Null. If $bDateTimeIsFixed is False, the format to display the Date or Time in. See Constants, $LOI_SLIDE_DT_FMT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oNotes not an Object.
;                  @Error: 1, @Extended: 2 = Object passed in $oNotes is a Master Notes Object.
;                  @Error: 1, @Extended: 3 = $bHeader not a Boolean.
;                  @Error: 1, @Extended: 4 = $sHeaderText not a String.
;                  @Error: 1, @Extended: 5 = $bDateTime not a Boolean.
;                  @Error: 1, @Extended: 6 = $bDateTimeIsFixed not a Boolean.
;                  @Error: 1, @Extended: 7 = $sDateTimeValue not a String.
;                  @Error: 1, @Extended: 8 = $iDateTimeFormat not an Integer, less than 4 or greater than 9 but not equal to one of the constant values. See Constants, $LOI_SLIDE_DT_FMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bHeader
;                  |                               2 = Error setting $sHeaderText
;                  |                               4 = Error setting $bDateTime
;                  |                               8 = Error setting $bDateTimeIsFixed
;                  |                               16 = Error setting $sDateTimeValue
;                  |                               32 = Error setting $iDateTimeFormat
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When retrieving current setting values, both $sDateTimeValue and $iDateTimeFormat may return a value. To determine which is currently valid, check $bDateTimeIsFixed. If $bDateTimeIsFixed is True, $sDateTimeValue is valid, else $iDateTimeFormat. If $bDateTime is false, neither will be valid.
;                  Apply to all is not added to this function as it is not an actual setting. The user can simulate this easily by making a loop to apply it to all slides.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  You can only set or retrieve header property values for a slide notes page, not a master notes page.
; Related .......: _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, _LOImpress_SlideHandoutHeader
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideNotesHeader(ByRef $oNotes, $bHeader = Null, $sHeaderText = Null, $bDateTime = Null, $bDateTimeIsFixed = Null, $sDateTimeValue = Null, $iDateTimeFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHeader[7]
	Local $sAllowed = $LOI_SLIDE_DT_FMT_24H_HM & ":" & $LOI_SLIDE_DT_FMT_MMDDYY_24H_HM & ":" & $LOI_SLIDE_DT_FMT_24H_HMS & ":" & $LOI_SLIDE_DT_FMT_12H_HM_AMPM & ":" & $LOI_SLIDE_DT_FMT_MMDDYY_12H_HM_AMPM & ":" & $LOI_SLIDE_DT_FMT_12H_HMS_AMPM

	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oNotes.supportsService("com.sun.star.drawing.MasterPage") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bHeader, $sHeaderText, $bDateTime, $bDateTimeIsFixed, $sDateTimeValue, $iDateTimeFormat) Then
		__LO_ArrayFill($avHeader, $oNotes.IsHeaderVisible(), $oNotes.HeaderText(), $oNotes.IsDateTimeVisible(), $oNotes.IsDateTimeFixed(), $oNotes.DateTimeText(), _
				$oNotes.DateTimeFormat())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avHeader)
	EndIf

	If ($bHeader <> Null) Then
		If Not IsBool($bHeader) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oNotes.IsHeaderVisible = $bHeader

		$iError = ($oNotes.IsHeaderVisible() = $bHeader) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sHeaderText <> Null) Then
		If Not IsString($sHeaderText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oNotes.HeaderText = $sHeaderText

		$iError = ($oNotes.HeaderText() = $sHeaderText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bDateTime <> Null) Then
		If Not IsBool($bDateTime) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oNotes.IsDateTimeVisible = $bDateTime

		$iError = ($oNotes.IsDateTimeVisible() = $bDateTime) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bDateTimeIsFixed <> Null) Then
		If Not IsBool($bDateTimeIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oNotes.IsDateTimeFixed = $bDateTimeIsFixed

		$iError = ($oNotes.IsDateTimeFixed() = $bDateTimeIsFixed) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sDateTimeValue <> Null) Then
		If Not IsString($sDateTimeValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oNotes.DateTimeText = $sDateTimeValue

		$iError = ($oNotes.DateTimeText() = $sDateTimeValue) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iDateTimeFormat <> Null) Then
		If Not __LO_IntIsBetween($iDateTimeFormat, $LOI_SLIDE_DT_FMT_MMDDYY, $LOI_SLIDE_DT_FMT_DOW_MMMM_DD_YYYY, "", $sAllowed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oNotes.DateTimeFormat = $iDateTimeFormat

		$iError = ($oNotes.DateTimeFormat() = $iDateTimeFormat) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideNotesHeader

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideNotesMargins
; Description ...: Set or Retrieve the notes page margin settings.
; Syntax ........: _LOImpress_SlideNotesMargins(ByRef $oNotes[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oNotes              - A Notes page object returned by a previous _LOImpress_SlideNotesGetObj or _LOImpress_SlideMasterNotesGetObj function.
;                  $iLeft               - [optional] Default is Null. The amount of space to leave between the left edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] Default is Null. The amount of space to leave between the right edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] Default is Null. The amount of space to leave between the upper edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] Default is Null. The amount of space to leave between the lower edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oNotes not an Object.
;                  @Error: 1, @Extended: 2 = $iLeft not an Integer.
;                  @Error: 1, @Extended: 3 = $iRight not an Integer.
;                  @Error: 1, @Extended: 4 = $iTop not an Integer.
;                  @Error: 1, @Extended: 5 = $iBottom not an Integer.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOImpress_SlideLayout, _LOImpress_SlideFormat, _LOImpress_SlideHandoutMargins, _LOImpress_SlideMasterMargins, _LOImpress_SlideMargins
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideNotesMargins(ByRef $oNotes, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_Margins($oNotes, $iLeft, $iRight, $iTop, $iBottom)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_SlideNotesMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlidesGetCount
; Description ...: Retrieve a count of slides.
; Syntax ........: _LOImpress_SlidesGetCount(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning count of slides contained in the document.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve a count of slides.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideDeleteByIndex, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideMastersGetCount
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
; Name ..........: _LOImpress_SlidesGetNames
; Description ...: Retrieve an array of names for all Slides contained in the document.
; Syntax ........: _LOImpress_SlidesGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array containing all Slide names. @Extended is set to the number of slide names returned.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve array of Slide names.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideExists, _LOImpress_SlideGetObjByName, _LOImpress_SlideMastersGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlidesGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asSlides[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asSlides = $oDoc.Links.getByName("Slide").Links.getElementNames()
	If Not IsArray($asSlides) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asSlides), $asSlides)
EndFunc   ;==>_LOImpress_SlidesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowActiveSettings
; Description ...: Set or Retrieve settings for an actively running presentation.
; Syntax ........: _LOImpress_SlideshowActiveSettings(ByRef $oDoc[, $bKeepOnTop = Null[, $bMouseVisible = Null[, $bMouseAsPen = Null[, $iPenColor = Null[, $iPenWidth = Null]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bKeepOnTop          - [optional] Default is Null. If True, the presentation will be always kept on top of other programs.
;                  $bMouseVisible       - [optional] Default is Null. If True, the mouse is visible in the presentation.
;                  $bMouseAsPen         - [optional] Default is Null. If True, the mouse can be used as a pen to draw on slides.
;                  $iPenColor           - [optional] (0-16777215) Default is Null. If $bMouseAsPen is True, the color of the drawn line, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iPenWidth           - [optional] (4-400) Default is Null. The width of the drawn line. L.O. 4.2+. See Constants, $LOI_SLIDESHOW_PEN_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters. If The current LibreOffice version is below 4.2, the $iPenWidth parameter will return a Null value.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bKeepOnTop not a Boolean.
;                  @Error: 1, @Extended: 3 = $bMouseVisible not a Boolean.
;                  @Error: 1, @Extended: 4 = $bMouseAsPen not a Boolean.
;                  @Error: 1, @Extended: 5 = $iPenColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 6 = $iPenWidth not an Integer, less than 4 or greater than 400. See Constants, $LOI_SLIDESHOW_PEN_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = There is no presentation currently running.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Object for currently running presentation.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bKeepOnTop
;                  |                               2 = Error setting $bMouseVisible
;                  |                               4 = Error setting $bMouseAsPen
;                  |                               8 = Error setting $iPenColor
;                  |                               16 = Error setting $iPenWidth
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version less than 4.2, $iPenWidth not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_SlideshowIsRunning, _LOImpress_SlideshowSettingsMode, _LOImpress_SlideshowSettingsOptions, _LOImpress_SlideshowSettingsRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowActiveSettings(ByRef $oDoc, $bKeepOnTop = Null, $bMouseVisible = Null, $bMouseAsPen = Null, $iPenColor = Null, $iPenWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oPresentation
	Local $avSlideShow[5]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.Presentation.isRunning() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; No Slideshow active.

	$oPresentation = $oDoc.Presentation.getController()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($bKeepOnTop, $bMouseVisible, $bMouseAsPen, $iPenColor, $iPenWidth) Then
		If __LO_VersionCheck(4.2) Then
			__LO_ArrayFill($avSlideShow, $oPresentation.AlwaysOnTop(), $oPresentation.MouseVisible(), $oPresentation.UsePen(), $oPresentation.PenColor(), $oPresentation.PenWidth())

		Else
			__LO_ArrayFill($avSlideShow, $oPresentation.AlwaysOnTop(), $oPresentation.MouseVisible(), $oPresentation.UsePen(), $oPresentation.PenColor(), Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSlideShow)
	EndIf

	If ($bKeepOnTop <> Null) Then
		If Not IsBool($bKeepOnTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oPresentation.AlwaysOnTop = $bKeepOnTop

		$iError = ($oPresentation.AlwaysOnTop() = $bKeepOnTop) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bMouseVisible <> Null) Then
		If Not IsBool($bMouseVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPresentation.MouseVisible = $bMouseVisible

		$iError = ($oPresentation.MouseVisible() = $bMouseVisible) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bMouseAsPen <> Null) Then
		If Not IsBool($bMouseAsPen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPresentation.UsePen = $bMouseAsPen

		$iError = ($oPresentation.UsePen() = $bMouseAsPen) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iPenColor <> Null) Then
		If Not __LO_IntIsBetween($iPenColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPresentation.PenColor = $iPenColor

		$iError = ($oPresentation.PenColor() = $iPenColor) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iPenWidth <> Null) Then
		If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
		If Not __LO_IntIsBetween($iPenWidth, $LOI_SLIDESHOW_PEN_WIDTH_VERY_THIN, $LOI_SLIDESHOW_PEN_WIDTH_VERY_THICK) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPresentation.PenWidth = $iPenWidth

		$iError = ($oPresentation.PenWidth() = $iPenWidth) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideshowActiveSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowCustomCreate
; Description ...: Create a Custom Slideshow.
; Syntax ........: _LOImpress_SlideshowCustomCreate(ByRef $oDoc, $sName, $asSlides)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sName               - The name of the Custom Slideshow to create.
;                  $asSlides            - A single column Array of Slide names. See remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully created new Custom Slideshow.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Name called in $sName already exists as a Custom Slideshow in Document.
;                  @Error: 1, @Extended: 4 = $asSlides not an Array.
;                  @Error: 1, @Extended: 5 = Array called in $asSlides has 0 elements.
;                  @Error: 1, @Extended: 6 = Element contained in $asSlides not a String. Returning problem element number.
;                  @Error: 1, @Extended: 7 = Slide name contained in $asSlides not found in Document. Returning problem element number.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a CustomPresentation Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve the Links Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Slide's Object.
;                  @Error: 3, @Extended: 3 = Failed to insert new Custom Slideshow.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The expected input for $asSlides is a single column array having the Slide names in the order the user wishes the Slides to appear in the presentation, slide names can be placed in the Array multiple times.
; Related .......: _LOImpress_SlideshowCustomDelete, _LOImpress_SlideshowsCustomGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowCustomCreate(ByRef $oDoc, $sName, $asSlides)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oLinks, $oCustomPres, $oSlide

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsArray($asSlides) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If (UBound($asSlides) < 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oLinks = $oDoc.Links.getByName("Slide").Links()
	If Not IsObj($oLinks) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($asSlides) - 1
		If Not IsString($asSlides[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)
		If Not $oLinks.hasByName($asSlides[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $i)
	Next

	$oCustomPres = $oDoc.CustomPresentations.createInstance()
	If Not IsObj($oCustomPres) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($asSlides) - 1
		$oSlide = $oLinks.getByName($asSlides[$i])
		If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oCustomPres.insertByIndex($oCustomPres.getCount(), $oSlide)
	Next

	$oDoc.CustomPresentations.insertByName($sName, $oCustomPres)
	If Not $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideshowCustomCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowCustomDelete
; Description ...: Deletes a Custom Slideshow.
; Syntax ........: _LOImpress_SlideshowCustomDelete(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sName               - The Custom Slideshow's name to delete.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Custom Slideshow was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Name called in $sName not found as a Custom Slideshow in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to delete the requested Custom Slideshow.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideshowCustomCreate, _LOImpress_SlideshowsCustomGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowCustomDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oDoc.CustomPresentations.removeByName($sName)
	If $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideshowCustomDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowCustomModify
; Description ...: Set or Retrieve the Slides and order of the slides contained in a Custom Slideshow.
; Syntax ........: _LOImpress_SlideshowCustomModify(ByRef $oDoc, $sName[, $asSlides = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sName               - The name of the Custom Slideshow to modify.
;                  $asSlides            - [optional] Default is Null. A single column Array of Slide names. See remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Custom Slideshow successfully modified.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning Array of Slide names contained in the Custom Slideshow. See remarks.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Name called in $sName already exists as a Custom Slideshow in Document.
;                  @Error: 1, @Extended: 4 = $asSlides not an Array.
;                  @Error: 1, @Extended: 5 = Array called in $asSlides has 0 elements.
;                  @Error: 1, @Extended: 6 = Element contained in $asSlides not a String. Returning problem element number.
;                  @Error: 1, @Extended: 7 = Slide name contained in $asSlides not found in Document. Returning problem element number.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a CustomPresentation Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Custom Slideshow's Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Slide's name.
;                  @Error: 3, @Extended: 3 = Failed to retrieve the Links Object.
;                  @Error: 3, @Extended: 4 = Failed to retrieve Slide's Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $asSlides
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The expected input for $asSlides is a single column array having the Slide names in the order the user wishes the Slides to appear in the presentation, slide names can be placed in the Array multiple times.
;                  When retrieving the current order and content of the Slideshow, an array is returned with all the Slides contained in the Custom Slideshow, in the order they are set to be played. Slides may be present multiple times.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOImpress_SlideshowCustomCreate, _LOImpress_SlideshowsCustomGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowCustomModify(ByRef $oDoc, $sName, $asSlides = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oLinks, $oCustomPres, $oNewCustomPres, $oSlide
	Local $asCurrSlides[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($asSlides) Then
		$oCustomPres = $oDoc.CustomPresentations.getByName($sName)
		If Not IsObj($oCustomPres) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		ReDim $asCurrSlides[$oCustomPres.getCount()]
		For $i = 0 To $oCustomPres.getCount() - 1
			$asCurrSlides[$i] = $oCustomPres.getByIndex($i).LinkDisplayName()
			If Not IsString($asCurrSlides[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $asCurrSlides)
	EndIf

	If Not IsArray($asSlides) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If (UBound($asSlides) < 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oLinks = $oDoc.Links.getByName("Slide").Links()
	If Not IsObj($oLinks) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asSlides) - 1
		If Not IsString($asSlides[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)
		If Not $oLinks.hasByName($asSlides[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $i)
	Next

	$oNewCustomPres = $oDoc.CustomPresentations.createInstance()
	If Not IsObj($oNewCustomPres) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($asSlides) - 1
		$oSlide = $oLinks.getByName($asSlides[$i])
		If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oNewCustomPres.insertByIndex($oNewCustomPres.getCount(), $oSlide)
	Next

	$oDoc.CustomPresentations.replaceByName($sName, $oNewCustomPres)
	$iError = ($oDoc.CustomPresentations.getByName($sName).getCount() = UBound($asSlides)) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideshowCustomModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowCustomSetName
; Description ...: Rename a Custom Slideshow.
; Syntax ........: _LOImpress_SlideshowCustomSetName(ByRef $oDoc, $sName, $sNewName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sName               - The name of the Custom Slideshow to rename.
;                  $sNewName            - The name to rename the Custom Slideshow to.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Custom Slideshow was successfully renamed.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Name called in $sName not found as a Custom Slideshow in Document.
;                  @Error: 1, @Extended: 4 = $sNewName not a String.
;                  @Error: 1, @Extended: 5 = Name called in $sNewName already exists as a Custom Slideshow in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested Custom Slideshow Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sNewName
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideshowCustomModify, _LOImpress_SlideshowsCustomGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowCustomSetName(ByRef $oDoc, $sName, $sNewName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oCustomPres

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $oDoc.CustomPresentations.hasByName($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oCustomPres = $oDoc.CustomPresentations.getByName($sName)
	If Not IsObj($oCustomPres) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oCustomPres.setName($sNewName)
	$iError = ($oCustomPres.Name() = $sNewName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideshowCustomSetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowIsRunning
; Description ...: Check whether there is a presentation currently running.
; Syntax ........: _LOImpress_SlideshowIsRunning(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if there is currently a Presentation running, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to determine if a Presentation is currently active.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideshowStart, _LOImpress_SlideshowStop
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowIsRunning(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsRunning

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsRunning = $oDoc.Presentation.IsRunning()
	If Not IsBool($bIsRunning) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsRunning)
EndFunc   ;==>_LOImpress_SlideshowIsRunning

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowPresentationControl
; Description ...: Query the status of, or send commands to, a currently running presentation.
; Syntax ........: _LOImpress_SlideshowPresentationControl(ByRef $oDoc, $iAction[, $vValue = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iAction             - The Query or Command to perform on the presentation. See Constants, $LOI_SLIDESHOW_PRES_* as defined in LibreOfficeImpress_Constants.au3.
;                  $vValue              - [optional] Default is Null. If the Query or Command requires an input value, it goes here. See Remarks.
; Return values .: Success: Boolean, Integer, or Object.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully processed a command.
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Successfully processed a query that returns a Boolean. (See description of the specific query to see what is returned.)
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Successfully processed a query that returns an Integer. (See description of the specific query to see what is returned.)
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully processed a query that returns an Object. (See description of the specific query to see what is returned.)
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iAction not an Integer, less than 0 or greater than 25. See Constants, $LOI_SLIDESHOW_PRES_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iAction called with $LOI_SLIDESHOW_PRES_QUERY_GET_SLIDE_BY_INDEX, and index value called in $vValue is not an Integer, less than 0 or greater than number of slides in the Presentation.
;                  @Error: 1, @Extended: 4 = $iAction called with $LOI_SLIDESHOW_PRES_COMMAND_ACTIVATE_BLANK_SCREEN, and color value called in $vValue is not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iAction called with $LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE, and value called in $vValue is not an Object.
;                  @Error: 1, @Extended: 6 = $iAction called with $LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE_BY_INDEX, and index value called in $vValue is not an Integer, less than 0 or greater than number of slides in the Presentation.
;                  @Error: 1, @Extended: 7 = $iAction called with $LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE_BY_NAME, and value called in $vValue is not a String.
;                  @Error: 1, @Extended: 8 = $iAction called with $LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE_BY_NAME, and Slide name called in $vValue does not exist.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = There is no presentation currently running.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Object for currently running presentation.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Object for current slide.
;                  @Error: 3, @Extended: 4 = Failed to retrieve current slide's index value.
;                  @Error: 3, @Extended: 5 = Failed to retrieve next slide's index value.
;                  @Error: 3, @Extended: 6 = Failed to retrieve Object for requested slide.
;                  @Error: 3, @Extended: 7 = Failed to retrieve count of slides in the presentation.
;                  @Error: 3, @Extended: 8 = Failed to determine if the presentation is Active.
;                  @Error: 3, @Extended: 9 = Failed to determine if the presentation is Endless.
;                  @Error: 3, @Extended: 10 = Failed to determine if the presentation is FullScreen.
;                  @Error: 3, @Extended: 11 = Failed to determine if the presentation is Paused.
;                  @Error: 3, @Extended: 12 = Failed to activate presentation.
;                  @Error: 3, @Extended: 13 = Failed to activate blank screen for presentation.
;                  @Error: 3, @Extended: 14 = Failed to move to first slide.
;                  @Error: 3, @Extended: 15 = Failed to move to last slide.
;                  @Error: 3, @Extended: 16 = Failed to move to requested slide by Object.
;                  @Error: 3, @Extended: 17 = Failed to move to requested slide by Index.
;                  @Error: 3, @Extended: 18 = Failed to move to requested slide by Name.
;                  @Error: 3, @Extended: 19 = Failed to Pause the presentation.
;                  @Error: 3, @Extended: 20 = Failed to Resume the presentation.
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = $iAction called with $LOI_SLIDESHOW_PRES_COMMAND_ERASE_ALL_INK, and current LibreOffice version is less than 7.2.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Any queries or commands that require an input parameter will have the type of input required indicated in the description for the Constant.
; Related .......: _LOImpress_SlideshowActiveSettings, _LOImpress_SlideshowIsRunning
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowPresentationControl(ByRef $oDoc, $iAction, $vValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn = 1
	Local $oPresentation

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iAction, $LOI_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE, $LOI_SLIDESHOW_PRES_COMMAND_STOP_SOUND) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.Presentation.isRunning() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; No Slideshow active.

	$oPresentation = $oDoc.Presentation.getController()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Switch $iAction
		Case $LOI_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE
			$vReturn = $oPresentation.getCurrentSlide()
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOI_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE_INDEX
			$vReturn = $oPresentation.getCurrentSlideIndex()
			If Not IsInt($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Case $LOI_SLIDESHOW_PRES_QUERY_GET_NEXT_SLIDE_INDEX
			$vReturn = $oPresentation.getNextSlideIndex()
			If Not IsInt($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		Case $LOI_SLIDESHOW_PRES_QUERY_GET_SLIDE_BY_INDEX
			If Not __LO_IntIsBetween($vValue, 0, $oPresentation.getSlideCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$vReturn = $oPresentation.getSlideByIndex($vValue)
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

		Case $LOI_SLIDESHOW_PRES_QUERY_GET_SLIDE_COUNT
			$vReturn = $oPresentation.getSlideCount()
			If Not IsInt($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

		Case $LOI_SLIDESHOW_PRES_QUERY_IS_ACTIVE
			$vReturn = $oPresentation.isActive()
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

		Case $LOI_SLIDESHOW_PRES_QUERY_IS_ENDLESS
			$vReturn = $oPresentation.isEndless()
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 9, 0)

		Case $LOI_SLIDESHOW_PRES_QUERY_IS_FULLSCREEN
			$vReturn = $oPresentation.isFullScreen()
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 10, 0)

		Case $LOI_SLIDESHOW_PRES_QUERY_IS_PAUSED
			$vReturn = $oPresentation.isPaused()
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 11, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_ACTIVATE
			$oPresentation.activate()
			If Not $oPresentation.isActive() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 12, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_ACTIVATE_BLANK_SCREEN
			If Not __LO_IntIsBetween($vValue, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oPresentation.blankScreen($vValue)
			If Not $oPresentation.isPaused() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 13, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_DEACTIVATE
			$oPresentation.deactivate() ; Doesn't seem to set IsActive to False!
;~ 			If $oPresentation.isActive() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 14, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_ERASE_ALL_INK
			If Not __LO_VersionCheck(7.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

			$oPresentation.setEraseAllInk(True)

		Case $LOI_SLIDESHOW_PRES_COMMAND_GOTO_FIRST_SLIDE
			$oPresentation.gotoFirstSlide()
			If ($oPresentation.getCurrentSlideIndex() <> 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 14, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_GOTO_LAST_SLIDE
			$oPresentation.gotoLastSlide()
			If ($oPresentation.getCurrentSlideIndex() <> $oPresentation.getSlideCount() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 15, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_GOTO_NEXT_EFFECT
			$oPresentation.gotoNextEffect()

		Case $LOI_SLIDESHOW_PRES_COMMAND_GOTO_NEXT_SLIDE
			$oPresentation.gotoNextSlide()

		Case $LOI_SLIDESHOW_PRES_COMMAND_GOTO_PREV_EFFECT
			$oPresentation.gotoPreviousEffect()

		Case $LOI_SLIDESHOW_PRES_COMMAND_GOTO_PREV_SLIDE
			$oPresentation.gotoPreviousSlide()

		Case $LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE
			If Not IsObj($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oPresentation.gotoSlide($vValue)
			If ($oPresentation.getCurrentSlide() <> $vValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 16, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE_BY_INDEX
			If Not __LO_IntIsBetween($vValue, 0, $oPresentation.getSlideCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oPresentation.gotoSlideIndex($vValue)
			If ($oPresentation.getCurrentSlideIndex() <> $vValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 17, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE_BY_NAME
			If Not IsString($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
			If Not $oDoc.Links.getByName("Slide").Links.hasByName($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oPresentation.gotoBookmark($vValue)
			If ($oPresentation.getCurrentSlide.LinkDisplayName() <> $vValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 18, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_PAUSE
			$oPresentation.pause()
			If Not $oPresentation.isPaused() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 19, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_RESUME
			$oPresentation.resume()
			If $oPresentation.isPaused() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 20, 0)

		Case $LOI_SLIDESHOW_PRES_COMMAND_STOP_SOUND
			$oPresentation.stopSound()
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, $vReturn)
EndFunc   ;==>_LOImpress_SlideshowPresentationControl

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowsCustomGetNames
; Description ...: Retrieve an array of Custom Slideshow names available in the document.
; Syntax ........: _LOImpress_SlideshowsCustomGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array containing all Custom Slideshow names. @Extended is set to the number of names returned.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Custom Presentations Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideshowCustomCreate, _LOImpress_SlideshowCustomDelete, _LOImpress_SlideshowCustomModify, _LOImpress_SlideshowCustomSetName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowsCustomGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asCustomSlideShows[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asCustomSlideShows = $oDoc.CustomPresentations.ElementNames()
	If Not IsArray($asCustomSlideShows) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asCustomSlideShows), $asCustomSlideShows)
EndFunc   ;==>_LOImpress_SlideshowsCustomGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowSettingsMode
; Description ...: Set or Retrieve the Slideshow's play mode settings.
; Syntax ........: _LOImpress_SlideshowSettingsMode(ByRef $oDoc[, $iPresMode = Null[, $iRepeatPause = Null[, $bShowLogo = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iPresMode           - [optional] (0-2) Default is Null. The mode the presentation is displayed in. See Constants, $LOI_SLIDESHOW_VIEW_MODE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRepeatPause        - [optional] (0-86399) Default is Null. If $iPresMode is set to $LOI_SLIDESHOW_VIEW_MODE_LOOP, the amount of seconds before the presentation is played again.
;                  $bShowLogo           - [optional] Default is Null. If True, the LibreOffice logo is displayed during the pause.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPresMode not an Integer, less than 0 or greater than 2. See Constants, $LOI_SLIDESHOW_VIEW_MODE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iRepeatPause not an Integer, less than 0 or greater than 86399.
;                  @Error: 1, @Extended: 4 = $bShowLogo not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Presentation Object.
;                  @Error: 3, @Extended: 2 = Failed to identify current mode.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iPresMode
;                  |                               2 = Error setting $iRepeatPause
;                  |                               4 = Error setting $bShowLogo
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_SlideshowActiveSettings, _LOImpress_SlideshowPresentationControl, _LOImpress_SlideshowSettingsOptions, _LOImpress_SlideshowSettingsRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowSettingsMode(ByRef $oDoc, $iPresMode = Null, $iRepeatPause = Null, $bShowLogo = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurrMode
	Local $oPresentation
	Local $avSlideShow[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPresentation = $oDoc.Presentation()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iPresMode, $iRepeatPause, $bShowLogo) Then
		If ($oPresentation.IsFullScreen() = True) And ($oPresentation.IsEndless() = False) Then
			$iCurrMode = $LOI_SLIDESHOW_VIEW_MODE_FULL_SCREEN

		ElseIf ($oPresentation.IsFullScreen() = False) And ($oPresentation.IsEndless() = False) Then
			$iCurrMode = $LOI_SLIDESHOW_VIEW_MODE_IN_WINDOW

		ElseIf ($oPresentation.IsEndless() = True) Then
			$iCurrMode = $LOI_SLIDESHOW_VIEW_MODE_LOOP

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Failed to identify current mode.
		EndIf

		__LO_ArrayFill($avSlideShow, $iCurrMode, $oPresentation.Pause(), $oPresentation.IsShowLogo())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSlideShow)
	EndIf

	If ($iPresMode <> Null) Then
		If Not __LO_IntIsBetween($iPresMode, $LOI_SLIDESHOW_VIEW_MODE_FULL_SCREEN, $LOI_SLIDESHOW_VIEW_MODE_LOOP) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		Switch $iPresMode
			Case $LOI_SLIDESHOW_VIEW_MODE_FULL_SCREEN
				$oPresentation.IsFullScreen = True
				$oPresentation.IsEndless = False
				$iError = ($oPresentation.IsFullScreen() = True) ? ($iError) : (BitOR($iError, 1))
				$iError = ($oPresentation.IsEndless = False) ? ($iError) : (BitOR($iError, 1))

			Case $LOI_SLIDESHOW_VIEW_MODE_IN_WINDOW
				$oPresentation.IsFullScreen = False
				$oPresentation.IsEndless = False
				$iError = ($oPresentation.IsFullScreen() = False) ? ($iError) : (BitOR($iError, 1))
				$iError = ($oPresentation.IsEndless = False) ? ($iError) : (BitOR($iError, 1))

			Case $LOI_SLIDESHOW_VIEW_MODE_LOOP
				$oPresentation.IsEndless = True
				$iError = ($oPresentation.IsEndless = True) ? ($iError) : (BitOR($iError, 1))
		EndSwitch
	EndIf

	If ($iRepeatPause <> Null) Then
		If Not __LO_IntIsBetween($iRepeatPause, 0, 86399) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; 0 seconds to 23:59:59 hours.

		$oPresentation.Pause = $iRepeatPause
		$iError = ($oPresentation.Pause() = $iRepeatPause) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bShowLogo <> Null) Then
		If Not IsBool($bShowLogo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPresentation.IsShowLogo = $bShowLogo
		$iError = ($oPresentation.IsShowLogo() = $bShowLogo) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideshowSettingsMode

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowSettingsOptions
; Description ...: Set or Retrieve the Slideshow's play options settings.
; Syntax ........: _LOImpress_SlideshowSettingsOptions(ByRef $oDoc[, $bDisableAutoSlides = Null[, $bChangeSlideByClick = Null[, $bMouseVisible = Null[, $bMouseAsPen = Null[, $bPlayAnimatedFiles = Null[, $bKeepOnTop = Null]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bDisableAutoSlides  - [optional] Default is Null. If True, slides will not transition to the next slide automatically (overriding individual slide settings).
;                  $bChangeSlideByClick - [optional] Default is Null. If True, slides will transition when the mouse is clicked.
;                  $bMouseVisible       - [optional] Default is Null. If True, the mouse is visible in the presentation.
;                  $bMouseAsPen         - [optional] Default is Null. If True, the mouse can be used as a pen to draw on slides.
;                  $bPlayAnimatedFiles  - [optional] Default is Null. If True, animated files (such as GIFs) will be played.
;                  $bKeepOnTop          - [optional] Default is Null. If True, the presentation will be always kept on top of other programs.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bDisableAutoSlides not a Boolean.
;                  @Error: 1, @Extended: 3 = $bChangeSlideByClick not a Boolean.
;                  @Error: 1, @Extended: 4 = $bMouseVisible not a Boolean.
;                  @Error: 1, @Extended: 5 = $bMouseAsPen not a Boolean.
;                  @Error: 1, @Extended: 6 = $bPlayAnimatedFiles not a Boolean.
;                  @Error: 1, @Extended: 7 = $bKeepOnTop not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Presentation Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bDisableAutoSlides
;                  |                               2 = Error setting $bChangeSlideByClick
;                  |                               4 = Error setting $bMouseVisible
;                  |                               8 = Error setting $bMouseAsPen
;                  |                               16 = Error setting $bPlayAnimatedFiles
;                  |                               32 = Error setting $bKeepOnTop
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_SlideshowActiveSettings, _LOImpress_SlideshowPresentationControl, _LOImpress_SlideshowSettingsMode, _LOImpress_SlideshowSettingsRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowSettingsOptions(ByRef $oDoc, $bDisableAutoSlides = Null, $bChangeSlideByClick = Null, $bMouseVisible = Null, $bMouseAsPen = Null, $bPlayAnimatedFiles = Null, $bKeepOnTop = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oPresentation
	Local $avSlideShow[6]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPresentation = $oDoc.Presentation()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bDisableAutoSlides, $bChangeSlideByClick, $bMouseVisible, $bMouseAsPen, $bPlayAnimatedFiles, $bKeepOnTop) Then
		__LO_ArrayFill($avSlideShow, $oPresentation.IsAutomatic(), $oPresentation.IsTransitionOnClick(), $oPresentation.IsMouseVisible(), _
				$oPresentation.UsePen(), $oPresentation.AllowAnimations(), $oPresentation.IsAlwaysOnTop())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSlideShow)
	EndIf

	If ($bDisableAutoSlides <> Null) Then
		If Not IsBool($bDisableAutoSlides) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oPresentation.IsAutomatic = $bDisableAutoSlides
		$iError = ($oPresentation.IsAutomatic() = $bDisableAutoSlides) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bChangeSlideByClick <> Null) Then
		If Not IsBool($bChangeSlideByClick) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPresentation.IsTransitionOnClick = $bChangeSlideByClick
		$iError = ($oPresentation.IsTransitionOnClick() = $bChangeSlideByClick) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bMouseVisible <> Null) Then
		If Not IsBool($bMouseVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPresentation.IsMouseVisible = $bMouseVisible
		$iError = ($oPresentation.IsMouseVisible() = $bMouseVisible) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bMouseAsPen <> Null) Then
		If Not IsBool($bMouseAsPen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPresentation.UsePen = $bMouseAsPen
		$iError = ($oPresentation.UsePen() = $bMouseAsPen) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bPlayAnimatedFiles <> Null) Then
		If Not IsBool($bPlayAnimatedFiles) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPresentation.AllowAnimations = $bPlayAnimatedFiles
		$iError = ($oPresentation.AllowAnimations() = $bPlayAnimatedFiles) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bKeepOnTop <> Null) Then
		If Not IsBool($bKeepOnTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPresentation.IsAlwaysOnTop = $bKeepOnTop
		$iError = ($oPresentation.IsAlwaysOnTop() = $bKeepOnTop) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideshowSettingsOptions

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowSettingsRange
; Description ...: Set or Retrieve the Slideshow's play Range settings.
; Syntax ........: _LOImpress_SlideshowSettingsRange(ByRef $oDoc[, $iRange = Null[, $sValue = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iRange              - [optional] (0-2) Default is Null. The Range of slides that will be shown when the Presentation is started. See Constants, $LOI_SLIDESHOW_RANGE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $sValue              - [optional] Default is Null. The "From" slide or Custom Slide Show name. See remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iRange not an Integer, less than 0 or greater than 2. See Constants, $LOI_SLIDESHOW_RANGE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $sValue not a String.
;                  @Error: 1, @Extended: 4 = Range set to $LOI_SLIDESHOW_RANGE_FROM, and the Slide name called in $sValue does not exist.
;                  @Error: 1, @Extended: 5 = Range set to $LOI_SLIDESHOW_RANGE_CUSTOM, and the Custom Slideshow name called in $sValue does not exist.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Presentation Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Start From slide value.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Custom Slideshow name.
;                  @Error: 3, @Extended: 4 = Failed to identify current range.
;                  @Error: 3, @Extended: 5 = $iRange is called with other than $LOI_SLIDESHOW_RANGE_ALL, and $sValue is not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRange
;                  |                               2 = Error setting $sValue
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you call $iRange with any other value than $LOI_SLIDESHOW_RANGE_ALL, $sValue must be called with an appropriate name, either a Slide name to start from, or a Custom Slideshow name.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  If there are two slides with the same name, and one is set to the "From Slide" property, there is no guarantee which slide will be the one used.
; Related .......: _LOImpress_SlideshowActiveSettings, _LOImpress_SlideshowPresentationControl, _LOImpress_SlideshowSettingsMode, _LOImpress_SlideshowSettingsOptions
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowSettingsRange(ByRef $oDoc, $iRange = Null, $sValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurrRange
	Local $oPresentation
	Local $avSlideShow[2]
	Local $sCurrValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPresentation = $oDoc.Presentation()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oPresentation.IsShowAll() = True) Then
		$iCurrRange = $LOI_SLIDESHOW_RANGE_ALL
		$sCurrValue = ""

	ElseIf ($oPresentation.FirstPage() <> "") Then
		$iCurrRange = $LOI_SLIDESHOW_RANGE_FROM
		$sCurrValue = $oPresentation.FirstPage()
		If Not IsString($sCurrValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$sCurrValue = $oDoc.DrawPages.getByName($sCurrValue).LinkDisplayName()    ; Get Link Display Name as it is more reliable?
		If Not IsString($sCurrValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	ElseIf ($oPresentation.CustomShow() <> "") Then
		$iCurrRange = $LOI_SLIDESHOW_RANGE_CUSTOM
		$sCurrValue = $oPresentation.CustomShow()
		If Not IsString($sCurrValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)     ; Failed to identify current range.
	EndIf

	If __LO_VarsAreNull($iRange, $sValue) Then
		__LO_ArrayFill($avSlideShow, $iCurrRange, $sCurrValue)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSlideShow)
	EndIf

	If ($iRange <> Null) Then
		If Not __LO_IntIsBetween($iRange, $LOI_SLIDESHOW_RANGE_ALL, $LOI_SLIDESHOW_RANGE_CUSTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		Switch $iRange
			Case $LOI_SLIDESHOW_RANGE_ALL
				$oPresentation.IsShowAll = True
				$oPresentation.FirstPage = ""
				$oPresentation.CustomShow = ""

				$iError = ($oPresentation.IsShowAll() = True) ? ($iError) : (BitOR($iError, 1))
				$iError = ($oPresentation.FirstPage() = "") ? ($iError) : (BitOR($iError, 1))
				$iError = ($oPresentation.CustomShow() = "") ? ($iError) : (BitOR($iError, 1))

			Case $LOI_SLIDESHOW_RANGE_FROM
				If ($iCurrRange <> $iRange) Then
					If ($sValue = Null) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Value not called

					$oPresentation.IsShowAll = False
					$oPresentation.CustomShow = ""

					$iError = ($oPresentation.IsShowAll() = False) ? ($iError) : (BitOR($iError, 1))
					$iError = ($oPresentation.CustomShow() = "") ? ($iError) : (BitOR($iError, 1))
				EndIf

			Case $LOI_SLIDESHOW_RANGE_CUSTOM
				If ($iCurrRange <> $iRange) Then
					If ($sValue = Null) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Value not called

					$oPresentation.IsShowAll = False
					$oPresentation.FirstPage = ""

					$iError = ($oPresentation.IsShowAll() = False) ? ($iError) : (BitOR($iError, 1))
					$iError = ($oPresentation.FirstPage() = "") ? ($iError) : (BitOR($iError, 1))
				EndIf
		EndSwitch

		$iCurrRange = $iRange
	EndIf

	If ($sValue <> Null) Then
		If Not IsString($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		Switch $iCurrRange
			Case $LOI_SLIDESHOW_RANGE_FROM
				If Not $oDoc.Links.getByName("Slide").Links.hasByName($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

				$sValue = $oDoc.Links.getByName("Slide").Links.getByName($sValue).Name()    ; Overwrite value (LinkDisplayName) with Slide's name, as that is what L.O. uses.

				$oPresentation.FirstPage = $sValue
				$iError = ($oPresentation.FirstPage() = $sValue) ? ($iError) : (BitOR($iError, 2))

			Case $LOI_SLIDESHOW_RANGE_CUSTOM
				If Not $oDoc.CustomPresentations.hasByName($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

				$oPresentation.CustomShow = $sValue
				$iError = ($oPresentation.CustomShow() = $sValue) ? ($iError) : (BitOR($iError, 2))
		EndSwitch
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideshowSettingsRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowStart
; Description ...: Begins a presentation.
; Syntax ........: _LOImpress_SlideshowStart(ByRef $oDoc[, $bRehearse = False[, $sStartSlide = ""[, $sCustomShow = ""]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bRehearse           - [optional] Default is False. If True, starts the presentation from the beginning and shows a rehearsal timer to the user.
;                  $sStartSlide         - [optional] Default is "". The Slide's name to begin this presentation from.
;                  $sCustomShow         - [optional] Default is "". The Custom Slideshow's name to play for this presentation.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. The Presentation was started successfully.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bRehearse not a Boolean.
;                  @Error: 1, @Extended: 3 = $sStartSlide not a String.
;                  @Error: 1, @Extended: 4 = $sCustomShow not a String.
;                  @Error: 1, @Extended: 5 = Slide name called in $sStartSlide not found.
;                  @Error: 1, @Extended: 6 = Custom Slideshow name called in $sCustomShow not found.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create FirstPage property.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Presentation Object.
;                  @Error: 3, @Extended: 2 = There is already a presentation running.
;                  @Error: 3, @Extended: 3 = Failed to start the presentation.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bRehearse is called with True, both $sStartSlide and $sCustomShow will be ignored.
;                  If both $sStartSlide and $sCustomShow are called with a parameter, $sCustomShow will be ignored.
; Related .......: _LOImpress_SlideshowIsRunning, _LOImpress_SlideshowStop
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowStart(ByRef $oDoc, $bRehearse = False, $sStartSlide = "", $sCustomShow = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atProperties[0]
	Local $oPresentation
	Local $bBackupShowAll
	Local $sBackupFirst, $sBackupCustom

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bRehearse) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sStartSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sCustomShow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oPresentation = $oDoc.Presentation()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oPresentation.IsRunning() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; A Slideshow is already active.

	If $bRehearse Then
		$oPresentation.rehearseTimings()

	ElseIf ($sStartSlide <> "") Then
		If Not $oDoc.Links.getByName("Slide").Links.hasByName($sStartSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		ReDim $atProperties[1]
		$atProperties[0] = __LO_SetPropertyValue("FirstPage", $sStartSlide)
		If Not IsObj($atProperties[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oPresentation.startWithArguments($atProperties)

	ElseIf ($sCustomShow <> "") Then
		If Not $oDoc.CustomPresentations.hasByName($sCustomShow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		; This does not work (IllegalArgument COM error, I think there is some form of bug in L.O., so I use a workaround.
		; ReDim $atProperties[1]
		; $atProperties[0] = __LO_SetPropertyValue("CustomShow", $sCustomShow)
		; If Not IsObj($atProperties[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		; $oPresentation.startWithArguments($atProperties)

		$bBackupShowAll = $oPresentation.IsShowAll()
		$sBackupFirst = $oPresentation.FirstPage()
		$sBackupCustom = $oPresentation.CustomShow()

		$oPresentation.IsShowAll = False
		$oPresentation.FirstPage = ""
		$oPresentation.CustomShow = $sCustomShow

		$oPresentation.start()

		$oPresentation.IsShowAll = $bBackupShowAll
		$oPresentation.FirstPage = $sBackupFirst
		$oPresentation.CustomShow = $sBackupCustom

	Else
		$oPresentation.start()
	EndIf

	If Not $oPresentation.IsRunning() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideshowStart

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideshowStop
; Description ...: Stop the presently playing presentation.
; Syntax ........: _LOImpress_SlideshowStop(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Presentation was successfully stopped.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to stop the running presentation.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_SlideshowIsRunning, _LOImpress_SlideshowStart
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideshowStop(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If $oDoc.Presentation.IsRunning() Then
		$oDoc.Presentation.end()
		If $oDoc.Presentation.IsRunning() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Failed to stop Slideshow.
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_SlideshowStop

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideSoundsGetNames
; Description ...: Retrieve an array of Sound files that are included with LibreOffice Impress.
; Syntax ........: _LOImpress_SlideSoundsGetNames()
; Parameters ....: None
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning array of included Impress Sound files. @Extended will be set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create the ServiceManager.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve the DefaultContext Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve the "/singletons/com.sun.star.util.thePathSettings" Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve the LibreOffice gallery path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: An example path that may be returned is: "C:\Program Files\LibreOffice\program\..\share\gallery\curve.wav"
; Related .......: _LOImpress_SlideTransition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideSoundsGetNames()
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oContext, $oPathSettings
	Local $asGalleryPath[0], $asFiles[35]
	Local $iCount = 0
	Local $hSearch
	Local $sFile

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oContext = $oServiceManager.DefaultContext()
	If Not IsObj($oContext) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oPathSettings = $oContext.getValueByName("/singletons/com.sun.star.util.thePathSettings")
	If Not IsObj($oPathSettings) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	; "file:///C:/Program%20Files/LibreOffice/program/../share/gallery"
	$asGalleryPath = $oPathSettings.Gallery_internal()
	If Not IsArray($asGalleryPath) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asGalleryPath) - 1
		$asGalleryPath[$i] = _LO_PathConvert($asGalleryPath[$i], $LO_PATHCONV_PCPATH_RETURN)

		$hSearch = FileFindFirstFile($asGalleryPath[$i] & "\sounds\*.wav")
		If ($hSearch <> -1) Then
			While 1
				$sFile = FileFindNextFile($hSearch)
				If @error Then ExitLoop

				If (UBound($asFiles) <= $iCount) Then ReDim $asFiles[UBound($asFiles) + 5]
				$asFiles[$iCount] = $asGalleryPath[$i] & "\sounds\" & $sFile
				$iCount += 1
			WEnd

			ExitLoop
		EndIf
	Next

	ReDim $asFiles[$iCount]

	Return SetError($__LO_STATUS_SUCCESS, UBound($asFiles), $asFiles)
EndFunc   ;==>_LOImpress_SlideSoundsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_SlideTransition
; Description ...: Set or Retrieve a Slide's Transition properties.
; Syntax ........: _LOImpress_SlideTransition(ByRef $oSlide[, $iTransition = Null[, $nDuration = Null[, $sSound = Null[, $bLoopSound = Null[, $nSlideAdvance = Null]]]]])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iTransition         - [optional] (0-78) Default is Null. The Transition effect. See Constants, $LOI_SLIDE_TRANSITION_* as defined in LibreOfficeImpress_Constants.au3.
;                  $nDuration           - [optional] (0-1000) Default is Null. The duration of the slide's transition effect, in seconds. L.O. 6.1+. See remarks.
;                  $sSound              - [optional] Default is Null. The path to the sound to play during slide transition. See remarks.
;                  $bLoopSound          - [optional] Default is Null. If True, the sound is repeated.
;                  $nSlideAdvance       - [optional] (-1-1000) Default is Null. The number of seconds before automatically advance the slide. Call with -1 to set to On Mouse Click.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iTransition not an Integer, less than 0 or greater than 78. See Constants, $LOI_SLIDE_TRANSITION_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $nDuration not a Number, less than 0 or greater than 1000.
;                  @Error: 1, @Extended: 4 = $sSound not a String.
;                  @Error: 1, @Extended: 5 = File called in $sSound does not exist.
;                  @Error: 1, @Extended: 6 = $bLoopSound not a Boolean.
;                  @Error: 1, @Extended: 7 = $nSlideAdvance not a Number, less than -1 or greater than 1000.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve the current Transition type.
;                  @Error: 3, @Extended: 2 = Failed to retrieve current Duration.
;                  @Error: 3, @Extended: 3 = Failed to retrieve current Sound value.
;                  @Error: 3, @Extended: 4 = Failed to retrieve current Slide advance value.
;                  @Error: 3, @Extended: 5 = Failed to set Transition type.
;                  @Error: 3, @Extended: 6 = Failed to convert Sound path to LibreOffice path.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTransition
;                  |                               2 = Error setting $nDuration
;                  |                               4 = Error setting $sSound
;                  |                               8 = Error setting $bLoopSound
;                  |                               16 = Error setting $nSlideAdvance
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Previous to LibreOffice 6.1 $nDuration was simply a three option selection of Slow, Medium, and Fast. To make both these work with this UDF the following method has been adopted:
;                  Previous to LibreOffice 6.1, if the Value called in $nDuration is from 0 to 1.99, the speed is set to Fast, if $nDuration is called with 2, Medium speed is set, and any value beyond 2.01 is considered Slow. This matches LibreOffice's internal behaviour.
;                  When retrieving current property values previous to LibreOffice 6.1, if Speed is set to Fast, 1 is returned for $nDuration. If Speed is set to Medium, 2 is returned. And if Speed is set to Slow, 3 is returned.
;                  $sSound can be called with an empty string to indicate that no sound should be played.
;                  If $sSound is called with the string "stop", this equals "Stop Previous Sound" in the UI.
;                  Otherwise call $sSound with a valid path to a sound file. See _LOImpress_SlideSoundsGetNames, to obtain a list of sound files included with Impress.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_SlideSoundsGetNames, _LOImpress_SlideLayout
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_SlideTransition(ByRef $oSlide, $iTransition = Null, $nDuration = Null, $sSound = Null, $bLoopSound = Null, $nSlideAdvance = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__LOI_CONST_CHANGE_MANUAL = 0, $__LOI_CONST_CHANGE_AUTO = 1 ;,  $__LOI_CONST_CHANGE_SEMI_MANUAL = 2 ; com.sun.star.presentation.DrawPage:Change
	Local Const $__LOI_CONST_SPEED_SLOW = 0, $__LOI_CONST_SPEED_MEDIUM = 1, $__LOI_CONST_SPEED_FAST = 2    ; com.sun.star.presentation:AnimationSpeed
	Local $iError = 0, $iCurrTransition, $nCurrDuration, $nCurrSlideAdvance
	Local $sCurrSound
	Local $avTransition[5]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransition, $nDuration, $sSound, $bLoopSound, $nSlideAdvance) Then
		$iCurrTransition = __LOImpress_Transition($oSlide)
		If Not IsInt($iCurrTransition) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If __LO_VersionCheck(6.1) Then
			$nCurrDuration = $oSlide.TransitionDuration()

		Else
			Switch $oSlide.Speed()
				Case $__LOI_CONST_SPEED_FAST ; 0 - 1.99 ; Matches L.O. Behaviour.
					$nCurrDuration = 1

				Case $__LOI_CONST_SPEED_MEDIUM ; 2
					$nCurrDuration = 2

				Case Else ; $__LOI_CONST_SPEED_SLOW
					$nCurrDuration = 3
			EndSwitch
		EndIf

		If Not IsNumber($nCurrDuration) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$sCurrSound = $oSlide.Sound()
		If IsBool($sCurrSound) And ($sCurrSound = True) Then $sCurrSound = "stop"
		If Not IsString($sCurrSound) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$sCurrSound = _LO_PathConvert($sCurrSound, $LO_PATHCONV_PCPATH_RETURN)

		If ($oSlide.Change() <> $__LOI_CONST_CHANGE_AUTO) Then
			$nCurrSlideAdvance = -1

		Else
			$nCurrSlideAdvance = $oSlide.HighResDuration()
		EndIf

		If Not IsNumber($nCurrSlideAdvance) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		__LO_ArrayFill($avTransition, $iCurrTransition, $nCurrDuration, $sCurrSound, $oSlide.LoopSound(), $nCurrSlideAdvance)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTransition)
	EndIf

	If ($iTransition <> Null) Then
		If Not __LO_IntIsBetween($iTransition, $LOI_SLIDE_TRANSITION_3D_VENETIAN_VERT, $LOI_SLIDE_TRANSITION_WIPE_TOP_TO_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		__LOImpress_Transition($oSlide, $iTransition)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		$iError = (__LOImpress_Transition($oSlide) = $iTransition) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($nDuration <> Null) Then
		If Not __LO_NumIsBetween($nDuration, 0, 1000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If __LO_VersionCheck(6.1) Then
			$oSlide.TransitionDuration = $nDuration
			$iError = ($oSlide.TransitionDuration() = $nDuration) ? ($iError) : (BitOR($iError, 2))

		Else
			Switch $nDuration
				Case 0 - 1.99 ; Matches L.O. Behaviour.
					$oSlide.Speed = $__LOI_CONST_SPEED_FAST
					$iError = ($oSlide.Speed() = $__LOI_CONST_SPEED_FAST) ? ($iError) : (BitOR($iError, 2))

				Case 2
					$oSlide.Speed = $__LOI_CONST_SPEED_MEDIUM
					$iError = ($oSlide.Speed() = $__LOI_CONST_SPEED_MEDIUM) ? ($iError) : (BitOR($iError, 2))

				Case Else
					$oSlide.Speed = $__LOI_CONST_SPEED_SLOW
					$iError = ($oSlide.Speed() = $__LOI_CONST_SPEED_SLOW) ? ($iError) : (BitOR($iError, 2))
			EndSwitch
		EndIf
	EndIf

	If ($sSound <> Null) Then
		If Not IsString($sSound) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If ($sSound = "") Then
			$oSlide.Sound = $sSound
			$iError = ($oSlide.Sound() = $sSound) ? ($iError) : (BitOR($iError, 4))

		ElseIf ($sSound = "stop") Then
			$oSlide.Sound = True
			$iError = ($oSlide.Sound() = True) ? ($iError) : (BitOR($iError, 4))

		Else
			If Not FileExists($sSound) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$sSound = _LO_PathConvert($sSound, $LO_PATHCONV_OFFICE_RETURN)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

			$oSlide.Sound = $sSound
			$iError = ($oSlide.Sound() = $sSound) ? ($iError) : (BitOR($iError, 4))
		EndIf
	EndIf

	If ($bLoopSound <> Null) Then
		If Not IsBool($bLoopSound) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oSlide.LoopSound = $bLoopSound
		$iError = ($oSlide.LoopSound() = $bLoopSound) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($nSlideAdvance <> Null) Then
		If Not __LO_NumIsBetween($nSlideAdvance, -1, 1000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		If ($nSlideAdvance = -1) Then
			$oSlide.Change = $__LOI_CONST_CHANGE_MANUAL
			$iError = ($oSlide.Change() = $__LOI_CONST_CHANGE_MANUAL) ? ($iError) : (BitOR($iError, 16))

		Else
			If ($oSlide.Change() <> $__LOI_CONST_CHANGE_AUTO) Then $oSlide.Change = $__LOI_CONST_CHANGE_AUTO
			$oSlide.Duration = $nSlideAdvance
			$oSlide.HighResDuration = $nSlideAdvance
			$iError = (($oSlide.Duration() = $nSlideAdvance) And ($oSlide.HighResDuration() = $nSlideAdvance)) ? ($iError) : (BitOR($iError, 16))
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_SlideTransition
