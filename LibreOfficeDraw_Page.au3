#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Draw
#include "LibreOfficeDraw_Internal.au3"
#include "LibreOfficeDraw_Constants.au3"

; Other includes for Draw

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Deleting, etc. L.O. Draw Pages.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LODraw_PageAdd
; _LODraw_PageBackColor
; _LODraw_PageBackFillStyle
; _LODraw_PageBackGradient
; _LODraw_PageBackTransparency
; _LODraw_PageBackTransparencyGradient
; _LODraw_PageCopy
; _LODraw_PageCurrent
; _LODraw_PageDeleteByIndex
; _LODraw_PageDeleteByObj
; _LODraw_PageExists
; _LODraw_PageFooter
; _LODraw_PageFormat
; _LODraw_PageGetObjByIndex
; _LODraw_PageGetObjByName
; _LODraw_PageHandoutFooter
; _LODraw_PageHandoutFormat
; _LODraw_PageHandoutGetObj
; _LODraw_PageHandoutHeader
; _LODraw_PageHandoutLayout
; _LODraw_PageHandoutMargins
; _LODraw_PageLayout
; _LODraw_PageMargins
; _LODraw_PageMasterAdd
; _LODraw_PageMasterBackColor
; _LODraw_PageMasterBackFillStyle
; _LODraw_PageMasterBackGradient
; _LODraw_PageMasterBackTransparency
; _LODraw_PageMasterBackTransparencyGradient
; _LODraw_PageMasterCurrent
; _LODraw_PageMasterDeleteByIndex
; _LODraw_PageMasterDeleteByObj
; _LODraw_PageMasterExists
; _LODraw_PageMasterFormat
; _LODraw_PageMasterGetObjByIndex
; _LODraw_PageMasterGetObjByName
; _LODraw_PageMasterMargins
; _LODraw_PageMasterName
; _LODraw_PageMasterNotesGetObj
; _LODraw_PageMastersGetCount
; _LODraw_PageMastersGetNames
; _LODraw_PageMove
; _LODraw_PageName
; _LODraw_PageNotesFooter
; _LODraw_PageNotesFormat
; _LODraw_PageNotesGetObj
; _LODraw_PageNotesHeader
; _LODraw_PageNotesMargins
; _LODraw_PagesGetCount
; _LODraw_PagesGetNames
; _LODraw_PageshowActiveSettings
; _LODraw_PageshowCustomCreate
; _LODraw_PageshowCustomDelete
; _LODraw_PageshowCustomModify
; _LODraw_PageshowCustomSetName
; _LODraw_PageshowIsRunning
; _LODraw_PageshowPresentationControl
; _LODraw_PageshowsCustomGetNames
; _LODraw_PageshowSettingsMode
; _LODraw_PageshowSettingsOptions
; _LODraw_PageshowSettingsRange
; _LODraw_PageshowStart
; _LODraw_PageshowStop
; _LODraw_PageSoundsGetNames
; _LODraw_PageTransition
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageAdd
; Description ...: Add a page to a presentation.
; Syntax ........: _LODraw_PageAdd(ByRef $oDoc[, $iPos = Null[, $sName = ""]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iPos                - [optional] Default is Null. The position to insert the new page in the collection of pages. 0 Based. See remarks.
;                  $sName               - [optional] Default is "". The unique name of the Page. If called with an empty string, LibreOffice automatically names it.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning new page's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPos not an Integer, less than 0 or greater than number of pages.
;                  @Error: 1, @Extended: 3 = $sName not a String.
;                  @Error: 1, @Extended: 4 = Name called in $sName already exists.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to create a page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $iPos is called with Null, the new page is inserted at the end.
;                  Call $iPos with the last page index to insert the page at the end. Call $iPos with 0 to insert the new page in the first page position.
;                  Due to limitations in the API, I have made a small workaround for inserting a page at the beginning. A dispatch is executed to move the page to the beginning. The current page will temporarily be set to the new page in order to move it.
; Related .......: _LODraw_PageDeleteByIndex, _LODraw_PageDeleteByObj, _LODraw_PageMasterAdd, _LODraw_PageExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageAdd(ByRef $oDoc, $iPos = Null, $sName = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPage, $oServiceManager, $oDispatcher, $oCurrPage
	Local $bMoveToFirst = False
	Local $aArray[0]
	Local $iCurrView

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($iPos = Null) Then $iPos = $oDoc.DrawPages.getCount()
	If Not __LO_IntIsBetween($iPos, 0, $oDoc.DrawPages.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sName <> "") And _LODraw_PageExists($oDoc, $sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iPos -= 1 ; -1 because when 0 is called in insertNewByIndex, it inserts it in position 1, etc. Also there is no way to insert a new page at position 0, so I made a workaround.

	If ($iPos = -1) Then
		$iPos = 0
		$bMoveToFirst = True
	EndIf

	$oPage = $oDoc.DrawPages.insertNewByIndex($iPos)
	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $bMoveToFirst Then
		$oCurrPage = $oDoc.getCurrentController.CurrentPage() ; Backup current page and view mode

		$iCurrView = __LODraw_DocCurrView($oDoc)

		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
		If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oDoc.getCurrentController.setCurrentPage($oPage)

		$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:MovePageFirst", "", 0, $aArray)

		If IsObj($oCurrPage) Then $oDoc.getCurrentController.setCurrentPage($oCurrPage) ; Restore current page and view mode.

		If IsInt($iCurrView) Then __LODraw_DocCurrView($oDoc, $iCurrView)
	EndIf

	If ($sName <> "") Then
		$oPage.Name = $sName
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPage)
EndFunc   ;==>_LODraw_PageAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageBackColor
; Description ...: Set or Retrieve the Page's background color.
; Syntax ........: _LODraw_PageBackColor(ByRef $oPage[, $iColor = Null])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Page background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Integer
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting as an Integer value. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
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
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the page, the Constant $LO_COLOR_OFF is returned.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LODraw_PageBackFillStyle, _LODraw_PageBackGradient, _LODraw_PageMasterBackColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageBackColor(ByRef $oPage, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBackground, $oDoc
	Local $iError = 0, $iCurColor

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oPage.Background()

	If __LO_VarsAreNull($iColor) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, $LO_COLOR_OFF) ; If no background is set, this will be void, instead of an Object.

		$iCurColor = __LODraw_ColorRemoveAlpha($oBackground.FillColor())
		If Not IsInt($iCurColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurColor)
	EndIf

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oDoc = __LODraw_GetParentDoc($oPage)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	$oBackground.FillStyle = $LOD_AREA_FILL_STYLE_SOLID
	$oBackground.FillColor = $iColor

	$oPage.Background = $oBackground
	$iError = ($oPage.Background.FillColor() = $iColor) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageBackFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LODraw_PageBackFillStyle(ByRef $oPage[, $bFillOff = False])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $bFillOff            - [optional] Default is False. If True, the Fill style will be set to Off. See remarks.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOD_AREA_FILL_STYLE_* as defined in LibreOfficeDraw_Constants.au3.
;                  @Error: 0, @Extended: 1, Return: 0 = Success. Fill style was successfully turned off.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $bFillOff not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Fill Style.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
;                  When the Fill style is disabled for a Page, the Fill properties are completely removed. This is how Draw works normally.
;                  $bFillOff will do nothing if it is called with False, and is not, of course, returned when retrieving the FillStyle value.
; Related .......: _LODraw_PageBackColor, _LODraw_PageBackGradient, _LODraw_PageMasterBackFillStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageBackFillStyle(ByRef $oPage, $bFillOff = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle
	Local $oBackground

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bFillOff) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bFillOff Then
		If IsObj($oPage.Background()) Then
			$oBackground = $oPage.Background
			If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, 0) ; If no Background Object, no Fillstyle is active.

			$oBackground.FillStyle = $LOD_AREA_FILL_STYLE_OFF
			$oPage.Background = $oBackground
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, 0)
	EndIf

	$oBackground = $oPage.Background
	If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 0, $LOD_AREA_FILL_STYLE_OFF) ; If no Background Object, no Fillstyle is active.

	$iFillStyle = $oBackground.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LODraw_PageBackFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageBackGradient
; Description ...: Set or Retrieve the settings for Page Background color Gradient.
; Syntax ........: _LODraw_PageBackGradient(ByRef $oPage[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $sGradientName       - [optional] Default is Null. A Preset Gradient Name. See remarks. See constants, $LOD_GRAD_NAME_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iType               - [optional] (-1-5) Default is Null. The gradient type to apply. See Constants, $LOD_GRAD_TYPE_* as defined in LibreOfficeDraw_Constants.au3.
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
;                  @Error: 0, @Extended: 2, Return: -1 = Success. All optional parameters were called with Null, no background is currently active for the page. Returning -1.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $sGradientName not a String.
;                  @Error: 1, @Extended: 3 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOD_GRAD_TYPE_* as defined in LibreOfficeDraw_Constants.au3.
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
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LODraw_PageBackColor, _LODraw_PageBackFillStyle, _LODraw_PageMasterBackGradient
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageBackGradient(ByRef $oPage, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBackground, $oDoc
	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $nRed, $nGreen, $nBlue
	Local $atColorStop
	Local $avGradient[11]
	Local $sGradName

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oPage.Background()

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

	$oDoc = __LODraw_GetParentDoc($oPage)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	$tStyleGradient = $oBackground.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oBackground.FillStyle() <> $LOD_AREA_FILL_STYLE_GRADIENT) Then $oBackground.FillStyle = $LOD_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		__LODraw_GradientPresets($oDoc, $oBackground, $tStyleGradient, $sGradientName)

		$oPage.Background = $oBackground

		$oBackground = $oPage.Background()
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$tStyleGradient = $oBackground.FillGradient()
		If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$iError = ($oBackground.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOD_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oBackground.FillStyle = $LOD_AREA_FILL_STYLE_OFF
			$oBackground.FillGradientName = ""
			$oPage.Background = $oBackground

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOD_GRAD_TYPE_LINEAR, $LOD_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

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

	If ($oBackground.FillGradientName() = "") Or __LODraw_GradientIsModified($tStyleGradient, $oBackground.FillGradientName()) Then
		$sGradName = __LODraw_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

		$oBackground.FillGradientName = $sGradName
		If ($oBackground.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)
	EndIf

	$oBackground.FillGradient = $tStyleGradient
	$oPage.Background = $oBackground

	; Error checking
	$iError = (__LO_VarsAreNull($iType)) ? $iError : ($oPage.Background.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2))
	$iError = (__LO_VarsAreNull($iXCenter)) ? $iError : ($oPage.Background.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8))
	$iError = (__LO_VarsAreNull($iYCenter)) ? $iError : ($oPage.Background.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16))
	$iError = (__LO_VarsAreNull($iAngle)) ? $iError : (($oPage.Background.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? $iError : ($oPage.Background.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64))
	$iError = (__LO_VarsAreNull($iFromColor)) ? $iError : ($oPage.Background.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128))
	$iError = (__LO_VarsAreNull($iToColor)) ? $iError : ($oPage.Background.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256))
	$iError = (__LO_VarsAreNull($iFromIntense)) ? $iError : ($oPage.Background.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512))
	$iError = (__LO_VarsAreNull($iToIntense)) ? $iError : ($oPage.Background.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageBackGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageBackTransparency
; Description ...: Set or retrieve Transparency settings for a Page.
; Syntax ........: _LODraw_PageBackTransparency(ByRef $oPage[, $iTransparency = Null])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $iTransparency       - [optional] (0-100) Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting for Transparency as an Integer. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
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
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the page, -1 is returned.
; Related .......: _LODraw_PageBackTransparencyGradient, _LODraw_PageMasterBackTransparency
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageBackTransparency(ByRef $oPage, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurTransp
	Local $oBackground, $oDoc

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oPage.Background()

	If __LO_VarsAreNull($iTransparency) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, -1) ; No background present.

		$iCurTransp = $oBackground.FillTransparence()
		If Not IsInt($iCurTransp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurTransp)
	EndIf

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oDoc = __LODraw_GetParentDoc($oPage)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	$oBackground.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oBackground.FillTransparence = $iTransparency

	$oPage.Background = $oBackground

	$iError = ($oPage.Background.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageBackTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageBackTransparencyGradient
; Description ...: Set or retrieve the Page's transparency gradient settings.
; Syntax ........: _LODraw_PageBackTransparencyGradient(ByRef $oPage[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $iType               - [optional] (-1-5) Default is Null. The type of transparency gradient to apply. See Constants, $LOD_GRAD_TYPE_* as defined in LibreOfficeDraw_Constants.au3. Call with $LOD_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error: 0, @Extended: 1, Return: -1 = Success. All optional parameters were called with Null no background is currently active for the page. Returning -1.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $iType Not an Integer, less than -1 or greater than 5. See constants, $LOD_GRAD_TYPE_* as defined in LibreOfficeDraw_Constants.au3.
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
; Related .......: _LODraw_PageBackTransparency, _LODraw_PageMasterBackTransparencyGradient
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageBackTransparencyGradient(ByRef $oPage, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop
	Local $oBackground, $oDoc
	Local $fValue

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oPage.Background()

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 2, -1)

		$tGradient = $oBackground.FillTransparenceGradient()
		If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LO_ArrayFill($aiTransparent, $tGradient.Style(), $tGradient.XOffset(), $tGradient.YOffset(), _
				($tGradient.Angle() / 10), $tGradient.Border(), __LODraw_TransparencyGradientConvert(Null, $tGradient.StartColor()), _
				__LODraw_TransparencyGradientConvert(Null, $tGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	$oDoc = __LODraw_GetParentDoc($oPage)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	$tGradient = $oBackground.FillTransparenceGradient()
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iType <> Null) Then
		If ($iType = $LOD_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oBackground.FillTransparenceGradientName = ""
			$oPage.Background = $oBackground

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOD_GRAD_TYPE_LINEAR, $LOD_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

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

		$tGradient.StartColor = __LODraw_TransparencyGradientConvert($iStart)

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

		$tGradient.EndColor = __LODraw_TransparencyGradientConvert($iEnd)

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
		$sTGradName = __LODraw_TransparencyGradientNameInsert($oDoc, $tGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		$oBackground.FillTransparenceGradientName = $sTGradName
		If ($oBackground.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
	EndIf

	$oBackground.FillTransparenceGradient = $tGradient
	$oPage.Background = $oBackground

	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oPage.Background.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oPage.Background.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oPage.Background.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((($oPage.Background.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oPage.Background.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iStart)) ? ($iError) : (($oPage.Background.FillTransparenceGradient.StartColor() = __LODraw_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iEnd)) ? ($iError) : (($oPage.Background.FillTransparenceGradient.EndColor() = __LODraw_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageBackTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageCopy
; Description ...: Create a copy of a page.
; Syntax ........: _LODraw_PageCopy(ByRef $oPage[, $iPos = Null])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $iPos                - [optional] Default is Null. The position to insert the new page in the collection of pages. 0 Based. See remarks.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully copied the page, returning the new page's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $iPos not an Integer, less than 0 or greater than number of pages.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Parent Document.
;                  @Error: 3, @Extended: 2 = Failed to copy page.
;                  @Error: 3, @Extended: 3 = Failed to identify copied page's position.
;                  @Error: 3, @Extended: 4 = Failed to move copied page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The copied page is inserted after the page to be copied.
;                  If $iPos is called with Null, the page is left in the position described above. Otherwise, due to limitations in the API, some dispatches are executed to move the page. The current page will temporarily be set to the new page in order to move it.
; Related .......: _LODraw_PageAdd, _LODraw_PageDeleteByIndex, _LODraw_PageDeleteByObj, _LODraw_PageMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageCopy(ByRef $oPage, $iPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc, $oNewPage, $oCurrPage, $oServiceManager, $oDispatcher
	Local $iNewPos, $iMove, $iCurrView
	Local $sDispatch
	Local $aArray[0]

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = __LODraw_GetParentDoc($oPage)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iPos <> Null) And Not __LO_IntIsBetween($iPos, 0, $oDoc.DrawPages.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oNewPage = $oDoc.Duplicate($oPage)
	If Not IsObj($oNewPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Failed to copy Page.

	If ($iPos <> Null) Then
		For $i = 0 To $oDoc.DrawPages.getCount() - 1
			If ($oDoc.DrawPages.getByIndex($i) = $oNewPage) Then
				$iNewPos = $i
				ExitLoop
			EndIf
			Sleep((IsInt($i / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		If Not IsInt($iNewPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$iMove = ($iNewPos > $iPos) ? ($iNewPos - $iPos) : (($iNewPos < $iPos) ? ($iPos - $iNewPos) : (0)) ; 0 = NewPos and current Pos are the same.
		$sDispatch = ($iNewPos > $iPos) ? (".uno:MovePageUp") : (".uno:MovePageDown")
		If ($iPos = 0) Then ; Move Page to beginning.
			$iMove = 1 ; Set to 1 so it will be called once.
			$sDispatch = ".uno:MovePageFirst"

		ElseIf ($iPos = $oDoc.DrawPages.getCount() - 1) Then ; Move page to end.
			$iMove = 1 ; Set to 1 so it will be called once.
			$sDispatch = ".uno:MovePageLast"
		EndIf

		$oCurrPage = $oDoc.getCurrentController.CurrentPage() ; Backup current page

		$iCurrView = __LODraw_DocCurrView($oDoc)

		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
		If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oDoc.getCurrentController.setCurrentPage($oNewPage)

		For $i = 0 To $iMove - 1
			$oDispatcher.executeDispatch($oDoc.CurrentController(), $sDispatch, "", 0, $aArray)
			Sleep((IsInt($i / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		For $i = 0 To $oDoc.DrawPages.getCount() - 1
			If ($oDoc.DrawPages.getByIndex($i) = $oNewPage) Then
				$iNewPos = $i
				ExitLoop
			EndIf
			Sleep((IsInt($i / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		If IsObj($oCurrPage) Then $oDoc.getCurrentController.setCurrentPage($oCurrPage) ; Restore current page and view mode.

		If IsInt($iCurrView) Then __LODraw_DocCurrView($oDoc, $iCurrView)

		If ($iNewPos <> $iPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNewPage)
EndFunc   ;==>_LODraw_PageCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageCurrent
; Description ...: Set or Retrieve the currently active page or master page.
; Syntax ........: _LODraw_PageCurrent(ByRef $oDoc[, $oObj = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $oObj                - [optional] Default is Null. A Page or Master Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, _LODraw_PageCopy, _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
; Return values .: Success: 1 or Object
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Object = Success. All optional parameters were called with Null, returning currently active page. @Extended page's type, see remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oObj not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current page's Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oObj
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current page.
;                  If this function fails to return an Object with processing error 1, it is possible the current view mode is set to Page sorter.
;                  You can only set the current page to either a Master page or a normal page. To change views to Notes, Handouts etc., see _LODraw_DocView.
;                  If the current view mode is set to Page outline or Page Notes, the current page Object is returned. If the current view mode is set to Master Page Notes or Master Page Handout, the current Master page Object is returned.
;                  When retrieving the current page, @Extended will be set to either $LOD_PAGE_VIEW_PAGE or $LOD_PAGE_VIEW_MASTER. See Constants, $LOD_PAGE_VIEW_* as defined in LibreOfficeDraw_Constants.au3. Use _LODraw_DocView to determine the current view mode active.
; Related .......: _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, _LODraw_PageMasterCurrent, _LODraw_DocView
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageCurrent(ByRef $oDoc, $oObj = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCurrPage
	Local $iError, $iPageType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($oObj) Then
		$oCurrPage = $oDoc.getCurrentController.CurrentPage()
		If Not IsObj($oCurrPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Could be because the current mode is set to Page Sorter.

		$iPageType = ($oDoc.getCurrentController.IsMasterPageMode()) ? ($LOD_PAGE_VIEW_MASTER) : ($LOD_PAGE_VIEW_PAGE)

		Return SetError($__LO_STATUS_SUCCESS, $iPageType, $oCurrPage)
	EndIf

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.getCurrentController.setCurrentPage($oObj)
	$iError = ($oDoc.getCurrentController.CurrentPage() = $oObj) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageDeleteByIndex
; Description ...: Delete a page by index.
; Syntax ........: _LODraw_PageDeleteByIndex(ByRef $oDoc, $iPage)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iPage              - The page to delete. 0 based.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Page was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPage not an Integer, less than 0 or greater than number of pages minus one.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve count of pages.
;                  @Error: 3, @Extended: 2 = Failed to retrieve page's Object.
;                  @Error: 3, @Extended: 3 = Failed to delete page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageDeleteByObj, _LODraw_PagesGetCount, _LODraw_PageMasterDeleteByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageDeleteByIndex(ByRef $oDoc, $iPage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPage
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iPage, 0, $oDoc.DrawPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iCount = $oDoc.DrawPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oPage = $oDoc.DrawPages.getByIndex($iPage)
	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.DrawPages.remove($oPage)
	If ($iCount = $oDoc.DrawPages.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to delete because the count is the same.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LODraw_PageDeleteByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageDeleteByObj
; Description ...: Delete a page using its Object.
; Syntax ........: _LODraw_PageDeleteByObj(ByRef $oPage)
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Page was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Parent Document.
;                  @Error: 3, @Extended: 2 = Failed to retrieve count of pages.
;                  @Error: 3, @Extended: 3 = Failed to delete page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageDeleteByIndex, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, _LODraw_PageMasterDeleteByObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageDeleteByObj(ByRef $oPage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc
	Local $iCount

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = __LODraw_GetParentDoc($oPage)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oDoc.DrawPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.DrawPages.Remove($oPage)
	If ($oDoc.DrawPages.getCount() = $iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to delete because the count is the same.

	$oPage = Null

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LODraw_PageDeleteByObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageExists
; Description ...: Check whether a page with a certain name exists in a document.
; Syntax ........: _LODraw_PageExists(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $sName               - The page name to check for.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Document contains a Page with the called name, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query for Page name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageAdd, _LODraw_PageName, _LODraw_PagesGetNames, _LODraw_PageMasterExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageExists(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bExists

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$bExists = $oDoc.Links.getByName("Page").Links.hasByName($sName)
	If Not IsBool($bExists) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bExists)
EndFunc   ;==>_LODraw_PageExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageFooter
; Description ...: Set or Retrieve Page Footer settings.
; Syntax ........: _LODraw_PageFooter(ByRef $oPage[, $bDateTime = Null[, $bDateTimeIsFixed = Null[, $sDateTimeValue = Null[, $iDateTimeFormat = Null[, $bFooter = Null[, $sFooterText = Null[, $bPageNum = Null]]]]]]])
; Parameters ....: $oPage              -  A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $bDateTime           - [optional] Default is Null. If True, a Date or Time entry is added to the footer of the page.
;                  $bDateTimeIsFixed    - [optional] Default is Null. If True, the Date or Time entry is fixed.
;                  $sDateTimeValue      - [optional] Default is Null. If $bDateTimeIsFixed is True, this is the custom date or time value to display.
;                  $iDateTimeFormat     - [optional] (4-112) Default is Null. If $bDateTimeIsFixed is False, the format to display the Date or Time in. See Constants, $LOD_PAGE_DT_FMT_* as defined in LibreOfficeDraw_Constants.au3.
;                  $bFooter             - [optional] Default is Null. If True, a Footer entry is added to the footer of the page.
;                  $sFooterText         - [optional] Default is Null. If $bFooter is True, the text to display in the footer of the Page.
;                  $bPageNum           - [optional] Default is Null. If True, a current Page number is added to the footer of the page.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $bDateTime not a Boolean.
;                  @Error: 1, @Extended: 3 = $bDateTimeIsFixed not a Boolean.
;                  @Error: 1, @Extended: 4 = $sDateTimeValue not a String.
;                  @Error: 1, @Extended: 5 = $iDateTimeFormat not an Integer, less than 4 or greater than 9 but not equal to one of the constant values. See Constants, $LOD_PAGE_DT_FMT_* as defined in LibreOfficeDraw_Constants.au3.
;                  @Error: 1, @Extended: 6 = $bFooter not a Boolean.
;                  @Error: 1, @Extended: 7 = $sFooterText not a String.
;                  @Error: 1, @Extended: 8 = $bPageNum not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bDateTime
;                  |                               2 = Error setting $bDateTimeIsFixed
;                  |                               4 = Error setting $sDateTimeValue
;                  |                               8 = Error setting $iDateTimeFormat
;                  |                               16 = Error setting $bFooter
;                  |                               32 = Error setting $sFooterText
;                  |                               64 = Error setting $bPageNum
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When retrieving current setting values, both $sDateTimeValue and $iDateTimeFormat may return a value. To determine which is currently valid, check $bDateTimeIsFixed. If $bDateTimeIsFixed is True, $sDateTimeValue is valid, else $iDateTimeFormat. If $bDateTime is false, neither will be valid.
;                  Skip first page, and Apply to all are not added to this function as they are not actual settings. The user can simulate these easily by making a loop to apply it to all pages, and skip the first page if required.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, _LODraw_PageHandoutFooter, _LODraw_PageNotesFooter
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageFooter(ByRef $oPage, $bDateTime = Null, $bDateTimeIsFixed = Null, $sDateTimeValue = Null, $iDateTimeFormat = Null, $bFooter = Null, $sFooterText = Null, $bPageNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFooter[7]
	Local $sAllowed = $LOD_PAGE_DT_FMT_24H_HM & ":" & $LOD_PAGE_DT_FMT_MMDDYY_24H_HM & ":" & $LOD_PAGE_DT_FMT_24H_HMS & ":" & $LOD_PAGE_DT_FMT_12H_HM_AMPM & ":" & $LOD_PAGE_DT_FMT_MMDDYY_12H_HM_AMPM & ":" & $LOD_PAGE_DT_FMT_12H_HMS_AMPM

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bDateTime, $bDateTimeIsFixed, $sDateTimeValue, $iDateTimeFormat, $bFooter, $sFooterText, $bPageNum) Then
		__LO_ArrayFill($avFooter, $oPage.IsDateTimeVisible(), $oPage.IsDateTimeFixed(), $oPage.DateTimeText(), $oPage.DateTimeFormat(), _
				$oPage.IsFooterVisible(), $oPage.FooterText(), $oPage.IsPageNumberVisible())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFooter)
	EndIf

	If ($bDateTime <> Null) Then
		If Not IsBool($bDateTime) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oPage.IsDateTimeVisible = $bDateTime

		$iError = ($oPage.IsDateTimeVisible() = $bDateTime) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bDateTimeIsFixed <> Null) Then
		If Not IsBool($bDateTimeIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPage.IsDateTimeFixed = $bDateTimeIsFixed

		$iError = ($oPage.IsDateTimeFixed() = $bDateTimeIsFixed) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sDateTimeValue <> Null) Then
		If Not IsString($sDateTimeValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPage.DateTimeText = $sDateTimeValue

		$iError = ($oPage.DateTimeText() = $sDateTimeValue) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iDateTimeFormat <> Null) Then
		If Not __LO_IntIsBetween($iDateTimeFormat, $LOD_PAGE_DT_FMT_MMDDYY, $LOD_PAGE_DT_FMT_DOW_MMMM_DD_YYYY, "", $sAllowed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPage.DateTimeFormat = $iDateTimeFormat

		$iError = ($oPage.DateTimeFormat() = $iDateTimeFormat) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bFooter <> Null) Then
		If Not IsBool($bFooter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPage.IsFooterVisible = $bFooter

		$iError = ($oPage.IsFooterVisible() = $bFooter) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($sFooterText <> Null) Then
		If Not IsString($sFooterText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPage.FooterText = $sFooterText

		$iError = ($oPage.FooterText() = $sFooterText) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bPageNum <> Null) Then
		If Not IsBool($bPageNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPage.IsPageNumberVisible = $bPageNum

		$iError = ($oPage.IsPageNumberVisible() = $bPageNum) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageFormat
; Description ...: Set or Retrieve the page format settings.
; Syntax ........: _LODraw_PageFormat(ByRef $oPage[, $iWidth = Null[, $iHeight = Null[, $iOrientation = Null]]])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $iWidth              - [optional] Default is Null. The Width of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOD_PAGE_WIDTH_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iHeight             - [optional] Default is Null. The Height of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOD_PAGE_HEIGHT_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iOrientation        - [optional] (0-1) Default is Null. The page orientation. See Constants, $LOD_PAGE_ORIENT_* as defined in LibreOfficeDraw_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 4 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants, $LOD_PAGE_ORIENT_* as defined in LibreOfficeDraw_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current page width.
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
;                  I am unable to find the properties to set for "FitObject to Paper Format", "Background covers margins", "Page numbers", and "Paper tray".
; Related .......: _LO_UnitConvert, _LODraw_PageLayout, _LODraw_PageMargins, _LODraw_PageHandoutFormat, _LODraw_PageMasterFormat, _LODraw_PageNotesFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageFormat(ByRef $oPage, $iWidth = Null, $iHeight = Null, $iOrientation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LODraw_Format($oPage, $iWidth, $iHeight, $iOrientation)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LODraw_PageFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageGetObjByIndex
; Description ...: Retrieve a Page's Object by index.
; Syntax ........: _LODraw_PageGetObjByIndex(ByRef $oDoc, $iPage)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iPage              - The page to retrieve. 0 based.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested page's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPage not an Integer, less than 0 or greater than number of pages minus one.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageGetObjByName, _LODraw_PagesGetCount, _LODraw_PageMasterGetObjByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageGetObjByIndex(ByRef $oDoc, $iPage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPage

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iPage, 0, $oDoc.DrawPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oPage = $oDoc.DrawPages.getByIndex($iPage)
	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPage)
EndFunc   ;==>_LODraw_PageGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageGetObjByName
; Description ...: Retrieve a Page's Object by name.
; Syntax ........: _LODraw_PageGetObjByName(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $sName               - The Page's name to retrieve the Object for.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested Page's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Page name called in $sName not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested Page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageGetObjByIndex, _LODraw_PagesGetNames, _LODraw_PageMasterGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageGetObjByName(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPage

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.Links.getByName("Page").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPage = $oDoc.Links.getByName("Page").Links.getByName($sName)
	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPage)
EndFunc   ;==>_LODraw_PageGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageFooter
; Description ...: Set or Retrieve handout page Footer settings.
; Syntax ........: _LODraw_PageFooter(ByRef $oHandout[, $bFooter = Null[, $sFooterText = Null[, $bPageNum = Null]]])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LODraw_PageHandoutGetObj function.
;                  $bFooter             - [optional] Default is Null. If True, a Footer entry is added to the footer of the page.
;                  $sFooterText         - [optional] Default is Null. If $bFooter is True, the text to display in the footer of the page.
;                  $bPageNum           - [optional] Default is Null. If True, a current Page number is added to the footer of the page.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oHandout not an Object.
;                  @Error: 1, @Extended: 2 = $bFooter not a Boolean.
;                  @Error: 1, @Extended: 3 = $sFooterText not a String.
;                  @Error: 1, @Extended: 4 = $bPageNum not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFooter
;                  |                               2 = Error setting $sFooterText
;                  |                               4 = Error setting $bPageNum
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Apply to all is not added to this function as they it is not an actual setting. The user can simulate this easily by making a loop to apply it to all pages.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  During basic testing, while the settings were successfully set, LibreOffice seems to ignore footer values set for handout pages.
; Related .......: _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, _LODraw_PageFooter, _LODraw_PageNotesFooter
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageHandoutFooter(ByRef $oHandout, $bFooter = Null, $sFooterText = Null, $bPageNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFooter[3]

	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bFooter, $sFooterText, $bPageNum) Then
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

	If ($bPageNum <> Null) Then
		If Not IsBool($bPageNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oHandout.IsPageNumberVisible = $bPageNum

		$iError = ($oHandout.IsPageNumberVisible() = $bPageNum) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageHandoutFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageHandoutFormat
; Description ...: Set or Retrieve the handout page format settings.
; Syntax ........: _LODraw_PageHandoutFormat(ByRef $oHandout[, $iWidth = Null[, $iHeight = Null[, $iOrientation = Null]]])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LODraw_PageHandoutGetObj function.
;                  $iWidth              - [optional] Default is Null. The Width of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOD_PAGE_WIDTH_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iHeight             - [optional] Default is Null. The Height of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOD_PAGE_HEIGHT_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iOrientation        - [optional] (0-1) Default is Null. The page orientation. See Constants, $LOD_PAGE_ORIENT_* as defined in LibreOfficeDraw_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oHandout not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 4 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants, $LOD_PAGE_ORIENT_* as defined in LibreOfficeDraw_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current page width.
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
; Related .......: _LO_UnitConvert, _LODraw_PageLayout, _LODraw_PageMargins, _LODraw_PageMasterFormat, _LODraw_PageNotesFormat, _LODraw_PageFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageHandoutFormat(ByRef $oHandout, $iWidth = Null, $iHeight = Null, $iOrientation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LODraw_Format($oHandout, $iWidth, $iHeight, $iOrientation)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LODraw_PageHandoutFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageHandoutGetObj
; Description ...: Retrieve the Handout page Object for an Draw document.
; Syntax ........: _LODraw_PageHandoutGetObj(ByRef $oDoc)
; Parameters ....: $oDoc                -  A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
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
; Related .......: _LODraw_PageNotesGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageHandoutGetObj(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oHandout

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oHandout = $oDoc.HandoutMasterPage()
	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oHandout)
EndFunc   ;==>_LODraw_PageHandoutGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageHandoutHeader
; Description ...: Set or Retrieve handout page header settings.
; Syntax ........: _LODraw_PageHandoutHeader(ByRef $oHandout[, $bHeader = Null[, $sHeaderText = Null[, $bDateTime = Null[, $bDateTimeIsFixed = Null[, $sDateTimeValue = Null[, $iDateTimeFormat = Null]]]]]])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LODraw_PageHandoutGetObj function.
;                  $bHeader             - [optional] Default is Null. If True, a Header entry is added to the Header of the page.
;                  $sHeaderText         - [optional] Default is Null. If $bHeader is True, the text to display in the Header of the page.
;                  $bDateTime           - [optional] Default is Null. If True, a Date or Time entry is added to the header of the page.
;                  $bDateTimeIsFixed    - [optional] Default is Null. If True, the Date or Time entry is fixed.
;                  $sDateTimeValue      - [optional] Default is Null. If $bDateTimeIsFixed is True, this is the custom date or time value to display.
;                  $iDateTimeFormat     - [optional] (4-112) Default is Null. If $bDateTimeIsFixed is False, the format to display the Date or Time in. See Constants, $LOD_PAGE_DT_FMT_* as defined in LibreOfficeDraw_Constants.au3.
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
;                  @Error: 1, @Extended: 7 = $iDateTimeFormat not an Integer, less than 4 or greater than 9 but not equal to one of the constant values. See Constants, $LOD_PAGE_DT_FMT_* as defined in LibreOfficeDraw_Constants.au3.
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
;                  Apply to all is not added to this function as it is not an actual setting. The user can simulate this easily by making a loop to apply it to all pages.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, _LODraw_PageNotesHeader
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageHandoutHeader(ByRef $oHandout, $bHeader = Null, $sHeaderText = Null, $bDateTime = Null, $bDateTimeIsFixed = Null, $sDateTimeValue = Null, $iDateTimeFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHeader[6]
	Local $sAllowed = $LOD_PAGE_DT_FMT_24H_HM & ":" & $LOD_PAGE_DT_FMT_MMDDYY_24H_HM & ":" & $LOD_PAGE_DT_FMT_24H_HMS & ":" & $LOD_PAGE_DT_FMT_12H_HM_AMPM & ":" & $LOD_PAGE_DT_FMT_MMDDYY_12H_HM_AMPM & ":" & $LOD_PAGE_DT_FMT_12H_HMS_AMPM

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
		If Not __LO_IntIsBetween($iDateTimeFormat, $LOD_PAGE_DT_FMT_MMDDYY, $LOD_PAGE_DT_FMT_DOW_MMMM_DD_YYYY, "", $sAllowed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oHandout.DateTimeFormat = $iDateTimeFormat

		$iError = ($oHandout.DateTimeFormat() = $iDateTimeFormat) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageHandoutHeader

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageHandoutLayout
; Description ...: Set or Retrieve the current Handout page's layout.
; Syntax ........: _LODraw_PageHandoutLayout(ByRef $oHandout[, $iLayout = Null])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LODraw_PageHandoutGetObj function.
;                  $iLayout             - [optional] (22-31) Default is Null. The layout format of the Handout page. See Constants, $LOD_HANDOUT_LAYOUT_* as defined in LibreOfficeDraw_Constants.au3.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current layout setting as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oHandout not an Object.
;                  @Error: 1, @Extended: 2 = $iLayout not an Integer, less than 22 or greater than 31. See Constants, $LOD_HANDOUT_LAYOUT_* as defined in LibreOfficeDraw_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Page's current layout.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLayout
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LODraw_PageHandoutFormat, _LODraw_PageHandoutMargins, _LODraw_PageLayout
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageHandoutLayout(ByRef $oHandout, $iLayout = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iCurrLayout

	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLayout) Then
		$iCurrLayout = $oHandout.Layout()
		If Not IsInt($iCurrLayout) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurrLayout)
	EndIf

	If Not __LO_IntIsBetween($iLayout, $LOD_HANDOUT_LAYOUT_ONE_PAGE, $LOD_HANDOUT_LAYOUT_NINE_PAGES) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oHandout.Layout = $iLayout
	$iError = ($oHandout.Layout() = $iLayout) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageHandoutLayout

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageHandoutMargins
; Description ...: Set or Retrieve the handout page margin settings.
; Syntax ........: _LODraw_PageHandoutMargins(ByRef $oHandout[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oHandout            - A Handout page object returned by a previous _LODraw_PageHandoutGetObj function.
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
; Related .......: _LO_UnitConvert, _LODraw_PageLayout, _LODraw_PageFormat, _LODraw_PageMasterMargins, _LODraw_PageNotesMargins, _LODraw_PageMargins
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageHandoutMargins(ByRef $oHandout, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oHandout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LODraw_Margins($oHandout, $iLeft, $iRight, $iTop, $iBottom)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LODraw_PageHandoutMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageLayout
; Description ...: Set or Retrieve the current Page's layout.
; Syntax ........: _LODraw_PageLayout(ByRef $oPage[, $iLayout = Null])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $iLayout             - [optional] (0-34) Default is Null. The layout format of the Page. See Constants, $LOD_PAGE_LAYOUT_* as defined in LibreOfficeDraw_Constants.au3.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current layout setting as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $iLayout not an Integer, less than 0 or greater than 34. See Constants, $LOD_PAGE_LAYOUT_* as defined in LibreOfficeDraw_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Page's current layout.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLayout
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LODraw_PageName, _LODraw_PageTransition, _LODraw_PageFormat, _LODraw_PageMargins, _LODraw_PageHandoutLayout
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageLayout(ByRef $oPage, $iLayout = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iCurrLayout

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLayout) Then
		$iCurrLayout = $oPage.Layout()
		If Not IsInt($iCurrLayout) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurrLayout)
	EndIf

	If Not __LO_IntIsBetween($iLayout, $LOD_PAGE_LAYOUT_TITLE, $LOD_PAGE_LAYOUT_TITLE_6_CONTENT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oPage.Layout = $iLayout
	$iError = ($oPage.Layout() = $iLayout) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageLayout

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMargins
; Description ...: Set or Retrieve the page page margin settings.
; Syntax ........: _LODraw_PageMargins(ByRef $oPage[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $iLeft               - [optional] Default is Null. The amount of space to leave between the left edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] Default is Null. The amount of space to leave between the right edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] Default is Null. The amount of space to leave between the upper edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] Default is Null. The amount of space to leave between the lower edge of the page and the page content. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
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
; Related .......: _LO_UnitConvert, _LODraw_PageLayout, _LODraw_PageFormat, _LODraw_PageHandoutMargins, _LODraw_PageMasterMargins, _LODraw_PageNotesMargins
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMargins(ByRef $oPage, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LODraw_Margins($oPage, $iLeft, $iRight, $iTop, $iBottom)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LODraw_PageMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterAdd
; Description ...: Add a master page to a presentation.
; Syntax ........: _LODraw_PageMasterAdd(ByRef $oDoc[, $iPos = Null[, $sName = ""[, $bBlank = True]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iPos                - [optional] Default is Null. The position to insert the new master page in the collection of pages. 0 Based. This is ignored if $bBlank is False.
;                  $sName               - [optional] Default is "". The unique name of the Master Page. If called with an empty string, LibreOffice automatically names it.
;                  $bBlank              - [optional] Default is True. If True, the new Master Page is blank. If False a Master Page with a preset layout is inserted. See remarks.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning new page's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPos not an Integer, less than 0 or greater than number of master pages.
;                  @Error: 1, @Extended: 3 = $sName not a String.
;                  @Error: 1, @Extended: 4 = Name called in $sName already exists.
;                  @Error: 1, @Extended: 5 = $bBlank not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to create a master page.
;                  @Error: 3, @Extended: 2 = Failed to retrieve master pages Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve count of master pages.
;                  @Error: 3, @Extended: 4 = Failed to retrieve master page Object
;                  @Error: 3, @Extended: 5 = Failed to identify new master page Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $iPos is called with Null, the new master page is inserted at the end.
;                  Call $iPos with the last master page index to insert the master page at the end. Call $iPos with 0 to insert the new master page at the beginning.
;                  When inserting a non-blank master page ($bBlank called with false), the new page will be inserted AFTER the last page, $iPos is ignored.
;                  This function uses two methods to insert a Master page. Using the API, the resulting new master page is blank, without text boxes etc., the second method uses a document dispatch command, which results in a normally formatted master page, like when you add a master page manually.
;                  When inserting a new page with $bBlank set to False, I use the dispatch command to accomplish the insertion, this method seems to only ever insert the new page at the end of all the pages.
;                  I have not found a way to import Master page from the LibreOffice templates yet.
; Related .......: _LODraw_PageMasterDeleteByIndex, _LODraw_PageMasterDeleteByObj, _LODraw_PageAdd, _LODraw_PageMasterExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterAdd(ByRef $oDoc, $iPos = Null, $sName = "", $bBlank = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMPage, $oServiceManager, $oDispatcher, $oMasters, $oMaster
	Local $aoMasters[0]
	Local $iMasters
	Local $aArray[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($iPos = Null) Then $iPos = ($bBlank) ? ($oDoc.MasterPages.getCount()) : ($oDoc.MasterPages.getCount() - 1) ; If I am inserting a Master using the dispatch, I have make position be 1 less than the count so I can retrieve the Object for the last master page.
	If ($iPos = $oDoc.MasterPages.getCount()) Then $iPos = $iPos - 1 ; If I am inserting a Master using the dispatch command, and the user called the last page position plus 1, I need to change it to be 1 less so I can retrieve the Object for the last master page.
	If Not __LO_IntIsBetween($iPos, 0, $oDoc.MasterPages.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sName <> "") And _LODraw_PageMasterExists($oDoc, $sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bBlank) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If $bBlank Then
		$oMPage = $oDoc.MasterPages.insertNewByIndex($iPos)
		If Not IsObj($oMPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else
		; When inserting a new Master using a Dispatch, I have to backup a copy of all current master page objects so I can identify the new page.
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

			Sleep((IsInt($i / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:InsertMasterPage", "", 0, $aArray)

		; Identify the new Master Page.
		For $i = 0 To $oMasters.getCount() - 1
			$oMaster = $oMasters.getByIndex($i)
			If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			For $j = 0 To $iMasters - 1
				; If the Object is a match, exit this loop and continue the top-level loop, bypassing the Objext assignment.
				If $aoMasters[$j] = $oMaster Then ContinueLoop 2

				Sleep((IsInt($j / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
			Next

			$oMPage = $oMaster
		Next

		If Not IsObj($oMPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	If ($sName <> "") Then $oMPage.Name = $sName

	Return SetError($__LO_STATUS_SUCCESS, 0, $oMPage)
EndFunc   ;==>_LODraw_PageMasterAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterBackColor
; Description ...: Set or Retrieve the Master Page's background color.
; Syntax ........: _LODraw_PageMasterBackColor(ByRef $oMaster[, $iColor = Null])
; Parameters ....: $oMaster             - A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Master Page background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
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
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the page, the Constant $LO_COLOR_OFF is returned.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LODraw_PageMasterBackFillStyle, _LODraw_PageMasterBackGradient, _LODraw_PageBackColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterBackColor(ByRef $oMaster, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBackground, $oDoc
	Local $iError = 0, $iCurColor

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBackground = $oMaster.Background()

	If __LO_VarsAreNull($iColor) Then
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, $LO_COLOR_OFF) ; If no background is set, this will be void, instead of an Object.

		$iCurColor = __LODraw_ColorRemoveAlpha($oBackground.FillColor())
		If Not IsInt($iCurColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurColor)
	EndIf

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oDoc = __LODraw_GetParentDoc($oMaster)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oMaster.Background = $oBackground
	EndIf

	$oBackground.FillStyle = $LOD_AREA_FILL_STYLE_SOLID
	$oBackground.FillColor = $iColor
	$iError = ($oMaster.Background.FillColor() = $iColor) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageMasterBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterBackFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LODraw_PageMasterBackFillStyle(ByRef $oMaster[, $bFillOff = False])
; Parameters ....: $oMaster             - A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
;                  $bFillOff            - [optional] Default is False. If True, the Fill style will be set to Off. See remarks.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOD_AREA_FILL_STYLE_* as defined in LibreOfficeDraw_Constants.au3.
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
;                  When the Fill style is disabled for a Master Page, the Fill properties are completely removed. This is how Draw works normally.
;                  $bFillOff will do nothing if it is called with False, and is not, of course, returned when retrieving the FillStyle value.
; Related .......: _LODraw_PageMasterBackColor, _LODraw_PageMasterBackGradient, _LODraw_PageBackFillStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterBackFillStyle(ByRef $oMaster, $bFillOff = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle
	Local $oBackground

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bFillOff) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bFillOff Then
		If IsObj($oMaster.Background()) Then
			$oBackground = $oMaster.Background
			If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 1, 0) ; If no Background Object, no Fillstyle is active.

			$oBackground.FillStyle = $LOD_AREA_FILL_STYLE_OFF
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, 0)
	EndIf

	$oBackground = $oMaster.Background()
	If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_SUCCESS, 0, $LOD_AREA_FILL_STYLE_OFF) ; If no Background Object, no Fillstyle is active.

	$iFillStyle = $oBackground.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LODraw_PageMasterBackFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterBackGradient
; Description ...: Set or Retrieve the settings for Master Page Background color Gradient.
; Syntax ........: _LODraw_PageMasterBackGradient(ByRef $oMaster[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oMaster             - A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
;                  $sGradientName       - [optional] Default is Null. A Preset Gradient Name. See remarks. See constants, $LOD_GRAD_NAME_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iType               - [optional] (-1-5) Default is Null. The gradient type to apply. See Constants, $LOD_GRAD_TYPE_* as defined in LibreOfficeDraw_Constants.au3.
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
;                  @Error: 0, @Extended: 2, Return: -1 = Success. All optional parameters were called with Null, no background is currently active for the master page. Returning -1.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $sGradientName not a String.
;                  @Error: 1, @Extended: 3 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOD_GRAD_TYPE_* as defined in LibreOfficeDraw_Constants.au3.
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
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LODraw_PageMasterBackColor, _LODraw_PageMasterBackFillStyle, _LODraw_PageBackGradient
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterBackGradient(ByRef $oMaster, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
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

	$oDoc = __LODraw_GetParentDoc($oMaster)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oMaster.Background = $oBackground
	EndIf

	$tStyleGradient = $oBackground.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oBackground.FillStyle() <> $LOD_AREA_FILL_STYLE_GRADIENT) Then $oBackground.FillStyle = $LOD_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		__LODraw_GradientPresets($oDoc, $oBackground, $tStyleGradient, $sGradientName)

		$tStyleGradient = $oBackground.FillGradient()
		If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$iError = ($oBackground.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOD_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oBackground.FillStyle = $LOD_AREA_FILL_STYLE_OFF
			$oBackground.FillGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOD_GRAD_TYPE_LINEAR, $LOD_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

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

	If ($oBackground.FillGradientName() = "") Or __LODraw_GradientIsModified($tStyleGradient, $oBackground.FillGradientName()) Then
		$sGradName = __LODraw_GradientNameInsert($oDoc, $tStyleGradient)
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
EndFunc   ;==>_LODraw_PageMasterBackGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterBackTransparency
; Description ...: Set or retrieve Transparency settings for a Master Page.
; Syntax ........: _LODraw_PageMasterBackTransparency(ByRef $oMaster[, $iTransparency = Null])
; Parameters ....: $oMaster             - A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
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
;                  If no background, of any kind (i.e. Solid fill, Gradient, etc., is set for the Master page, -1 is returned.
; Related .......: _LODraw_PageMasterBackTransparencyGradient, _LODraw_PageBackTransparency
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterBackTransparency(ByRef $oMaster, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
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
		$oDoc = __LODraw_GetParentDoc($oMaster)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oMaster.Background = $oBackground
	EndIf

	$oBackground.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oBackground.FillTransparence = $iTransparency

	$iError = ($oMaster.Background.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageMasterBackTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterBackTransparencyGradient
; Description ...: Set or retrieve the Master Page's transparency gradient settings.
; Syntax ........: _LODraw_PageMasterBackTransparencyGradient(ByRef $oMaster[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oMaster             - A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
;                  $iType               - [optional] (-1-5) Default is Null. The type of transparency gradient to apply. See Constants, $LOD_GRAD_TYPE_* as defined in LibreOfficeDraw_Constants.au3. Call with $LOD_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error: 0, @Extended: 1, Return: -1 = Success. All optional parameters were called with Null no background is currently active for the master page. Returning -1.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $iType Not an Integer, less than -1 or greater than 5. See constants, $LOD_GRAD_TYPE_* as defined in LibreOfficeDraw_Constants.au3.
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
;                  While these properties can be set successfully, LibreOffice doesn't seem to apply it to the master page, even when done using the UI.
; Related .......: _LODraw_PageMasterBackTransparency, _LODraw_PageBackTransparencyGradient
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterBackTransparencyGradient(ByRef $oMaster, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
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
				($tGradient.Angle() / 10), $tGradient.Border(), __LODraw_TransparencyGradientConvert(Null, $tGradient.StartColor()), _
				__LODraw_TransparencyGradientConvert(Null, $tGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	$oDoc = __LODraw_GetParentDoc($oMaster)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not IsObj($oBackground) Then ; Have to create the Background service.
		$oBackground = $oDoc.createInstance("com.sun.star.drawing.Background")
		If Not IsObj($oBackground) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oMaster.Background = $oBackground
	EndIf

	$tGradient = $oBackground.FillTransparenceGradient()
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iType <> Null) Then
		If ($iType = $LOD_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oBackground.FillTransparenceGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOD_GRAD_TYPE_LINEAR, $LOD_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

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

		$tGradient.StartColor = __LODraw_TransparencyGradientConvert($iStart)

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

		$tGradient.EndColor = __LODraw_TransparencyGradientConvert($iEnd)

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
		$sTGradName = __LODraw_TransparencyGradientNameInsert($oDoc, $tGradient)
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
	$iError = (__LO_VarsAreNull($iStart)) ? ($iError) : (($oMaster.Background.FillTransparenceGradient.StartColor() = __LODraw_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iEnd)) ? ($iError) : (($oMaster.Background.FillTransparenceGradient.EndColor() = __LODraw_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageMasterBackTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterCurrent
; Description ...: Set or Retrieve the currently applied Master page to a page.
; Syntax ........: _LODraw_PageMasterCurrent(ByRef $oPage[, $oMaster = Null])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $oMaster             - [optional] Default is Null. A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
; Return values .: Success: 1 or Object.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Object = Success. All optional parameters were called with Null, returning currently applied Master Page as an Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $oMaster not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve currently applied Master page.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $oMaster
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LODraw_PageMasterGetObjByIndex, _LODraw_PageMasterGetObjByName, _LODraw_PageCurrent, _LODraw_PageCurrent
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterCurrent(ByRef $oPage, $oMaster = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCurrMaster
	Local $iError

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($oMaster) Then
		$oCurrMaster = $oPage.MasterPage()
		If Not IsObj($oCurrMaster) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $oCurrMaster)
	EndIf

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oPage.MasterPage = $oMaster
	$iError = ($oPage.MasterPage() = $oMaster) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageMasterCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterDeleteByIndex
; Description ...: Delete a master page by index.
; Syntax ........: _LODraw_PageMasterDeleteByIndex(ByRef $oDoc, $iMaster)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iMaster             - The index of the master page to delete. 0 based.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Master page was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iMaster not an Integer, less than 0 or greater than number of Master pages minus one.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve count of master pages.
;                  @Error: 3, @Extended: 2 = Failed to retrieve master page's Object.
;                  @Error: 3, @Extended: 3 = Failed to delete master page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Trying to delete a Master Page that is used by a page will result in a processing error. I currently have no way of checking if a master page is free to be deleted.
; Related .......: _LODraw_PageMasterDeleteByObj, _LODraw_PageMastersGetCount, _LODraw_PageDeleteByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterDeleteByIndex(ByRef $oDoc, $iMaster)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMPage
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iMaster, 0, $oDoc.MasterPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iCount = $oDoc.MasterPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oMPage = $oDoc.MasterPages.getByIndex($iMaster)
	If Not IsObj($oMPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.MasterPages.remove($oMPage)
	If ($iCount = $oDoc.MasterPages.getCount()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to delete because the count is the same.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LODraw_PageMasterDeleteByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterDeleteByObj
; Description ...: Delete a master page using its Object.
; Syntax ........: _LODraw_PageMasterDeleteByObj(ByRef $oMaster)
; Parameters ....: $oMaster             -  A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Page was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Parent Document.
;                  @Error: 3, @Extended: 2 = Failed to retrieve count of master pages.
;                  @Error: 3, @Extended: 3 = Failed to delete master page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Trying to delete a Master Page that is used by a page will result in a processing error. I currently have no way of checking if a master page is free to be deleted.
; Related .......: _LODraw_PageMasterDeleteByIndex, _LODraw_PageMasterGetObjByIndex, _LODraw_PageMasterGetObjByName, _LODraw_PageDeleteByObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterDeleteByObj(ByRef $oMaster)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc
	Local $iCount

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = __LODraw_GetParentDoc($oMaster)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oDoc.MasterPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.MasterPages.Remove($oMaster)
	If ($oDoc.MasterPages.getCount() = $iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to delete because the count is the same.

	$oMaster = Null

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LODraw_PageMasterDeleteByObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterExists
; Description ...: Check whether a master page with a certain name exists in a document.
; Syntax ........: _LODraw_PageMasterExists(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $sName               - The master page name to check for.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Document contains a master Page with the called name, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query for master Page name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageMasterAdd, _LODraw_PageMasterName, _LODraw_PageMastersGetNames, _LODraw_PageExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterExists(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bExists

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$bExists = $oDoc.Links.getByName("Master Page").Links.hasByName($sName)
	If Not IsBool($bExists) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bExists)
EndFunc   ;==>_LODraw_PageMasterExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterFormat
; Description ...: Set or Retrieve the master page format settings.
; Syntax ........: _LODraw_PageMasterFormat(ByRef $oMaster[, $iWidth = Null[, $iHeight = Null[, $iOrientation = Null]]])
; Parameters ....: $oMaster             - A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
;                  $iWidth              - [optional] Default is Null. The Width of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOD_PAGE_WIDTH_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iHeight             - [optional] Default is Null. The Height of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOD_PAGE_HEIGHT_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iOrientation        - [optional] (0-1) Default is Null. The page orientation. See Constants, $LOD_PAGE_ORIENT_* as defined in LibreOfficeDraw_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 4 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants, $LOD_PAGE_ORIENT_* as defined in LibreOfficeDraw_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current page width.
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
;                  I am unable to find the properties to set for "FitObject to Paper Format", "Background covers margins", "Page numbers", and "Paper tray".
; Related .......: _LO_UnitConvert, _LODraw_PageMasterMargins, _LODraw_PageHandoutFormat, _LODraw_PageNotesFormat, _LODraw_PageFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterFormat(ByRef $oMaster, $iWidth = Null, $iHeight = Null, $iOrientation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LODraw_Format($oMaster, $iWidth, $iHeight, $iOrientation)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LODraw_PageMasterFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterGetObjByIndex
; Description ...: Retrieve a Master Page's Object by index.
; Syntax ........: _LODraw_PageMasterGetObjByIndex(ByRef $oDoc, $iMaster)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iMaster             - The index of the master page to retrieve. 0 based.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested master page's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iMaster not an Integer, less than 0 or greater than number of master pages minus one.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested master page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageMasterGetObjByName, _LODraw_PageMastersGetCount, _LODraw_PageGetObjByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterGetObjByIndex(ByRef $oDoc, $iMaster)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMPage

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iMaster, 0, $oDoc.MasterPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oMPage = $oDoc.MasterPages.getByIndex($iMaster)
	If Not IsObj($oMPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oMPage)
EndFunc   ;==>_LODraw_PageMasterGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterGetObjByName
; Description ...: Retrieve a Master Page's Object by name.
; Syntax ........: _LODraw_PageMasterGetObjByName(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $sName               - The Master Page's name to retrieve the Object for.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested Master Page's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Master Page name called in $sName not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested Master Page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I have not found a way to import Master page from the LibreOffice templates yet.
; Related .......: _LODraw_PageMasterGetObjByIndex, _LODraw_PageMastersGetNames, _LODraw_PageGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterGetObjByName(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMPage

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.Links.getByName("Master Page").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oMPage = $oDoc.Links.getByName("Master Page").Links.getByName($sName)
	If Not IsObj($oMPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oMPage)
EndFunc   ;==>_LODraw_PageMasterGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterMargins
; Description ...: Set or Retrieve the master page page margin settings.
; Syntax ........: _LODraw_PageMasterMargins(ByRef $oMaster[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oMaster             - A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
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
; Related .......: _LO_UnitConvert, _LODraw_PageMasterFormat, _LODraw_PageHandoutMargins, _LODraw_PageNotesMargins, _LODraw_PageMargins
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterMargins(ByRef $oMaster, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LODraw_Margins($oMaster, $iLeft, $iRight, $iTop, $iBottom)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LODraw_PageMasterMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterName
; Description ...: Set or Retrieve a Master Page's name.
; Syntax ........: _LODraw_PageMasterName(ByRef $oMaster[, $sName = Null])
; Parameters ....: $oMaster             - A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
;                  $sName               - [optional] Default is Null. The new name to set the Master page to. See Remarks.
; Return values .: Success: 1 or String.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. All optional parameters were called with Null, returning current Master Page name as a String.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oMaster not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Master Page name called in $sName already exists in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Master Page name.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Document Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If setting the Master page name to a name and a number, there is a good chance the name won't stay applied, as LibreOffice will assume it is an auto-numbered page.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LODraw_PageMasterExists, _LODraw_PageMastersGetNames, _LODraw_PageName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterName(ByRef $oMaster, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
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

	$oDoc = __LODraw_GetParentDoc($oMaster)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If $oDoc.Links.getByName("Master Page").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oMaster.Name = $sName
	$iError = ($oMaster.LinkDisplayName() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageMasterName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMasterNotesGetObj
; Description ...: Retrieve the Notes Object for a Master Page.
; Syntax ........: _LODraw_PageMasterNotesGetObj(ByRef $oMaster)
; Parameters ....: $oMaster             - A Master Page object returned by a previous _LODraw_PageMasterAdd, _LODraw_PageMasterGetObjByIndex, or _LODraw_PageMasterGetObjByName function.
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
; Related .......: _LODraw_PageNotesGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterNotesGetObj(ByRef $oMaster)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNotes

	If Not IsObj($oMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oNotes = $oMaster.NotesPage()
	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNotes)
EndFunc   ;==>_LODraw_PageMasterNotesGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMastersGetCount
; Description ...: Retrieve a count of master pages.
; Syntax ........: _LODraw_PageMastersGetCount(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning count of master pages contained in the document.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve a count of master pages.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This only returns a count of master pages already loaded into the document.
; Related .......: _LODraw_PageMasterDeleteByIndex, _LODraw_PageMasterGetObjByIndex, _LODraw_PagesGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMastersGetCount(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oDoc.MasterPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LODraw_PageMastersGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMastersGetNames
; Description ...: Retrieve an array of names for all Master Pages contained in the document.
; Syntax ........: _LODraw_PageMastersGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array containing all Master Page names. @Extended is set to the number of page names returned.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Master Pages Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve count of Master Pages.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Master Page name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This only returns a list of master page names already loaded into the document.
;                  I have not found a way to import Master page from the LibreOffice templates yet.
; Related .......: _LODraw_PageMasterExists, _LODraw_PageMasterGetObjByName, _LODraw_PagesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMastersGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
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

		Sleep((IsInt($i / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iMasters, $asMasters)
EndFunc   ;==>_LODraw_PageMastersGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMove
; Description ...: Move a page in the collection of pages.
; Syntax ........: _LODraw_PageMove(ByRef $oPage, $iPos)
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $iPos                - The position to move the page to in the collection of pages. 0 Based. See remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Page was successfully moved.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $iPos not an Integer, less than 0 or greater than number of pages minus 1.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Parent Document.
;                  @Error: 3, @Extended: 2 = Failed to identify page's current position.
;                  @Error: 3, @Extended: 3 = Failed to move page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Due to limitations in the API, some dispatches are executed to move the page. The current page will temporarily be set to the new page in order to move it.
; Related .......: _LODraw_PageCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMove(ByRef $oPage, $iPos)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc, $oCurrPage, $oServiceManager, $oDispatcher
	Local $iCurrPos, $iMove, $iCurrView
	Local $sDispatch
	Local $aArray[0]

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = __LODraw_GetParentDoc($oPage)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iPos <> Null) And Not __LO_IntIsBetween($iPos, 0, $oDoc.DrawPages.getCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	For $i = 0 To $oDoc.DrawPages.getCount() - 1
		If ($oDoc.DrawPages.getByIndex($i) = $oPage) Then
			$iCurrPos = $i
			ExitLoop
		EndIf
		Sleep((IsInt($i / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If Not IsInt($iCurrPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iMove = ($iCurrPos > $iPos) ? ($iCurrPos - $iPos) : (($iCurrPos < $iPos) ? ($iPos - $iCurrPos) : (0))     ; 0 = CurrPos and New Pos are the same.
	$sDispatch = ($iCurrPos > $iPos) ? (".uno:MovePageUp") : (".uno:MovePageDown")
	If ($iPos = 0) Then     ; Move Page to beginning.
		$iMove = 1    ; Set to 1 so it will be called once.
		$sDispatch = ".uno:MovePageFirst"

	ElseIf ($iPos = $oDoc.DrawPages.getCount() - 1) Then     ; Move page to end.
		$iMove = 1    ; Set to 1 so it will be called once.
		$sDispatch = ".uno:MovePageLast"
	EndIf

	$oCurrPage = $oDoc.getCurrentController.CurrentPage()     ; Backup current page

	$iCurrView = __LODraw_DocCurrView($oDoc)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oDoc.getCurrentController.setCurrentPage($oPage)

	For $i = 0 To $iMove - 1
		$oDispatcher.executeDispatch($oDoc.CurrentController(), $sDispatch, "", 0, $aArray)
		Sleep((IsInt($i / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	For $i = 0 To $oDoc.DrawPages.getCount() - 1
		If ($oDoc.DrawPages.getByIndex($i) = $oPage) Then
			$iCurrPos = $i
			ExitLoop
		EndIf
		Sleep((IsInt($i / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If IsObj($oCurrPage) Then $oDoc.getCurrentController.setCurrentPage($oCurrPage)     ; Restore current page and view mode.

	If IsInt($iCurrView) Then __LODraw_DocCurrView($oDoc, $iCurrView)

	If ($iCurrPos <> $iPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LODraw_PageMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageName
; Description ...: Set or Retrieve a Page's name.
; Syntax ........: _LODraw_PageName(ByRef $oPage[, $sName = Null])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $sName               - [optional] Default is Null. The new name to set the page to. See Remarks.
; Return values .: Success: 1 or String.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. All optional parameters were called with Null, returning current Page name as a String.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Page name called in $sName already exists in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Page name.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Document Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If setting the page name to a name and a number, there is a good chance the name won't stay applied, as LibreOffice will assume it is an auto-numbered page.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LODraw_PageExists, _LODraw_PagesGetNames, _LODraw_PageMasterName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageName(ByRef $oPage, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurrName
	Local $oDoc

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sName) Then
		$sCurrName = $oPage.LinkDisplayName()
		If Not IsString($sCurrName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrName)
	EndIf

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc = __LODraw_GetParentDoc($oPage)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If $oDoc.Links.getByName("Page").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPage.Name = $sName
	$iError = ($oPage.LinkDisplayName() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageFooter
; Description ...: Set or Retrieve notes page Footer settings.
; Syntax ........: _LODraw_PageFooter(ByRef $oNotes[, $bFooter = Null[, $sFooterText = Null[, $bPageNum = Null]]])
; Parameters ....: $oNotes              - A Notes page object returned by a previous _LODraw_PageNotesGetObj or _LODraw_PageMasterNotesGetObj function.
;                  $bFooter             - [optional] Default is Null. If True, a Footer entry is added to the footer of the page.
;                  $sFooterText         - [optional] Default is Null. If $bFooter is True, the text to display in the footer of the page.
;                  $bPageNum           - [optional] Default is Null. If True, a current Page number is added to the footer of the page.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oNotes not an Object.
;                  @Error: 1, @Extended: 2 = Object passed in $oNotes is a Master Notes Object.
;                  @Error: 1, @Extended: 3 = $bFooter not a Boolean.
;                  @Error: 1, @Extended: 4 = $sFooterText not a String.
;                  @Error: 1, @Extended: 5 = $bPageNum not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFooter
;                  |                               2 = Error setting $sFooterText
;                  |                               4 = Error setting $bPageNum
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Apply to all is not added to this function as they it is not an actual setting. The user can simulate this easily by making a loop to apply it to all pages.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  You can only set or retrieve footer property values for a page notes page, not a master notes page.
; Related .......: _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, _LODraw_PageFooter, _LODraw_PageHandoutFooter
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageNotesFooter(ByRef $oNotes, $bFooter = Null, $sFooterText = Null, $bPageNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFooter[3]

	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oNotes.supportsService("com.sun.star.drawing.MasterPage") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bFooter, $sFooterText, $bPageNum) Then
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

	If ($bPageNum <> Null) Then
		If Not IsBool($bPageNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oNotes.IsPageNumberVisible = $bPageNum

		$iError = ($oNotes.IsPageNumberVisible() = $bPageNum) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageNotesFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageNotesFormat
; Description ...: Set or Retrieve the notes page format settings.
; Syntax ........: _LODraw_PageNotesFormat(ByRef $oNotes[, $iWidth = Null[, $iHeight = Null[, $iOrientation = Null]]])
; Parameters ....: $oNotes              - A Notes page object returned by a previous _LODraw_PageNotesGetObj or _LODraw_PageMasterNotesGetObj function.
;                  $iWidth              - [optional] Default is Null. The Width of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOD_PAGE_WIDTH_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iHeight             - [optional] Default is Null. The Height of the page, may be a custom value in Hundredths of a Millimeter (HMM), or one of the constants, $LOD_PAGE_HEIGHT_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iOrientation        - [optional] (0-1) Default is Null. The page orientation. See Constants, $LOD_PAGE_ORIENT_* as defined in LibreOfficeDraw_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oNotes not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 4 = $iOrientation not an Integer, less than 0 or greater than 1. See Constants, $LOD_PAGE_ORIENT_* as defined in LibreOfficeDraw_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current page width.
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
; Related .......: _LO_UnitConvert, _LODraw_PageLayout, _LODraw_PageMargins, _LODraw_PageHandoutFormat, _LODraw_PageMasterFormat, _LODraw_PageFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageNotesFormat(ByRef $oNotes, $iWidth = Null, $iHeight = Null, $iOrientation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LODraw_Format($oNotes, $iWidth, $iHeight, $iOrientation)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LODraw_PageNotesFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageNotesGetObj
; Description ...: Retrieve the Notes Object for a Page.
; Syntax ........: _LODraw_PageNotesGetObj(ByRef $oPage)
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning Notes page Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Notes Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageMasterNotesGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageNotesGetObj(ByRef $oPage)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNotes

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oNotes = $oPage.NotesPage()
	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNotes)
EndFunc   ;==>_LODraw_PageNotesGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageNotesHeader
; Description ...: Set or Retrieve notes page header settings.
; Syntax ........: _LODraw_PageNotesHeader(ByRef $oNotes[, $bHeader = Null[, $sHeaderText = Null[, $bDateTime = Null[, $bDateTimeIsFixed = Null[, $sDateTimeValue = Null[, $iDateTimeFormat = Null]]]]]])
; Parameters ....: $oNotes              - A Notes page object returned by a previous _LODraw_PageNotesGetObj function.
;                  $bHeader             - [optional] Default is Null. If True, a Header entry is added to the Header of the page.
;                  $sHeaderText         - [optional] Default is Null. If $bHeader is True, the text to display in the Header of the page.
;                  $bDateTime           - [optional] Default is Null. If True, a Date or Time entry is added to the header of the page.
;                  $bDateTimeIsFixed    - [optional] Default is Null. If True, the Date or Time entry is fixed.
;                  $sDateTimeValue      - [optional] Default is Null. If $bDateTimeIsFixed is True, this is the custom date or time value to display.
;                  $iDateTimeFormat     - [optional] (4-112) Default is Null. If $bDateTimeIsFixed is False, the format to display the Date or Time in. See Constants, $LOD_PAGE_DT_FMT_* as defined in LibreOfficeDraw_Constants.au3.
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
;                  @Error: 1, @Extended: 8 = $iDateTimeFormat not an Integer, less than 4 or greater than 9 but not equal to one of the constant values. See Constants, $LOD_PAGE_DT_FMT_* as defined in LibreOfficeDraw_Constants.au3.
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
;                  Apply to all is not added to this function as it is not an actual setting. The user can simulate this easily by making a loop to apply it to all pages.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  You can only set or retrieve header property values for a page notes page, not a master notes page.
; Related .......: _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, _LODraw_PageHandoutHeader
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageNotesHeader(ByRef $oNotes, $bHeader = Null, $sHeaderText = Null, $bDateTime = Null, $bDateTimeIsFixed = Null, $sDateTimeValue = Null, $iDateTimeFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHeader[7]
	Local $sAllowed = $LOD_PAGE_DT_FMT_24H_HM & ":" & $LOD_PAGE_DT_FMT_MMDDYY_24H_HM & ":" & $LOD_PAGE_DT_FMT_24H_HMS & ":" & $LOD_PAGE_DT_FMT_12H_HM_AMPM & ":" & $LOD_PAGE_DT_FMT_MMDDYY_12H_HM_AMPM & ":" & $LOD_PAGE_DT_FMT_12H_HMS_AMPM

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
		If Not __LO_IntIsBetween($iDateTimeFormat, $LOD_PAGE_DT_FMT_MMDDYY, $LOD_PAGE_DT_FMT_DOW_MMMM_DD_YYYY, "", $sAllowed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oNotes.DateTimeFormat = $iDateTimeFormat

		$iError = ($oNotes.DateTimeFormat() = $iDateTimeFormat) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageNotesHeader

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageNotesMargins
; Description ...: Set or Retrieve the notes page margin settings.
; Syntax ........: _LODraw_PageNotesMargins(ByRef $oNotes[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oNotes              - A Notes page object returned by a previous _LODraw_PageNotesGetObj or _LODraw_PageMasterNotesGetObj function.
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
; Related .......: _LO_UnitConvert, _LODraw_PageLayout, _LODraw_PageFormat, _LODraw_PageHandoutMargins, _LODraw_PageMasterMargins, _LODraw_PageMargins
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageNotesMargins(ByRef $oNotes, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oNotes) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LODraw_Margins($oNotes, $iLeft, $iRight, $iTop, $iBottom)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LODraw_PageNotesMargins

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PagesGetCount
; Description ...: Retrieve a count of pages.
; Syntax ........: _LODraw_PagesGetCount(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning count of pages contained in the document.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve a count of pages.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageDeleteByIndex, _LODraw_PageGetObjByIndex, _LODraw_PageMastersGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PagesGetCount(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oDoc.DrawPages.getCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LODraw_PagesGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PagesGetNames
; Description ...: Retrieve an array of names for all Pages contained in the document.
; Syntax ........: _LODraw_PagesGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array containing all Page names. @Extended is set to the number of page names returned.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve array of Page names.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageExists, _LODraw_PageGetObjByName, _LODraw_PageMastersGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PagesGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asPages[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asPages = $oDoc.Links.getByName("Page").Links.getElementNames()
	If Not IsArray($asPages) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asPages), $asPages)
EndFunc   ;==>_LODraw_PagesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowActiveSettings
; Description ...: Set or Retrieve settings for an actively running presentation.
; Syntax ........: _LODraw_PageshowActiveSettings(ByRef $oDoc[, $bKeepOnTop = Null[, $bMouseVisible = Null[, $bMouseAsPen = Null[, $iPenColor = Null[, $iPenWidth = Null]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $bKeepOnTop          - [optional] Default is Null. If True, the presentation will be always kept on top of other programs.
;                  $bMouseVisible       - [optional] Default is Null. If True, the mouse is visible in the presentation.
;                  $bMouseAsPen         - [optional] Default is Null. If True, the mouse can be used as a pen to draw on pages.
;                  $iPenColor           - [optional] (0-16777215) Default is Null. If $bMouseAsPen is True, the color of the drawn line, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iPenWidth           - [optional] (4-400) Default is Null. The width of the drawn line. L.O. 4.2+. See Constants, $LOD_SLIDESHOW_PEN_WIDTH_* as defined in LibreOfficeDraw_Constants.au3.
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
;                  @Error: 1, @Extended: 6 = $iPenWidth not an Integer, less than 4 or greater than 400. See Constants, $LOD_SLIDESHOW_PEN_WIDTH_* as defined in LibreOfficeDraw_Constants.au3.
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
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LODraw_PageshowIsRunning, _LODraw_PageshowSettingsMode, _LODraw_PageshowSettingsOptions, _LODraw_PageshowSettingsRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowActiveSettings(ByRef $oDoc, $bKeepOnTop = Null, $bMouseVisible = Null, $bMouseAsPen = Null, $iPenColor = Null, $iPenWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oPresentation
	Local $avPageShow[5]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.Presentation.isRunning() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; No Pageshow active.

	$oPresentation = $oDoc.Presentation.getController()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($bKeepOnTop, $bMouseVisible, $bMouseAsPen, $iPenColor, $iPenWidth) Then
		If __LO_VersionCheck(4.2) Then
			__LO_ArrayFill($avPageShow, $oPresentation.AlwaysOnTop(), $oPresentation.MouseVisible(), $oPresentation.UsePen(), $oPresentation.PenColor(), $oPresentation.PenWidth())

		Else
			__LO_ArrayFill($avPageShow, $oPresentation.AlwaysOnTop(), $oPresentation.MouseVisible(), $oPresentation.UsePen(), $oPresentation.PenColor(), Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPageShow)
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
		If Not __LO_IntIsBetween($iPenWidth, $LOD_SLIDESHOW_PEN_WIDTH_VERY_THIN, $LOD_SLIDESHOW_PEN_WIDTH_VERY_THICK) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPresentation.PenWidth = $iPenWidth

		$iError = ($oPresentation.PenWidth() = $iPenWidth) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageshowActiveSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowCustomCreate
; Description ...: Create a Custom Pageshow.
; Syntax ........: _LODraw_PageshowCustomCreate(ByRef $oDoc, $sName, $asPages)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $sName               - The name of the Custom Pageshow to create.
;                  $asPages            - A single column Array of Page names. See remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully created new Custom Pageshow.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Name called in $sName already exists as a Custom Pageshow in Document.
;                  @Error: 1, @Extended: 4 = $asPages not an Array.
;                  @Error: 1, @Extended: 5 = Array called in $asPages has 0 elements.
;                  @Error: 1, @Extended: 6 = Element contained in $asPages not a String. Returning problem element number.
;                  @Error: 1, @Extended: 7 = Page name contained in $asPages not found in Document. Returning problem element number.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a CustomPresentation Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve the Links Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Page's Object.
;                  @Error: 3, @Extended: 3 = Failed to insert new Custom Pageshow.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The expected input for $asPages is a single column array having the Page names in the order the user wishes the Pages to appear in the presentation, page names can be placed in the Array multiple times.
; Related .......: _LODraw_PageshowCustomDelete, _LODraw_PageshowsCustomGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowCustomCreate(ByRef $oDoc, $sName, $asPages)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oLinks, $oCustomPres, $oPage

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsArray($asPages) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If (UBound($asPages) < 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oLinks = $oDoc.Links.getByName("Page").Links()
	If Not IsObj($oLinks) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($asPages) - 1
		If Not IsString($asPages[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)
		If Not $oLinks.hasByName($asPages[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $i)
	Next

	$oCustomPres = $oDoc.CustomPresentations.createInstance()
	If Not IsObj($oCustomPres) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($asPages) - 1
		$oPage = $oLinks.getByName($asPages[$i])
		If Not IsObj($oPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oCustomPres.insertByIndex($oCustomPres.getCount(), $oPage)
	Next

	$oDoc.CustomPresentations.insertByName($sName, $oCustomPres)
	If Not $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LODraw_PageshowCustomCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowCustomDelete
; Description ...: Deletes a Custom Pageshow.
; Syntax ........: _LODraw_PageshowCustomDelete(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $sName               - The Custom Pageshow's name to delete.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Custom Pageshow was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Name called in $sName not found as a Custom Pageshow in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to delete the requested Custom Pageshow.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageshowCustomCreate, _LODraw_PageshowsCustomGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowCustomDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oDoc.CustomPresentations.removeByName($sName)
	If $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LODraw_PageshowCustomDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowCustomModify
; Description ...: Set or Retrieve the Pages and order of the pages contained in a Custom Pageshow.
; Syntax ........: _LODraw_PageshowCustomModify(ByRef $oDoc, $sName[, $asPages = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $sName               - The name of the Custom Pageshow to modify.
;                  $asPages            - [optional] Default is Null. A single column Array of Page names. See remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Custom Pageshow successfully modified.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning Array of Page names contained in the Custom Pageshow. See remarks.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Name called in $sName already exists as a Custom Pageshow in Document.
;                  @Error: 1, @Extended: 4 = $asPages not an Array.
;                  @Error: 1, @Extended: 5 = Array called in $asPages has 0 elements.
;                  @Error: 1, @Extended: 6 = Element contained in $asPages not a String. Returning problem element number.
;                  @Error: 1, @Extended: 7 = Page name contained in $asPages not found in Document. Returning problem element number.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a CustomPresentation Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Custom Pageshow's Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Page's name.
;                  @Error: 3, @Extended: 3 = Failed to retrieve the Links Object.
;                  @Error: 3, @Extended: 4 = Failed to retrieve Page's Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $asPages
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The expected input for $asPages is a single column array having the Page names in the order the user wishes the Pages to appear in the presentation, page names can be placed in the Array multiple times.
;                  When retrieving the current order and content of the Pageshow, an array is returned with all the Pages contained in the Custom Pageshow, in the order they are set to be played. Pages may be present multiple times.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LODraw_PageshowCustomCreate, _LODraw_PageshowsCustomGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowCustomModify(ByRef $oDoc, $sName, $asPages = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oLinks, $oCustomPres, $oNewCustomPres, $oPage
	Local $asCurrPages[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.CustomPresentations.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($asPages) Then
		$oCustomPres = $oDoc.CustomPresentations.getByName($sName)
		If Not IsObj($oCustomPres) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		ReDim $asCurrPages[$oCustomPres.getCount()]
		For $i = 0 To $oCustomPres.getCount() - 1
			$asCurrPages[$i] = $oCustomPres.getByIndex($i).LinkDisplayName()
			If Not IsString($asCurrPages[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Sleep((IsInt($i / $__LODCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $asCurrPages)
	EndIf

	If Not IsArray($asPages) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If (UBound($asPages) < 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oLinks = $oDoc.Links.getByName("Page").Links()
	If Not IsObj($oLinks) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asPages) - 1
		If Not IsString($asPages[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)
		If Not $oLinks.hasByName($asPages[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $i)
	Next

	$oNewCustomPres = $oDoc.CustomPresentations.createInstance()
	If Not IsObj($oNewCustomPres) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($asPages) - 1
		$oPage = $oLinks.getByName($asPages[$i])
		If Not IsObj($oPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oNewCustomPres.insertByIndex($oNewCustomPres.getCount(), $oPage)
	Next

	$oDoc.CustomPresentations.replaceByName($sName, $oNewCustomPres)
	$iError = ($oDoc.CustomPresentations.getByName($sName).getCount() = UBound($asPages)) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageshowCustomModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowCustomSetName
; Description ...: Rename a Custom Pageshow.
; Syntax ........: _LODraw_PageshowCustomSetName(ByRef $oDoc, $sName, $sNewName)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $sName               - The name of the Custom Pageshow to rename.
;                  $sNewName            - The name to rename the Custom Pageshow to.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Custom Pageshow was successfully renamed.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Name called in $sName not found as a Custom Pageshow in Document.
;                  @Error: 1, @Extended: 4 = $sNewName not a String.
;                  @Error: 1, @Extended: 5 = Name called in $sNewName already exists as a Custom Pageshow in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested Custom Pageshow Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sNewName
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageshowCustomModify, _LODraw_PageshowsCustomGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowCustomSetName(ByRef $oDoc, $sName, $sNewName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
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
EndFunc   ;==>_LODraw_PageshowCustomSetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowIsRunning
; Description ...: Check whether there is a presentation currently running.
; Syntax ........: _LODraw_PageshowIsRunning(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
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
; Related .......: _LODraw_PageshowStart, _LODraw_PageshowStop
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowIsRunning(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsRunning

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsRunning = $oDoc.Presentation.IsRunning()
	If Not IsBool($bIsRunning) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsRunning)
EndFunc   ;==>_LODraw_PageshowIsRunning

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowPresentationControl
; Description ...: Query the status of, or send commands to, a currently running presentation.
; Syntax ........: _LODraw_PageshowPresentationControl(ByRef $oDoc, $iAction[, $vValue = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iAction             - The Query or Command to perform on the presentation. See Constants, $LOD_SLIDESHOW_PRES_* as defined in LibreOfficeDraw_Constants.au3.
;                  $vValue              - [optional] Default is Null. If the Query or Command requires an input value, it goes here. See Remarks.
; Return values .: Success: Boolean, Integer, or Object.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully processed a command.
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Successfully processed a query that returns a Boolean. (See description of the specific query to see what is returned.)
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Successfully processed a query that returns an Integer. (See description of the specific query to see what is returned.)
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully processed a query that returns an Object. (See description of the specific query to see what is returned.)
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iAction not an Integer, less than 0 or greater than 25. See Constants, $LOD_SLIDESHOW_PRES_* as defined in LibreOfficeDraw_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iAction called with $LOD_SLIDESHOW_PRES_QUERY_GET_PAGE_BY_INDEX, and index value called in $vValue is not an Integer, less than 0 or greater than number of pages in the Presentation.
;                  @Error: 1, @Extended: 4 = $iAction called with $LOD_SLIDESHOW_PRES_COMMAND_ACTIVATE_BLANK_SCREEN, and color value called in $vValue is not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iAction called with $LOD_SLIDESHOW_PRES_COMMAND_GOTO_PAGE, and value called in $vValue is not an Object.
;                  @Error: 1, @Extended: 6 = $iAction called with $LOD_SLIDESHOW_PRES_COMMAND_GOTO_PAGE_BY_INDEX, and index value called in $vValue is not an Integer, less than 0 or greater than number of pages in the Presentation.
;                  @Error: 1, @Extended: 7 = $iAction called with $LOD_SLIDESHOW_PRES_COMMAND_GOTO_PAGE_BY_NAME, and value called in $vValue is not a String.
;                  @Error: 1, @Extended: 8 = $iAction called with $LOD_SLIDESHOW_PRES_COMMAND_GOTO_PAGE_BY_NAME, and Page name called in $vValue does not exist.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = There is no presentation currently running.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Object for currently running presentation.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Object for current page.
;                  @Error: 3, @Extended: 4 = Failed to retrieve current page's index value.
;                  @Error: 3, @Extended: 5 = Failed to retrieve next page's index value.
;                  @Error: 3, @Extended: 6 = Failed to retrieve Object for requested page.
;                  @Error: 3, @Extended: 7 = Failed to retrieve count of pages in the presentation.
;                  @Error: 3, @Extended: 8 = Failed to determine if the presentation is Active.
;                  @Error: 3, @Extended: 9 = Failed to determine if the presentation is Endless.
;                  @Error: 3, @Extended: 10 = Failed to determine if the presentation is FullScreen.
;                  @Error: 3, @Extended: 11 = Failed to determine if the presentation is Paused.
;                  @Error: 3, @Extended: 12 = Failed to activate presentation.
;                  @Error: 3, @Extended: 13 = Failed to activate blank screen for presentation.
;                  @Error: 3, @Extended: 14 = Failed to move to first page.
;                  @Error: 3, @Extended: 15 = Failed to move to last page.
;                  @Error: 3, @Extended: 16 = Failed to move to requested page by Object.
;                  @Error: 3, @Extended: 17 = Failed to move to requested page by Index.
;                  @Error: 3, @Extended: 18 = Failed to move to requested page by Name.
;                  @Error: 3, @Extended: 19 = Failed to Pause the presentation.
;                  @Error: 3, @Extended: 20 = Failed to Resume the presentation.
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = $iAction called with $LOD_SLIDESHOW_PRES_COMMAND_ERASE_ALL_INK, and current LibreOffice version is less than 7.2.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Any queries or commands that require an input parameter will have the type of input required indicated in the description for the Constant.
; Related .......: _LODraw_PageshowActiveSettings, _LODraw_PageshowIsRunning
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowPresentationControl(ByRef $oDoc, $iAction, $vValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn = 1
	Local $oPresentation

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iAction, $LOD_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE, $LOD_SLIDESHOW_PRES_COMMAND_STOP_SOUND) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.Presentation.isRunning() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; No Pageshow active.

	$oPresentation = $oDoc.Presentation.getController()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Switch $iAction
		Case $LOD_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE
			$vReturn = $oPresentation.getCurrentPage()
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOD_SLIDESHOW_PRES_QUERY_GET_CURRENT_PAGE_INDEX
			$vReturn = $oPresentation.getCurrentPageIndex()
			If Not IsInt($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Case $LOD_SLIDESHOW_PRES_QUERY_GET_NEXT_PAGE_INDEX
			$vReturn = $oPresentation.getNextPageIndex()
			If Not IsInt($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		Case $LOD_SLIDESHOW_PRES_QUERY_GET_PAGE_BY_INDEX
			If Not __LO_IntIsBetween($vValue, 0, $oPresentation.getPageCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$vReturn = $oPresentation.getPageByIndex($vValue)
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

		Case $LOD_SLIDESHOW_PRES_QUERY_GET_PAGE_COUNT
			$vReturn = $oPresentation.getPageCount()
			If Not IsInt($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

		Case $LOD_SLIDESHOW_PRES_QUERY_IS_ACTIVE
			$vReturn = $oPresentation.isActive()
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

		Case $LOD_SLIDESHOW_PRES_QUERY_IS_ENDLESS
			$vReturn = $oPresentation.isEndless()
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 9, 0)

		Case $LOD_SLIDESHOW_PRES_QUERY_IS_FULLSCREEN
			$vReturn = $oPresentation.isFullScreen()
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 10, 0)

		Case $LOD_SLIDESHOW_PRES_QUERY_IS_PAUSED
			$vReturn = $oPresentation.isPaused()
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 11, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_ACTIVATE
			$oPresentation.activate()
			If Not $oPresentation.isActive() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 12, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_ACTIVATE_BLANK_SCREEN
			If Not __LO_IntIsBetween($vValue, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oPresentation.blankScreen($vValue)
			If Not $oPresentation.isPaused() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 13, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_DEACTIVATE
			$oPresentation.deactivate() ; Doesn't seem to set IsActive to False!
;~ 			If $oPresentation.isActive() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 14, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_ERASE_ALL_INK
			If Not __LO_VersionCheck(7.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

			$oPresentation.setEraseAllInk(True)

		Case $LOD_SLIDESHOW_PRES_COMMAND_GOTO_FIRST_SLIDE
			$oPresentation.gotoFirstPage()
			If ($oPresentation.getCurrentPageIndex() <> 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 14, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_GOTO_LAST_SLIDE
			$oPresentation.gotoLastPage()
			If ($oPresentation.getCurrentPageIndex() <> $oPresentation.getPageCount() - 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 15, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_GOTO_NEXT_EFFECT
			$oPresentation.gotoNextEffect()

		Case $LOD_SLIDESHOW_PRES_COMMAND_GOTO_NEXT_SLIDE
			$oPresentation.gotoNextPage()

		Case $LOD_SLIDESHOW_PRES_COMMAND_GOTO_PREV_EFFECT
			$oPresentation.gotoPreviousEffect()

		Case $LOD_SLIDESHOW_PRES_COMMAND_GOTO_PREV_SLIDE
			$oPresentation.gotoPreviousPage()

		Case $LOD_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE
			If Not IsObj($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oPresentation.gotoPage($vValue)
			If ($oPresentation.getCurrentPage() <> $vValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 16, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_GOTO_PAGE_BY_INDEX
			If Not __LO_IntIsBetween($vValue, 0, $oPresentation.getPageCount() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oPresentation.gotoPageIndex($vValue)
			If ($oPresentation.getCurrentPageIndex() <> $vValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 17, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_GOTO_PAGE_BY_NAME
			If Not IsString($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
			If Not $oDoc.Links.getByName("Page").Links.hasByName($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oPresentation.gotoBookmark($vValue)
			If ($oPresentation.getCurrentPage.LinkDisplayName() <> $vValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 18, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_PAUSE
			$oPresentation.pause()
			If Not $oPresentation.isPaused() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 19, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_RESUME
			$oPresentation.resume()
			If $oPresentation.isPaused() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 20, 0)

		Case $LOD_SLIDESHOW_PRES_COMMAND_STOP_SOUND
			$oPresentation.stopSound()
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, $vReturn)
EndFunc   ;==>_LODraw_PageshowPresentationControl

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowsCustomGetNames
; Description ...: Retrieve an array of Custom Pageshow names available in the document.
; Syntax ........: _LODraw_PageshowsCustomGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array containing all Custom Pageshow names. @Extended is set to the number of names returned.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Custom Presentations Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LODraw_PageshowCustomCreate, _LODraw_PageshowCustomDelete, _LODraw_PageshowCustomModify, _LODraw_PageshowCustomSetName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowsCustomGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asCustomPageShows[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asCustomPageShows = $oDoc.CustomPresentations.ElementNames()
	If Not IsArray($asCustomPageShows) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asCustomPageShows), $asCustomPageShows)
EndFunc   ;==>_LODraw_PageshowsCustomGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowSettingsMode
; Description ...: Set or Retrieve the Pageshow's play mode settings.
; Syntax ........: _LODraw_PageshowSettingsMode(ByRef $oDoc[, $iPresMode = Null[, $iRepeatPause = Null[, $bShowLogo = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iPresMode           - [optional] (0-2) Default is Null. The mode the presentation is displayed in. See Constants, $LOD_SLIDESHOW_VIEW_MODE_* as defined in LibreOfficeDraw_Constants.au3.
;                  $iRepeatPause        - [optional] (0-86399) Default is Null. If $iPresMode is set to $LOD_SLIDESHOW_VIEW_MODE_LOOP, the amount of seconds before the presentation is played again.
;                  $bShowLogo           - [optional] Default is Null. If True, the LibreOffice logo is displayed during the pause.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPresMode not an Integer, less than 0 or greater than 2. See Constants, $LOD_SLIDESHOW_VIEW_MODE_* as defined in LibreOfficeDraw_Constants.au3.
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
; Related .......: _LODraw_PageshowActiveSettings, _LODraw_PageshowPresentationControl, _LODraw_PageshowSettingsOptions, _LODraw_PageshowSettingsRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowSettingsMode(ByRef $oDoc, $iPresMode = Null, $iRepeatPause = Null, $bShowLogo = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurrMode
	Local $oPresentation
	Local $avPageShow[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPresentation = $oDoc.Presentation()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iPresMode, $iRepeatPause, $bShowLogo) Then
		If ($oPresentation.IsFullScreen() = True) And ($oPresentation.IsEndless() = False) Then
			$iCurrMode = $LOD_SLIDESHOW_VIEW_MODE_FULL_SCREEN

		ElseIf ($oPresentation.IsFullScreen() = False) And ($oPresentation.IsEndless() = False) Then
			$iCurrMode = $LOD_SLIDESHOW_VIEW_MODE_IN_WINDOW

		ElseIf ($oPresentation.IsEndless() = True) Then
			$iCurrMode = $LOD_SLIDESHOW_VIEW_MODE_LOOP

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Failed to identify current mode.
		EndIf

		__LO_ArrayFill($avPageShow, $iCurrMode, $oPresentation.Pause(), $oPresentation.IsShowLogo())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPageShow)
	EndIf

	If ($iPresMode <> Null) Then
		If Not __LO_IntIsBetween($iPresMode, $LOD_SLIDESHOW_VIEW_MODE_FULL_SCREEN, $LOD_SLIDESHOW_VIEW_MODE_LOOP) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		Switch $iPresMode
			Case $LOD_SLIDESHOW_VIEW_MODE_FULL_SCREEN
				$oPresentation.IsFullScreen = True
				$oPresentation.IsEndless = False
				$iError = ($oPresentation.IsFullScreen() = True) ? ($iError) : (BitOR($iError, 1))
				$iError = ($oPresentation.IsEndless = False) ? ($iError) : (BitOR($iError, 1))

			Case $LOD_SLIDESHOW_VIEW_MODE_IN_WINDOW
				$oPresentation.IsFullScreen = False
				$oPresentation.IsEndless = False
				$iError = ($oPresentation.IsFullScreen() = False) ? ($iError) : (BitOR($iError, 1))
				$iError = ($oPresentation.IsEndless = False) ? ($iError) : (BitOR($iError, 1))

			Case $LOD_SLIDESHOW_VIEW_MODE_LOOP
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
EndFunc   ;==>_LODraw_PageshowSettingsMode

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowSettingsOptions
; Description ...: Set or Retrieve the Pageshow's play options settings.
; Syntax ........: _LODraw_PageshowSettingsOptions(ByRef $oDoc[, $bDisableAutoPages = Null[, $bChangePageByClick = Null[, $bMouseVisible = Null[, $bMouseAsPen = Null[, $bPlayAnimatedFiles = Null[, $bKeepOnTop = Null]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $bDisableAutoPages  - [optional] Default is Null. If True, pages will not transition to the next page automatically (overriding individual page settings).
;                  $bChangePageByClick - [optional] Default is Null. If True, pages will transition when the mouse is clicked.
;                  $bMouseVisible       - [optional] Default is Null. If True, the mouse is visible in the presentation.
;                  $bMouseAsPen         - [optional] Default is Null. If True, the mouse can be used as a pen to draw on pages.
;                  $bPlayAnimatedFiles  - [optional] Default is Null. If True, animated files (such as GIFs) will be played.
;                  $bKeepOnTop          - [optional] Default is Null. If True, the presentation will be always kept on top of other programs.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bDisableAutoPages not a Boolean.
;                  @Error: 1, @Extended: 3 = $bChangePageByClick not a Boolean.
;                  @Error: 1, @Extended: 4 = $bMouseVisible not a Boolean.
;                  @Error: 1, @Extended: 5 = $bMouseAsPen not a Boolean.
;                  @Error: 1, @Extended: 6 = $bPlayAnimatedFiles not a Boolean.
;                  @Error: 1, @Extended: 7 = $bKeepOnTop not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Presentation Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bDisableAutoPages
;                  |                               2 = Error setting $bChangePageByClick
;                  |                               4 = Error setting $bMouseVisible
;                  |                               8 = Error setting $bMouseAsPen
;                  |                               16 = Error setting $bPlayAnimatedFiles
;                  |                               32 = Error setting $bKeepOnTop
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LODraw_PageshowActiveSettings, _LODraw_PageshowPresentationControl, _LODraw_PageshowSettingsMode, _LODraw_PageshowSettingsRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowSettingsOptions(ByRef $oDoc, $bDisableAutoPages = Null, $bChangePageByClick = Null, $bMouseVisible = Null, $bMouseAsPen = Null, $bPlayAnimatedFiles = Null, $bKeepOnTop = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oPresentation
	Local $avPageShow[6]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPresentation = $oDoc.Presentation()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bDisableAutoPages, $bChangePageByClick, $bMouseVisible, $bMouseAsPen, $bPlayAnimatedFiles, $bKeepOnTop) Then
		__LO_ArrayFill($avPageShow, $oPresentation.IsAutomatic(), $oPresentation.IsTransitionOnClick(), $oPresentation.IsMouseVisible(), _
				$oPresentation.UsePen(), $oPresentation.AllowAnimations(), $oPresentation.IsAlwaysOnTop())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPageShow)
	EndIf

	If ($bDisableAutoPages <> Null) Then
		If Not IsBool($bDisableAutoPages) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oPresentation.IsAutomatic = $bDisableAutoPages
		$iError = ($oPresentation.IsAutomatic() = $bDisableAutoPages) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bChangePageByClick <> Null) Then
		If Not IsBool($bChangePageByClick) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPresentation.IsTransitionOnClick = $bChangePageByClick
		$iError = ($oPresentation.IsTransitionOnClick() = $bChangePageByClick) ? ($iError) : (BitOR($iError, 2))
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
EndFunc   ;==>_LODraw_PageshowSettingsOptions

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowSettingsRange
; Description ...: Set or Retrieve the Pageshow's play Range settings.
; Syntax ........: _LODraw_PageshowSettingsRange(ByRef $oDoc[, $iRange = Null[, $sValue = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iRange              - [optional] (0-2) Default is Null. The Range of pages that will be shown when the Presentation is started. See Constants, $LOD_SLIDESHOW_RANGE_* as defined in LibreOfficeDraw_Constants.au3.
;                  $sValue              - [optional] Default is Null. The "From" page or Custom Page Show name. See remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iRange not an Integer, less than 0 or greater than 2. See Constants, $LOD_SLIDESHOW_RANGE_* as defined in LibreOfficeDraw_Constants.au3.
;                  @Error: 1, @Extended: 3 = $sValue not a String.
;                  @Error: 1, @Extended: 4 = Range set to $LOD_SLIDESHOW_RANGE_FROM, and the Page name called in $sValue does not exist.
;                  @Error: 1, @Extended: 5 = Range set to $LOD_SLIDESHOW_RANGE_CUSTOM, and the Custom Pageshow name called in $sValue does not exist.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Presentation Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Start From page value.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Custom Pageshow name.
;                  @Error: 3, @Extended: 4 = Failed to identify current range.
;                  @Error: 3, @Extended: 5 = $iRange is called with other than $LOD_SLIDESHOW_RANGE_ALL, and $sValue is not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRange
;                  |                               2 = Error setting $sValue
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you call $iRange with any other value than $LOD_SLIDESHOW_RANGE_ALL, $sValue must be called with an appropriate name, either a Page name to start from, or a Custom Pageshow name.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  If there are two pages with the same name, and one is set to the "From Page" property, there is no guarantee which page will be the one used.
; Related .......: _LODraw_PageshowActiveSettings, _LODraw_PageshowPresentationControl, _LODraw_PageshowSettingsMode, _LODraw_PageshowSettingsOptions
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowSettingsRange(ByRef $oDoc, $iRange = Null, $sValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurrRange
	Local $oPresentation
	Local $avPageShow[2]
	Local $sCurrValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oPresentation = $oDoc.Presentation()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oPresentation.IsShowAll() = True) Then
		$iCurrRange = $LOD_SLIDESHOW_RANGE_ALL
		$sCurrValue = ""

	ElseIf ($oPresentation.FirstPage() <> "") Then
		$iCurrRange = $LOD_SLIDESHOW_RANGE_FROM
		$sCurrValue = $oPresentation.FirstPage()
		If Not IsString($sCurrValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$sCurrValue = $oDoc.DrawPages.getByName($sCurrValue).LinkDisplayName()    ; Get Link Display Name as it is more reliable?
		If Not IsString($sCurrValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	ElseIf ($oPresentation.CustomShow() <> "") Then
		$iCurrRange = $LOD_SLIDESHOW_RANGE_CUSTOM
		$sCurrValue = $oPresentation.CustomShow()
		If Not IsString($sCurrValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)     ; Failed to identify current range.
	EndIf

	If __LO_VarsAreNull($iRange, $sValue) Then
		__LO_ArrayFill($avPageShow, $iCurrRange, $sCurrValue)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPageShow)
	EndIf

	If ($iRange <> Null) Then
		If Not __LO_IntIsBetween($iRange, $LOD_SLIDESHOW_RANGE_ALL, $LOD_SLIDESHOW_RANGE_CUSTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		Switch $iRange
			Case $LOD_SLIDESHOW_RANGE_ALL
				$oPresentation.IsShowAll = True
				$oPresentation.FirstPage = ""
				$oPresentation.CustomShow = ""

				$iError = ($oPresentation.IsShowAll() = True) ? ($iError) : (BitOR($iError, 1))
				$iError = ($oPresentation.FirstPage() = "") ? ($iError) : (BitOR($iError, 1))
				$iError = ($oPresentation.CustomShow() = "") ? ($iError) : (BitOR($iError, 1))

			Case $LOD_SLIDESHOW_RANGE_FROM
				If ($iCurrRange <> $iRange) Then
					If ($sValue = Null) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Value not called

					$oPresentation.IsShowAll = False
					$oPresentation.CustomShow = ""

					$iError = ($oPresentation.IsShowAll() = False) ? ($iError) : (BitOR($iError, 1))
					$iError = ($oPresentation.CustomShow() = "") ? ($iError) : (BitOR($iError, 1))
				EndIf

			Case $LOD_SLIDESHOW_RANGE_CUSTOM
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
			Case $LOD_SLIDESHOW_RANGE_FROM
				If Not $oDoc.Links.getByName("Page").Links.hasByName($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

				$sValue = $oDoc.Links.getByName("Page").Links.getByName($sValue).Name()    ; Overwrite value (LinkDisplayName) with Page's name, as that is what L.O. uses.

				$oPresentation.FirstPage = $sValue
				$iError = ($oPresentation.FirstPage() = $sValue) ? ($iError) : (BitOR($iError, 2))

			Case $LOD_SLIDESHOW_RANGE_CUSTOM
				If Not $oDoc.CustomPresentations.hasByName($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

				$oPresentation.CustomShow = $sValue
				$iError = ($oPresentation.CustomShow() = $sValue) ? ($iError) : (BitOR($iError, 2))
		EndSwitch
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageshowSettingsRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowStart
; Description ...: Begins a presentation.
; Syntax ........: _LODraw_PageshowStart(ByRef $oDoc[, $bRehearse = False[, $sStartPage = ""[, $sCustomShow = ""]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $bRehearse           - [optional] Default is False. If True, starts the presentation from the beginning and shows a rehearsal timer to the user.
;                  $sStartPage         - [optional] Default is "". The Page's name to begin this presentation from.
;                  $sCustomShow         - [optional] Default is "". The Custom Pageshow's name to play for this presentation.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. The Presentation was started successfully.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bRehearse not a Boolean.
;                  @Error: 1, @Extended: 3 = $sStartPage not a String.
;                  @Error: 1, @Extended: 4 = $sCustomShow not a String.
;                  @Error: 1, @Extended: 5 = Page name called in $sStartPage not found.
;                  @Error: 1, @Extended: 6 = Custom Pageshow name called in $sCustomShow not found.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create FirstPage property.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Presentation Object.
;                  @Error: 3, @Extended: 2 = There is already a presentation running.
;                  @Error: 3, @Extended: 3 = Failed to start the presentation.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bRehearse is called with True, both $sStartPage and $sCustomShow will be ignored.
;                  If both $sStartPage and $sCustomShow are called with a parameter, $sCustomShow will be ignored.
; Related .......: _LODraw_PageshowIsRunning, _LODraw_PageshowStop
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowStart(ByRef $oDoc, $bRehearse = False, $sStartPage = "", $sCustomShow = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atProperties[0]
	Local $oPresentation
	Local $bBackupShowAll
	Local $sBackupFirst, $sBackupCustom

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bRehearse) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sStartPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sCustomShow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oPresentation = $oDoc.Presentation()
	If Not IsObj($oPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oPresentation.IsRunning() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; A Pageshow is already active.

	If $bRehearse Then
		$oPresentation.rehearseTimings()

	ElseIf ($sStartPage <> "") Then
		If Not $oDoc.Links.getByName("Page").Links.hasByName($sStartPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		ReDim $atProperties[1]
		$atProperties[0] = __LO_SetPropertyValue("FirstPage", $sStartPage)
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
EndFunc   ;==>_LODraw_PageshowStart

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageshowStop
; Description ...: Stop the presently playing presentation.
; Syntax ........: _LODraw_PageshowStop(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
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
; Related .......: _LODraw_PageshowIsRunning, _LODraw_PageshowStart
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageshowStop(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If $oDoc.Presentation.IsRunning() Then
		$oDoc.Presentation.end()
		If $oDoc.Presentation.IsRunning() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Failed to stop Pageshow.
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LODraw_PageshowStop

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageSoundsGetNames
; Description ...: Retrieve an array of Sound files that are included with LibreOffice Draw.
; Syntax ........: _LODraw_PageSoundsGetNames()
; Parameters ....: None
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning array of included Draw Sound files. @Extended will be set to number of results.
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
; Related .......: _LODraw_PageTransition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageSoundsGetNames()
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
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
EndFunc   ;==>_LODraw_PageSoundsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageTransition
; Description ...: Set or Retrieve a Page's Transition properties.
; Syntax ........: _LODraw_PageTransition(ByRef $oPage[, $iTransition = Null[, $nDuration = Null[, $sSound = Null[, $bLoopSound = Null[, $nPageAdvance = Null]]]]])
; Parameters ....: $oPage              - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
;                  $iTransition         - [optional] (0-78) Default is Null. The Transition effect. See Constants, $LOD_PAGE_TRANSITION_* as defined in LibreOfficeDraw_Constants.au3.
;                  $nDuration           - [optional] (0-1000) Default is Null. The duration of the page's transition effect, in seconds. L.O. 6.1+. See remarks.
;                  $sSound              - [optional] Default is Null. The path to the sound to play during page transition. See remarks.
;                  $bLoopSound          - [optional] Default is Null. If True, the sound is repeated.
;                  $nPageAdvance       - [optional] (-1-1000) Default is Null. The number of seconds before automatically advance the page. Call with -1 to set to On Mouse Click.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oPage not an Object.
;                  @Error: 1, @Extended: 2 = $iTransition not an Integer, less than 0 or greater than 78. See Constants, $LOD_PAGE_TRANSITION_* as defined in LibreOfficeDraw_Constants.au3.
;                  @Error: 1, @Extended: 3 = $nDuration not a Number, less than 0 or greater than 1000.
;                  @Error: 1, @Extended: 4 = $sSound not a String.
;                  @Error: 1, @Extended: 5 = File called in $sSound does not exist.
;                  @Error: 1, @Extended: 6 = $bLoopSound not a Boolean.
;                  @Error: 1, @Extended: 7 = $nPageAdvance not a Number, less than -1 or greater than 1000.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve the current Transition type.
;                  @Error: 3, @Extended: 2 = Failed to retrieve current Duration.
;                  @Error: 3, @Extended: 3 = Failed to retrieve current Sound value.
;                  @Error: 3, @Extended: 4 = Failed to retrieve current Page advance value.
;                  @Error: 3, @Extended: 5 = Failed to set Transition type.
;                  @Error: 3, @Extended: 6 = Failed to convert Sound path to LibreOffice path.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTransition
;                  |                               2 = Error setting $nDuration
;                  |                               4 = Error setting $sSound
;                  |                               8 = Error setting $bLoopSound
;                  |                               16 = Error setting $nPageAdvance
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Previous to LibreOffice 6.1 $nDuration was simply a three option selection of Slow, Medium, and Fast. To make both these work with this UDF the following method has been adopted:
;                  Previous to LibreOffice 6.1, if the Value called in $nDuration is from 0 to 1.99, the speed is set to Fast, if $nDuration is called with 2, Medium speed is set, and any value beyond 2.01 is considered Slow. This matches LibreOffice's internal behaviour.
;                  When retrieving current property values previous to LibreOffice 6.1, if Speed is set to Fast, 1 is returned for $nDuration. If Speed is set to Medium, 2 is returned. And if Speed is set to Slow, 3 is returned.
;                  $sSound can be called with an empty string to indicate that no sound should be played.
;                  If $sSound is called with the string "stop", this equals "Stop Previous Sound" in the UI.
;                  Otherwise call $sSound with a valid path to a sound file. See _LODraw_PageSoundsGetNames, to obtain a list of sound files included with Draw.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LODraw_PageSoundsGetNames, _LODraw_PageLayout
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageTransition(ByRef $oPage, $iTransition = Null, $nDuration = Null, $sSound = Null, $bLoopSound = Null, $nPageAdvance = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__LOD_CONST_CHANGE_MANUAL = 0, $__LOD_CONST_CHANGE_AUTO = 1 ;,  $__LOD_CONST_CHANGE_SEMI_MANUAL = 2 ; com.sun.star.presentation.DrawPage:Change
	Local Const $__LOD_CONST_SPEED_SLOW = 0, $__LOD_CONST_SPEED_MEDIUM = 1, $__LOD_CONST_SPEED_FAST = 2    ; com.sun.star.presentation:AnimationSpeed
	Local $iError = 0, $iCurrTransition, $nCurrDuration, $nCurrPageAdvance
	Local $sCurrSound
	Local $avTransition[5]

	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransition, $nDuration, $sSound, $bLoopSound, $nPageAdvance) Then
		$iCurrTransition = __LODraw_Transition($oPage)
		If Not IsInt($iCurrTransition) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If __LO_VersionCheck(6.1) Then
			$nCurrDuration = $oPage.TransitionDuration()

		Else
			Switch $oPage.Speed()
				Case $__LOD_CONST_SPEED_FAST ; 0 - 1.99 ; Matches L.O. Behaviour.
					$nCurrDuration = 1

				Case $__LOD_CONST_SPEED_MEDIUM ; 2
					$nCurrDuration = 2

				Case Else ; $__LOD_CONST_SPEED_SLOW
					$nCurrDuration = 3
			EndSwitch
		EndIf

		If Not IsNumber($nCurrDuration) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$sCurrSound = $oPage.Sound()
		If IsBool($sCurrSound) And ($sCurrSound = True) Then $sCurrSound = "stop"
		If Not IsString($sCurrSound) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$sCurrSound = _LO_PathConvert($sCurrSound, $LO_PATHCONV_PCPATH_RETURN)

		If ($oPage.Change() <> $__LOD_CONST_CHANGE_AUTO) Then
			$nCurrPageAdvance = -1

		Else
			$nCurrPageAdvance = $oPage.HighResDuration()
		EndIf

		If Not IsNumber($nCurrPageAdvance) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		__LO_ArrayFill($avTransition, $iCurrTransition, $nCurrDuration, $sCurrSound, $oPage.LoopSound(), $nCurrPageAdvance)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTransition)
	EndIf

	If ($iTransition <> Null) Then
		If Not __LO_IntIsBetween($iTransition, $LOD_PAGE_TRANSITION_3D_VENETIAN_VERT, $LOD_PAGE_TRANSITION_WIPE_TOP_TO_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		__LODraw_Transition($oPage, $iTransition)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		$iError = (__LODraw_Transition($oPage) = $iTransition) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($nDuration <> Null) Then
		If Not __LO_NumIsBetween($nDuration, 0, 1000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If __LO_VersionCheck(6.1) Then
			$oPage.TransitionDuration = $nDuration
			$iError = ($oPage.TransitionDuration() = $nDuration) ? ($iError) : (BitOR($iError, 2))

		Else
			Switch $nDuration
				Case 0 - 1.99 ; Matches L.O. Behaviour.
					$oPage.Speed = $__LOD_CONST_SPEED_FAST
					$iError = ($oPage.Speed() = $__LOD_CONST_SPEED_FAST) ? ($iError) : (BitOR($iError, 2))

				Case 2
					$oPage.Speed = $__LOD_CONST_SPEED_MEDIUM
					$iError = ($oPage.Speed() = $__LOD_CONST_SPEED_MEDIUM) ? ($iError) : (BitOR($iError, 2))

				Case Else
					$oPage.Speed = $__LOD_CONST_SPEED_SLOW
					$iError = ($oPage.Speed() = $__LOD_CONST_SPEED_SLOW) ? ($iError) : (BitOR($iError, 2))
			EndSwitch
		EndIf
	EndIf

	If ($sSound <> Null) Then
		If Not IsString($sSound) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If ($sSound = "") Then
			$oPage.Sound = $sSound
			$iError = ($oPage.Sound() = $sSound) ? ($iError) : (BitOR($iError, 4))

		ElseIf ($sSound = "stop") Then
			$oPage.Sound = True
			$iError = ($oPage.Sound() = True) ? ($iError) : (BitOR($iError, 4))

		Else
			If Not FileExists($sSound) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$sSound = _LO_PathConvert($sSound, $LO_PATHCONV_OFFICE_RETURN)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

			$oPage.Sound = $sSound
			$iError = ($oPage.Sound() = $sSound) ? ($iError) : (BitOR($iError, 4))
		EndIf
	EndIf

	If ($bLoopSound <> Null) Then
		If Not IsBool($bLoopSound) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPage.LoopSound = $bLoopSound
		$iError = ($oPage.LoopSound() = $bLoopSound) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($nPageAdvance <> Null) Then
		If Not __LO_NumIsBetween($nPageAdvance, -1, 1000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		If ($nPageAdvance = -1) Then
			$oPage.Change = $__LOD_CONST_CHANGE_MANUAL
			$iError = ($oPage.Change() = $__LOD_CONST_CHANGE_MANUAL) ? ($iError) : (BitOR($iError, 16))

		Else
			If ($oPage.Change() <> $__LOD_CONST_CHANGE_AUTO) Then $oPage.Change = $__LOD_CONST_CHANGE_AUTO
			$oPage.Duration = $nPageAdvance
			$oPage.HighResDuration = $nPageAdvance
			$iError = (($oPage.Duration() = $nPageAdvance) And ($oPage.HighResDuration() = $nPageAdvance)) ? ($iError) : (BitOR($iError, 16))
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageTransition
