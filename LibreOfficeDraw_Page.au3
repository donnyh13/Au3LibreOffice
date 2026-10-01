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
; _LODraw_PageFormat
; _LODraw_PageGetObjByIndex
; _LODraw_PageGetObjByName
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
; _LODraw_PageMastersGetCount
; _LODraw_PageMastersGetNames
; _LODraw_PageMove
; _LODraw_PageName
; _LODraw_PagesGetCount
; _LODraw_PagesGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageAdd
; Description ...: Add a page to a Draw Document.
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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
;                  $iPage               - The page to delete. 0 based.
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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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

	$bExists = $oDoc.Links.getByName("Slide").Links.hasByName($sName)
	If Not IsBool($bExists) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bExists)
EndFunc   ;==>_LODraw_PageExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageFormat
; Description ...: Set or Retrieve the page format settings.
; Syntax ........: _LODraw_PageFormat(ByRef $oPage[, $iWidth = Null[, $iHeight = Null[, $iOrientation = Null]]])
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
; Related .......: _LO_UnitConvert, _LODraw_PageMargins, _LODraw_PageMasterFormat
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
;                  $iPage               - The page to retrieve. 0 based.
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
	If Not $oDoc.Links.getByName("Slide").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPage = $oDoc.Links.getByName("Slide").Links.getByName($sName)
	If Not IsObj($oPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPage)
EndFunc   ;==>_LODraw_PageGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LODraw_PageMargins
; Description ...: Set or Retrieve the page margin settings.
; Syntax ........: _LODraw_PageMargins(ByRef $oPage[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
; Related .......: _LO_UnitConvert, _LODraw_PageFormat, _LODraw_PageMasterMargins
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
; Description ...: Add a master page to a Draw Document.
; Syntax ........: _LODraw_PageMasterAdd(ByRef $oDoc[, $iPos = Null[, $sName = ""]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LODraw_DocOpen, _LODraw_DocConnect, or _LODraw_DocCreate function.
;                  $iPos                - [optional] Default is Null. The position to insert the new master page in the collection of pages. 0 Based.
;                  $sName               - [optional] Default is "". The unique name of the Master Page. If called with an empty string, LibreOffice automatically names it.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning new page's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPos not an Integer, less than 0 or greater than number of master pages.
;                  @Error: 1, @Extended: 3 = $sName not a String.
;                  @Error: 1, @Extended: 4 = Name called in $sName already exists.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to create a master page.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $iPos is called with Null, the new master page is inserted at the end.
;                  Call $iPos with the last master page index to insert the master page at the end. Call $iPos with 0 to insert the new master page at the beginning.
; Related .......: _LODraw_PageMasterDeleteByIndex, _LODraw_PageMasterDeleteByObj, _LODraw_PageAdd, _LODraw_PageMasterExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LODraw_PageMasterAdd(ByRef $oDoc, $iPos = Null, $sName = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LODraw_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMPage

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($iPos = Null) Then $iPos = $oDoc.MasterPages.getCount()
	If Not __LO_IntIsBetween($iPos, 0, $oDoc.MasterPages.getCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sName <> "") And _LODraw_PageMasterExists($oDoc, $sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oMPage = $oDoc.MasterPages.insertNewByIndex($iPos)
	If Not IsObj($oMPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

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
	EndIf

	$oBackground.FillStyle = $LOD_AREA_FILL_STYLE_SOLID
	$oBackground.FillColor = $iColor

	$oMaster.Background = $oBackground

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

	$oMaster.Background = $oBackground

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
	EndIf

	$oBackground.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oBackground.FillTransparence = $iTransparency

	$oMaster.Background = $oBackground

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

	$oMaster.Background = $oBackground

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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
; Related .......: _LO_UnitConvert, _LODraw_PageMasterMargins, _LODraw_PageFormat
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
; Description ...: Set or Retrieve the master page margin settings.
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
; Related .......: _LO_UnitConvert, _LODraw_PageMasterFormat, _LODraw_PageMargins
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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
; Parameters ....: $oPage               - A Page object returned by a previous _LODraw_PageAdd, _LODraw_PageGetObjByIndex, _LODraw_PageGetObjByName, or _LODraw_PageCopy function.
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
	If $oDoc.Links.getByName("Slide").Links.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oPage.Name = $sName
	$iError = ($oPage.LinkDisplayName() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LODraw_PageName

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

	$asPages = $oDoc.Links.getByName("Slide").Links.getElementNames()
	If Not IsArray($asPages) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asPages), $asPages)
EndFunc   ;==>_LODraw_PagesGetNames
