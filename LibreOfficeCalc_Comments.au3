#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Constants.au3"
#include "LibreOfficeCalc_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Removing, etc. L.O. Calc document Cell Comments.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_CommentAdd
; _LOCalc_CommentAreaColor
; _LOCalc_CommentAreaFillStyle
; _LOCalc_CommentAreaGradient
; _LOCalc_CommentAreaGradientMulticolor
; _LOCalc_CommentAreaShadow
; _LOCalc_CommentAreaTransparency
; _LOCalc_CommentAreaTransparencyGradient
; _LOCalc_CommentAreaTransparencyGradientMulti
; _LOCalc_CommentCallout
; _LOCalc_CommentCreateTextCursor
; _LOCalc_CommentDelete
; _LOCalc_CommentGetCell
; _LOCalc_CommentGetLastEdit
; _LOCalc_CommentGetObjByCell
; _LOCalc_CommentGetObjByIndex
; _LOCalc_CommentLineArrowStyles
; _LOCalc_CommentLineProperties
; _LOCalc_CommentPosition
; _LOCalc_CommentRotate
; _LOCalc_CommentsGetCount
; _LOCalc_CommentsGetList
; _LOCalc_CommentSize
; _LOCalc_CommentText
; _LOCalc_CommentTextAnchor
; _LOCalc_CommentTextAnimation
; _LOCalc_CommentTextColumns
; _LOCalc_CommentTextSettings
; _LOCalc_CommentVisible
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentAdd
; Description ...: Add a comment to a cell.
; Syntax ........: _LOCalc_CommentAdd(ByRef $oCell, $sText)
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
;                  $sText               - a string value. The initial text of the Comment. Cannot be empty.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell not a Cell Object.
;                  @Error 1 @Extended 3 Return 0 = $sText not a String or string is empty.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotations Object.
;                  @Error 3 @Extended 2 Return 0 = Called Cell already contains a Comment.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Cell Address.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve new Comment Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning newly inserted Comment's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_CommentDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentAdd(ByRef $oCell, $sText)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tAddress
	Local $oAnnotations, $oAnnotation

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.SupportsService("com.sun.star.sheet.SheetCell") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sText) Or ($sText = "") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oAnnotations = $oCell.Spreadsheet.Annotations()
	If Not IsObj($oAnnotations) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oAnnotation = __LOCalc_CommentGetObjByCell($oCell)
	If IsObj($oAnnotation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tAddress = $oCell.CellAddress()
	If Not IsObj($tAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oAnnotations.insertNew($tAddress, $sText)

	$oAnnotation = __LOCalc_CommentGetObjByCell($oCell)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oAnnotation)
EndFunc   ;==>_LOCalc_CommentAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentAreaColor
; Description ...: Set or Retrieve the Comment's background color.
; Syntax ........: _LOCalc_CommentAreaColor(ByRef $oComment[, $iColor = Null])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The color for the background of the comment, set in Long Color Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current setting as an Integer value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current setting.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentAreaColor(ByRef $oComment, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iColor = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oAnnotationShape.FillColor())

	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oAnnotationShape.FillStyle = $LOC_AREA_FILL_STYLE_SOLID
	$oAnnotationShape.FillColor = $iColor

	If ($oAnnotationShape.FillColor() <> $iColor) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CommentAreaColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentAreaFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOCalc_CommentAreaFillStyle(ByRef $oComment)
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Fill Style.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning current background fill style. Return will be one of the constants $LOC_AREA_FILL_STYLE_* as defined in LibreOfficeCalc_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentAreaFillStyle(ByRef $oComment)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oComment.AnnotationShape.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOCalc_CommentAreaFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentAreaGradient
; Description ...: Modify or retrieve the settings for Comment Background color Gradient.
; Syntax ........: _LOCalc_CommentAreaGradient(ByRef $oComment[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See remarks. See constants, $LOC_GRAD_NAME_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient type to apply. See Constants, $LOC_GRAD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
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
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sGradientName not a String.
;                  @Error 1 @Extended 3 Return 0 = $iType not an Integer, less than -1, or greater than 5. See Constants, $LOC_GRAD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
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
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Annotation Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving "FillGradient" Object.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving Color Stop Array for "From" color
;                  @Error 3 @Extended 4 Return 0 = Error retrieving Color Stop Array for "To" color
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
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentAreaGradient(ByRef $oComment, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $nRed, $nGreen, $nBlue
	Local $atColorStop
	Local $avGradient[11]

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tStyleGradient = $oAnnotationShape.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oAnnotationShape.FillGradientName(), $tStyleGradient.Style(), _
				$oAnnotationShape.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oAnnotationShape.FillStyle() <> $LOC_AREA_FILL_STYLE_GRADIENT) Then $oAnnotationShape.FillStyle = $LOC_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oAnnotationShape.FillGradientName = $sGradientName
		$iError = ($oAnnotationShape.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))

		$tStyleGradient = $oAnnotationShape.FillGradient()
		If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOC_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oAnnotationShape.FillStyle = $LOC_AREA_FILL_STYLE_OFF
			$oAnnotationShape.FillGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOC_GRAD_TYPE_LINEAR, $LOC_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oAnnotationShape.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oAnnotationShape.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
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

	$oAnnotationShape.FillGradient = $tStyleGradient

	; Error checking
	$iError = ($iType = Null) ? ($iError) : (($oAnnotationShape.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oAnnotationShape.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oAnnotationShape.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oAnnotationShape.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oAnnotationShape.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = ($iFromColor = Null) ? ($iError) : (($oAnnotationShape.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = ($iToColor = Null) ? ($iError) : (($oAnnotationShape.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = ($iFromIntense = Null) ? ($iError) : (($oAnnotationShape.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = ($iToIntense = Null) ? ($iError) : (($oAnnotationShape.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentAreaGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentAreaGradientMulticolor
; Description ...: Set or Retrieve a Comment's Multicolor Gradient settings. See remarks.
; Syntax ........: _LOCalc_CommentAreaGradientMulticolor(ByRef $oComment[, $avColorStops = Null])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 4 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 5 Return ? = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve AnnotationShape Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve FillGradient Struct.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve StopColor Struct.
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
; Related .......: _LOCalc_GradientMulticolorAdd, _LOCalc_GradientMulticolorDelete, _LOCalc_GradientMulticolorModify, _LOCalc_CommentAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentAreaGradientMulticolor(ByRef $oComment, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tStyleGradient = $oAnnotationShape.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($avColorStops) Then
		$atColorStops = $tStyleGradient.ColorStops()
		If Not IsArray($atColorStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		ReDim $avNewColorStops[UBound($atColorStops)][2]

		For $i = 0 To UBound($atColorStops) - 1
			$avNewColorStops[$i][0] = $atColorStops[$i].StopOffset()
			$tStopColor = $atColorStops[$i].StopColor()
			If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			$avNewColorStops[$i][1] = Int(BitShift(($tStopColor.Red() * 255), -16) + BitShift(($tStopColor.Green() * 255), -8) + ($tStopColor.Blue() * 255)) ; RGB to Long
			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
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
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tStopColor.Red = (BitAND(BitShift($avColorStops[$i][1], 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($avColorStops[$i][1], 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($avColorStops[$i][1], 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oAnnotationShape.FillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oAnnotationShape.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentAreaGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentAreaShadow
; Description ...: Set or Retrieve the shadow settings for a Comment.
; Syntax ........: _LOCalc_CommentAreaShadow(ByRef $oComment[, $bShadow = Null[, $iColor = Null[, $iDistance = Null[, $iTransparency = Null[, $iBlur = Null[, $iLocation = Null]]]]]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $bShadow             - [optional] a boolean value. Default is Null. If True, a Shadow is present for the Comment.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Shadow color, set in Long Integer format, can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iDistance           - [optional] an integer value. Default is Null. The distance of the Shadow from the Comment box, set in Micrometers.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The percentage of Shadow transparency. 100% means completely transparent.
;                  $iBlur               - [optional] an integer value (0-150). Default is Null. The amount of blur applied to the Shadow, set in Printer's Points.
;                  $iLocation           - [optional] an integer value (0-8). Default is Null. The Location of the Shadow, must be one of the Constants, $LOC_COMMENT_SHADOW_* as defined in LibreOfficeCalc_Constants.au3..
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bShadow not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $iColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 4 Return 0 = $iDistance not an Integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iTransparency not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iBlur not an Integer, less than 0, or greater than 150 Printer's Points (0-5292 Micrometers).
;                  @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, less than 0, or greater than 8. See Constants, $LOC_COMMENT_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotation Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Distance and Location Values.
;                  @Error 3 @Extended 3 Return 0 = Failed to modify Distance property.
;                  @Error 3 @Extended 4 Return 0 = Failed to modify Location property.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bShadow
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $iDistance
;                  |                               8 = Error setting $iTransparency
;                  |                               16 = Error setting $iBlur
;                  |                               32 = Error setting $iLocation
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  LibreOffice may change the shadow distance +/- a Micrometer.
;                  Presently only location settings applying the Shadow to the bottom, right, or bottom-right corners of the Comment visually work, both in LibreOffice and using this function. Though it still can be set to the other locations.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentAreaShadow(ByRef $oComment, $bShadow = Null, $iColor = Null, $iDistance = Null, $iTransparency = Null, $iBlur = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $iError = 0, $iInternalLocation, $iInternalDistance
	Local $avShadow[6]

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bShadow, $iColor, $iDistance, $iTransparency, $iBlur, $iLocation) Then
		$iInternalDistance = __LOCalc_CommentAreaShadowModify($oAnnotationShape)
		$iInternalLocation = @extended
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		__LO_ArrayFill($avShadow, $oAnnotationShape.Shadow(), $oAnnotationShape.ShadowColor(), $iInternalDistance, $oAnnotationShape.ShadowTransparence(), _
				__LO_UnitConvert($oAnnotationShape.ShadowBlur(), $__LOCONST_CONVERT_UM_PT), $iInternalLocation)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($bShadow <> Null) Then
		If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oAnnotationShape.Shadow = $bShadow
		$iError = ($oAnnotationShape.Shadow() = $bShadow) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oAnnotationShape.ShadowColor = $iColor
		$iError = ($oAnnotationShape.ShadowColor() = $iColor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iDistance <> Null) Then
		If Not __LO_IntIsBetween($iDistance, 0, $iDistance) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		__LOCalc_CommentAreaShadowModify($oAnnotationShape, Null, $iDistance)
		If (@error = $__LO_STATUS_PROP_SETTING_ERROR) Then
			$iError = BitOR($iError, 4)

		ElseIf @error Then

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		EndIf
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oAnnotationShape.ShadowTransparence = $iTransparency
		$iError = ($oAnnotationShape.ShadowTransparence = $iTransparency) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iBlur <> Null) Then
		If Not __LO_IntIsBetween($iBlur, 0, 150) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; 0 - 5292 max Micrometers.

		$oAnnotationShape.ShadowBlur = __LO_UnitConvert($iBlur, $__LOCONST_CONVERT_PT_UM)
		$iError = ($oAnnotationShape.ShadowBlur() = __LO_UnitConvert($iBlur, $__LOCONST_CONVERT_PT_UM)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOC_COMMENT_SHADOW_TOP_LEFT, $LOC_COMMENT_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		__LOCalc_CommentAreaShadowModify($oAnnotationShape, $iLocation)
		If (@error = $__LO_STATUS_PROP_SETTING_ERROR) Then
			$iError = BitOR($iError, 32)

		ElseIf @error Then

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentAreaShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentAreaTransparency
; Description ...: Set or retrieve Transparency settings for a Comment.
; Syntax ........: _LOCalc_CommentAreaTransparency(ByRef $oComment[, $iTransparency = Null])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTransparency not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotation Shape Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iTransparency
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current setting for Transparency in integer format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentAreaTransparency(ByRef $oComment, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oAnnotationShape

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iTransparency) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oAnnotationShape.FillTransparence())

	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oAnnotationShape.FillTransparenceGradientName = "" ; Turn off Gradient if it is on, else settings wont be applied.
	$oAnnotationShape.FillTransparence = $iTransparency
	$iError = ($oAnnotationShape.FillTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentAreaTransparency

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentAreaTransparencyGradient
; Description ...: Set or retrieve the Comment transparency gradient settings.
; Syntax ........: _LOCalc_CommentAreaTransparencyGradient(ByRef $oDoc, ByRef $oComment[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient to apply. See Constants, $LOC_GRAD_TYPE_* as defined in LibreOfficeCalc_Constants.au3. Set to $LOC_GRAD_TYPE_OFF to turn Transparency Gradient off.
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
;                  @Error 1 @Extended 2 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iType Not an Integer, less than -1, or greater than 5, see constants, $LOC_GRAD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iXCenter Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 5 Return 0 = $iYCenter Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 6 Return 0 = $iAngle Not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 7 Return 0 = $iTransitionStart Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iStart Not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 9 Return 0 = $iEnd Not an Integer, less than 0, or greater than 100.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotation Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving "FillTransparenceGradient" Object.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentAreaTransparencyGradient(ByRef $oDoc, ByRef $oComment, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tGradient, $tColorStop, $tStopColor
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]
	Local $atColorStop
	Local $oAnnotationShape
	Local $fValue

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tGradient = $oAnnotationShape.FillTransparenceGradient()
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iStart, $iEnd) Then
		__LO_ArrayFill($aiTransparent, $tGradient.Style(), $tGradient.XOffset(), $tGradient.YOffset(), _
				Int($tGradient.Angle() / 10), $tGradient.Border(), __LOCalc_TransparencyGradientConvert(Null, $tGradient.StartColor()), _
				__LOCalc_TransparencyGradientConvert(Null, $tGradient.EndColor())) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOC_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oAnnotationShape.FillTransparenceGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOC_GRAD_TYPE_LINEAR, $LOC_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

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

		$tGradient.StartColor = __LOCalc_TransparencyGradientConvert($iStart)

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
		If Not __LO_IntIsBetween($iEnd, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tGradient.EndColor = __LOCalc_TransparencyGradientConvert($iEnd)

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

	If ($oAnnotationShape.FillTransparenceGradientName() = "") Then
		$sTGradName = __LOCalc_TransparencyGradientNameInsert($oDoc, $tGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		$oAnnotationShape.FillTransparenceGradientName = $sTGradName
		If ($oAnnotationShape.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
	EndIf

	$oAnnotationShape.FillTransparenceGradient = $tGradient

	$iError = ($iType = Null) ? ($iError) : (($oAnnotationShape.FillTransparenceGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iXCenter = Null) ? ($iError) : (($oAnnotationShape.FillTransparenceGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iYCenter = Null) ? ($iError) : (($oAnnotationShape.FillTransparenceGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iAngle = Null) ? ($iError) : ((Int($oAnnotationShape.FillTransparenceGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iTransitionStart = Null) ? ($iError) : (($oAnnotationShape.FillTransparenceGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($iStart = Null) ? ($iError) : (($oAnnotationShape.FillTransparenceGradient.StartColor() = __LOCalc_TransparencyGradientConvert($iStart)) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($iEnd = Null) ? ($iError) : (($oAnnotationShape.FillTransparenceGradient.EndColor() = __LOCalc_TransparencyGradientConvert($iEnd)) ? ($iError) : (BitOR($iError, 64)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentAreaTransparencyGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentAreaTransparencyGradientMulti
; Description ...: Set or Retrieve a Comment's Multi Transparency Gradient settings. See remarks.
; Syntax ........: _LOCalc_CommentAreaTransparencyGradientMulti(ByRef $oComment[, $avColorStops = Null])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $avColorStops        - [optional] an array of variants. Default is Null. A Two column array of Transparency values and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops not an Array, or does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $avColorStops contains less than two rows.
;                  @Error 1 @Extended 4 Return ? = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error 1 @Extended 5 Return ? = ColorStop Transparency value not an Integer, less than 0 or greater than 100. Returning problem element index.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve AnnotationShape Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve FillTransparenceGradient Struct.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve ColorStops Array.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve StopColor Struct.
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
; Related .......: _LOCalc_TransparencyGradientMultiModify, _LOCalc_TransparencyGradientMultiDelete, _LOCalc_TransparencyGradientMultiAdd, _LOCalc_CommentAreaGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentAreaTransparencyGradientMulti(ByRef $oComment, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tStyleGradient = $oAnnotationShape.FillTransparenceGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($avColorStops) Then
		$atColorStops = $tStyleGradient.ColorStops()
		If Not IsArray($atColorStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		ReDim $avNewColorStops[UBound($atColorStops)][2]

		For $i = 0 To UBound($atColorStops) - 1
			$avNewColorStops[$i][0] = $atColorStops[$i].StopOffset()
			$tStopColor = $atColorStops[$i].StopColor()
			If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			$avNewColorStops[$i][1] = Int($tStopColor.Red() * 100) ; One value is the same as all.
			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
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
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tStopColor.Red = ($avColorStops[$i][1] / 100)
		$tStopColor.Green = ($avColorStops[$i][1] / 100)
		$tStopColor.Blue = ($avColorStops[$i][1] / 100)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oAnnotationShape.FillTransparenceGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oAnnotationShape.FillTransparenceGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentAreaTransparencyGradientMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentCallout
; Description ...: Set or Retrieve Comment Callout settings.
; Syntax ........: _LOCalc_CommentCallout(ByRef $oComment[, $iCalloutStyle = Null[, $iSpacing = Null[, $iExtension = Null[, $iExtendBy = Null[, $bOptimal = Null[, $iLength = Null]]]]]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $iCalloutStyle       - [optional] an integer value (0-2). Default is Null. The Style of Callout connector line. See Constants $LOC_COMMENT_CALLOUT_STYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iSpacing            - [optional] an integer value (0-240005). Default is Null. The amount of space between the Callout connector line end and the comment box, in Micrometers.
;                  $iExtension          - [optional] an integer value (0-4). Default is Null. The position to extend the Callout line from. See Constants $LOC_COMMENT_CALLOUT_EXT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iExtendBy           - [optional] an integer value (0-240005;0,5000,10000). Default is Null. The length to extend the Callout line, in Micrometers, or the alignment of the line depending on the current setting of $iExtension. See remarks. See Constants $LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_*, or $LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bOptimal            - [optional] a boolean value. Default is Null. If True a angled line will be used optimally.
;                  $iLength             - [optional] an integer value (0-240005). Default is Null. The length of the callout line, in Micrometers.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iCalloutStyle not an Integer, less than 0 or greater than 2. See Constants $LOC_COMMENT_CALLOUT_STYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iSpacing not an Integer, less than 0 or greater than 240,005 Micrometers.
;                  @Error 1 @Extended 4 Return 0 = $iExtension not an Integer, less than 0 or greater than 4. See Constants $LOC_COMMENT_CALLOUT_EXT_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iExtendBy not an Integer, not equal to 0, 5,000, or 10,000. See Constants $LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_*, as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iExtendBy not an Integer, not equal to 0, 5,000, or 10,000. See Constants $LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iExtendBy not an Integer, less than 0 or greater than 240,005 Micrometers.
;                  @Error 1 @Extended 8 Return 0 = $bOptimal not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iLength not an Integer, less than 0 or greater than 240,005 Micrometers.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotation Shape Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iCalloutStyle
;                  |                               2 = Error setting $iSpacing
;                  |                               4 = Error setting $iExtension
;                  |                               8 = Error setting $iExtendBy
;                  |                               16 = Error setting $bOptimal
;                  |                               32 = Error setting $iLength
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $iExtension is set to $LOC_COMMENT_CALLOUT_EXT_HORI, or $LOC_COMMENT_CALLOUT_EXT_VERT, $iExtendBy will be set to the alignment value of either constants $LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_*, or $LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_*.
;                  If $iExtension is set to $LOC_COMMENT_CALLOUT_EXT_OPTIMAL, $LOC_COMMENT_CALLOUT_EXT_FROM_LEFT, or $LOC_COMMENT_CALLOUT_EXT_FROM_TOP, $iExtendBy will be set to the length to extend the Callout line from the Comment box, in Micrometers.
;                  If $iCalloutStyle is not set to $LOC_COMMENT_CALLOUT_STYLE_ANGLED_CONNECTOR, both $bOptimal and $iLength, are not used/unavailable for setting.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentCallout(ByRef $oComment, $iCalloutStyle = Null, $iSpacing = Null, $iExtension = Null, $iExtendBy = Null, $bOptimal = Null, $iLength = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $iError = 0
	Local $aiCallout[6]

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iCalloutStyle, $iSpacing, $iExtension, $iExtendBy, $bOptimal, $iLength) Then
		__LO_ArrayFill($aiCallout, $oAnnotationShape.CaptionType(), $oAnnotationShape.CaptionGap(), $oAnnotationShape.CaptionEscapeDirection(), _
				(($oAnnotationShape.CaptionIsEscapeRelative()) ? ($oAnnotationShape.CaptionEscapeRelative()) : ($oAnnotationShape.CaptionEscapeAbsolute())), _
				$oAnnotationShape.CaptionIsFitLineLength(), $oAnnotationShape.CaptionLineLength())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiCallout)
	EndIf

	If ($iCalloutStyle <> Null) Then
		If Not __LO_IntIsBetween($iCalloutStyle, $LOC_COMMENT_CALLOUT_STYLE_STRAIGHT, $LOC_COMMENT_CALLOUT_STYLE_ANGLED_CONNECTOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oAnnotationShape.CaptionType = $iCalloutStyle
		$iError = ($oAnnotationShape.CaptionType() = $iCalloutStyle) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iSpacing <> Null) Then
		If Not __LO_IntIsBetween($iSpacing, 0, 240005) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oAnnotationShape.CaptionGap = $iSpacing
		$iError = ($oAnnotationShape.CaptionGap() = $iSpacing) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iExtension <> Null) Then
		If Not __LO_IntIsBetween($iExtension, $LOC_COMMENT_CALLOUT_EXT_HORI, $LOC_COMMENT_CALLOUT_EXT_FROM_TOP) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If __LO_IntIsBetween($iExtension, $LOC_COMMENT_CALLOUT_EXT_OPTIMAL, $LOC_COMMENT_CALLOUT_EXT_FROM_TOP) Then
			If ($oAnnotationShape.CaptionIsEscapeRelative() = True) Then
				$oAnnotationShape.CaptionIsEscapeRelative = False
				$oAnnotationShape.CaptionEscapeAbsolute = 0
			EndIf

			Switch $iExtension
				Case $LOC_COMMENT_CALLOUT_EXT_OPTIMAL
					$oAnnotationShape.CaptionEscapeDirection = $iExtension
					$iError = ($oAnnotationShape.CaptionEscapeDirection() = $iExtension) ? ($iError) : (BitOR($iError, 4))

				Case $LOC_COMMENT_CALLOUT_EXT_FROM_LEFT ; From Left, is the same as Vertical setting with different CaptionIsEscapeRelative
					$oAnnotationShape.CaptionEscapeDirection = $LOC_COMMENT_CALLOUT_EXT_VERT
					$iError = ($oAnnotationShape.CaptionEscapeDirection() = $LOC_COMMENT_CALLOUT_EXT_VERT) ? ($iError) : (BitOR($iError, 4))

				Case $LOC_COMMENT_CALLOUT_EXT_FROM_TOP ; From Top, is the same as Horizontal setting with different CaptionIsEscapeRelative
					$oAnnotationShape.CaptionEscapeDirection = $LOC_COMMENT_CALLOUT_EXT_HORI
					$iError = ($oAnnotationShape.CaptionEscapeDirection() = $LOC_COMMENT_CALLOUT_EXT_HORI) ? ($iError) : (BitOR($iError, 4))
			EndSwitch

		Else
			If ($oAnnotationShape.CaptionIsEscapeRelative() = False) Then
				$oAnnotationShape.CaptionIsEscapeRelative = True
				$oAnnotationShape.CaptionEscapeRelative = (($iExtension = $LOC_COMMENT_CALLOUT_EXT_HORI) ? ($LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_TOP) : ($LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_LEFT))
			EndIf
			$oAnnotationShape.CaptionEscapeDirection = $iExtension
			$iError = ($oAnnotationShape.CaptionEscapeDirection() = $iExtension) ? ($iError) : (BitOR($iError, 4))
		EndIf
	EndIf

	If ($iExtendBy <> Null) Then
		If ($oAnnotationShape.CaptionIsEscapeRelative() = True) Then
			If ($oAnnotationShape.CaptionEscapeDirection() = $LOC_COMMENT_CALLOUT_EXT_HORI) Then
				If Not __LO_IntIsBetween($iExtendBy, $LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_TOP, $LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_TOP, "", String($LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_MIDDLE & ":" & $LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_BOTTOM)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			Else
				If Not __LO_IntIsBetween($iExtendBy, $LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_LEFT, $LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_LEFT, "", String($LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_CENTER & ":" & $LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_RIGHT)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
			EndIf
			$oAnnotationShape.CaptionEscapeRelative = $iExtendBy
			$iError = ($oAnnotationShape.CaptionEscapeRelative() = $iExtendBy) ? ($iError) : (BitOR($iError, 8))

		Else
			If Not __LO_IntIsBetween($iExtendBy, 0, 240005) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oAnnotationShape.CaptionEscapeAbsolute = $iExtendBy
			$iError = ($oAnnotationShape.CaptionEscapeAbsolute() = $iExtendBy) ? ($iError) : (BitOR($iError, 8))
		EndIf
	EndIf

	If ($bOptimal <> Null) Then
		If Not IsBool($bOptimal) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oAnnotationShape.CaptionIsFitLineLength = $bOptimal
		$iError = ($oAnnotationShape.CaptionIsFitLineLength() = $bOptimal) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iLength <> Null) Then
		If Not __LO_IntIsBetween($iLength, 0, 240005) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oAnnotationShape.CaptionLineLength = $iLength
		$iError = ($oAnnotationShape.CaptionLineLength() = $iLength) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentCallout

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentCreateTextCursor
; Description ...: Create a Text Cursor in a Comment.
; Syntax ........: _LOCalc_CommentCreateTextCursor(ByRef $oComment[, $bAtEnd = True])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $bAtEnd              - [optional] a boolean value. Default is True. If True, The Text Cursor will be created at the end of any Text Content present.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bAtEnd not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Text Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning newly created Text Cursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentCreateTextCursor(ByRef $oComment, $bAtEnd = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bAtEnd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextCursor = $oComment.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $bAtEnd Then
		$oTextCursor.gotoEnd(False)

	Else
		$oTextCursor.gotoStart(False)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOCalc_CommentCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentDelete
; Description ...: Delete a comment from a Cell.
; Syntax ........: _LOCalc_CommentDelete(ByRef $oComment)
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Cell Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Annotations Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Comment's Index number.
;                  @Error 3 @Extended 4 Return 0 = Failed to delete Comment.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Comment was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_CommentAdd, _LOCalc_CommentGetObjByCell, _LOCalc_CommentGetObjByIndex, _LOCalc_CommentsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentDelete(ByRef $oComment)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotations, $oCell
	Local $iIndex

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oCell = $oComment.Parent()
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oAnnotations = $oCell.Spreadsheet.Annotations()
	If Not IsObj($oAnnotations) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iIndex = __LOCalc_CommentGetObjByCell($oCell, True)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oAnnotations.removeByIndex($iIndex)

	__LOCalc_CommentGetObjByCell($oCell, True)
	If (@error = 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; Comment still exists.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CommentDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentGetCell
; Description ...: Retrieve the Object of the Cell containing this Object.
; Syntax ........: _LOCalc_CommentGetCell(ByRef $oComment)
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve parent Cell Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning Object for Cell containing this comment.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentGetCell(ByRef $oComment)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oCell = $oComment.Parent()
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCell)
EndFunc   ;==>_LOCalc_CommentGetCell

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentGetLastEdit
; Description ...: Retrieve the Last Edit Date and Author of the Comment.
; Syntax ........: _LOCalc_CommentGetLastEdit(ByRef $oComment)
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Author.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve last Edit date.
;                  --Success--
;                  @Error 0 @Extended 0 Return Array = Success. Returning 2 element array containing Author and Date values.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The returned array will have two elements, the first element will contain the Author, and the second will be the last edit Date and Time.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentGetLastEdit(ByRef $oComment)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asEdit[2]
	Local $sAuthor, $sDate

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sAuthor = $oComment.Author()
	If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sDate = $oComment.Date()
	If Not IsString($sDate) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	__LO_ArrayFill($asEdit, $sAuthor, $sDate)

	Return SetError($__LO_STATUS_SUCCESS, 0, $asEdit)
EndFunc   ;==>_LOCalc_CommentGetLastEdit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentGetObjByCell
; Description ...: Retrieve a comment Object for a particular cell, if one exists.
; Syntax ........: _LOCalc_CommentGetObjByCell(ByRef $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell not a Cell Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve requested Comment Object, or no Comment present in Cell.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Comment's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_CommentDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentGetObjByCell(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotation

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.SupportsService("com.sun.star.sheet.SheetCell") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oAnnotation = __LOCalc_CommentGetObjByCell($oCell)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oAnnotation)
EndFunc   ;==>_LOCalc_CommentGetObjByCell

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentGetObjByIndex
; Description ...: Retrieve a comment object by Index.
; Syntax ........: _LOCalc_CommentGetObjByIndex(ByRef $oSheet, $iComment)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iComment            - an integer value. The Index number of the comment to retrieve. 0 based.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iComment not an Integer.
;                  @Error 1 @Extended 3 Return 0 = Index number called in $iComment less than 0, or greater than number of Comments contained in Sheet.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotations Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve total count of Comments.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve requested Comment Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Comment's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_CommentDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentGetObjByIndex(ByRef $oSheet, $iComment)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotations, $oAnnotation
	Local $iCount = 0

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oAnnotations = $oSheet.Annotations()
	If Not IsObj($oAnnotations) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oAnnotations.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iComment, 0, ($iCount - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oAnnotation = $oAnnotations.getByIndex($iComment)
	If Not IsObj($oAnnotation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oAnnotation)
EndFunc   ;==>_LOCalc_CommentGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentLineArrowStyles
; Description ...: Set or Retrieve Comment Line Start and End Arrow Style settings.
; Syntax ........: _LOCalc_CommentLineArrowStyles(ByRef $oComment[, $vStartStyle = Null[, $iStartWidth = Null[, $bStartCenter = Null[, $bSync = Null[, $vEndStyle = Null[, $iEndWidth = Null[, $bEndCenter = Null]]]]]]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $vStartStyle         - [optional] a variant value (0-32, or String). Default is Null. The Arrow head to apply to the start of the line. Can be a Custom Arrowhead name, or one of the constants, $LOC_COMMENT_LINE_ARROW_TYPE_* as defined in LibreOfficeCalc_Constants.au3. See remarks.
;                  $iStartWidth         - [optional] an integer value (0-5004). Default is Null. The Width of the Starting Arrowhead, in Micrometers.
;                  $bStartCenter        - [optional] a boolean value. Default is Null. If True, Places the center of the Start arrowhead on the endpoint of the line.
;                  $bSync               - [optional] a boolean value. Default is Null. If True, Synchronizes the Start Arrowhead settings with the end Arrowhead settings. See remarks.
;                  $vEndStyle           - [optional] a variant value (0-32, or String). Default is Null. The Arrow head to apply to the end of the line. Can be a Custom Arrowhead name, or one of the constants, $LOC_COMMENT_LINE_ARROW_TYPE_* as defined in LibreOfficeCalc_Constants.au3. See remarks.
;                  $iEndWidth           - [optional] an integer value (0-5004). Default is Null. The Width of the Ending Arrowhead, in Micrometers.
;                  $bEndCenter          - [optional] a boolean value. Default is Null. If True, Places the center of the End arrowhead on the endpoint of the line.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $vStartStyle not a String, and not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $vStartStyle is an Integer, but less than 0, or greater than 32. See constants $LOC_COMMENT_LINE_ARROW_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iStartWidth not an Integer, less than 0, or greater than 5004.
;                  @Error 1 @Extended 5 Return 0 = $bStartCenter not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bSync not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $vEndStyle not a String, and not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $vSEndStyle is an Integer, but less than 0, or greater than 32. See constants $LOC_COMMENT_LINE_ARROW_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $iEndWidth not an Integer, less than 0, or greater than 5004.
;                  @Error 1 @Extended 10 Return 0 = $bEndCenter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Annotation Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to convert Constant to Arrowhead name.
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
; Remarks .......: Libre Office has no setting for $bSync, so I have made a manual version of it in this function. It only accepts True, and must be called with True each time you want it to synchronize.
;                  When retrieving the current settings, $bSync will be a Boolean value of whether the Start Arrowhead settings are currently equal to the End Arrowhead setting values.
;                  Both $vStartStyle and $vEndStyle accept a String or an Integer because there is the possibility of a custom Arrowhead being available the user may want to use.
;                  When retrieving the current settings, both $vStartStyle and $vEndStyle could be either an integer or a String. It will be a String if the current Arrowhead is a custom Arrowhead, else an Integer, corresponding to one of the constants, $LOC_COMMENT_LINE_ARROW_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_CommentLineProperties
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentLineArrowStyles(ByRef $oComment, $vStartStyle = Null, $iStartWidth = Null, $bStartCenter = Null, $bSync = Null, $vEndStyle = Null, $iEndWidth = Null, $bEndCenter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $iError = 0
	Local $avArrow[7]
	Local $sStartStyle, $sEndStyle

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($vStartStyle, $iStartWidth, $bStartCenter, $bSync, $vEndStyle, $iEndWidth, $bEndCenter) Then
		__LO_ArrayFill($avArrow, __LOCalc_CommentArrowStyleName(Null, $oAnnotationShape.LineStartName()), $oAnnotationShape.LineStartWidth(), $oAnnotationShape.LineStartCenter(), _
				((($oAnnotationShape.LineStartName() = $oAnnotationShape.LineEndName()) And ($oAnnotationShape.LineStartWidth() = $oAnnotationShape.LineEndWidth()) And ($oAnnotationShape.LineStartCenter() = $oAnnotationShape.LineEndCenter())) ? (True) : (False)), _ ; See if Start and End are the same.
				__LOCalc_CommentArrowStyleName(Null, $oAnnotationShape.LineEndName()), $oAnnotationShape.LineEndWidth(), $oAnnotationShape.LineEndCenter())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avArrow)
	EndIf

	If ($vStartStyle <> Null) Then
		If Not IsString($vStartStyle) And Not IsInt($vStartStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If IsInt($vStartStyle) Then
			If Not __LO_IntIsBetween($vStartStyle, $LOC_COMMENT_LINE_ARROW_TYPE_NONE, $LOC_COMMENT_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$sStartStyle = __LOCalc_CommentArrowStyleName($vStartStyle)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Else
			$sStartStyle = $vStartStyle
		EndIf

		$oAnnotationShape.LineStartName = $sStartStyle
		$iError = ($oAnnotationShape.LineStartName() = $sStartStyle) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iStartWidth <> Null) Then
		If Not __LO_IntIsBetween($iStartWidth, 0, 5004) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oAnnotationShape.LineStartWidth = $iStartWidth
		$iError = (__LO_IntIsBetween($oAnnotationShape.LineStartWidth(), $iStartWidth - 1, $iStartWidth + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bStartCenter <> Null) Then
		If Not IsBool($bStartCenter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oAnnotationShape.LineStartCenter = $bStartCenter
		$iError = ($oAnnotationShape.LineStartCenter() = $bStartCenter) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bSync <> Null) Then
		If Not IsBool($bSync) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		If ($bSync = True) Then
			$oAnnotationShape.LineEndName = $oAnnotationShape.LineStartName()
			$oAnnotationShape.LineEndWidth = $oAnnotationShape.LineStartWidth()
			$oAnnotationShape.LineEndCenter = $oAnnotationShape.LineStartCenter()
			$iError = (($oAnnotationShape.LineStartName() = $oAnnotationShape.LineEndName()) And _
					($oAnnotationShape.LineStartWidth() = $oAnnotationShape.LineEndWidth()) And _
					($oAnnotationShape.LineStartCenter() = $oAnnotationShape.LineEndCenter())) ? ($iError) : (BitOR($iError, 8))
		EndIf
	EndIf

	If ($vEndStyle <> Null) Then
		If Not IsString($vEndStyle) And Not IsInt($vEndStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		If IsInt($vEndStyle) Then
			If Not __LO_IntIsBetween($vEndStyle, $LOC_COMMENT_LINE_ARROW_TYPE_NONE, $LOC_COMMENT_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$sEndStyle = __LOCalc_CommentArrowStyleName($vEndStyle)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Else
			$sEndStyle = $vEndStyle
		EndIf

		$oAnnotationShape.LineEndName = $sEndStyle
		$iError = ($oAnnotationShape.LineEndName() = $sEndStyle) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iEndWidth <> Null) Then
		If Not __LO_IntIsBetween($iEndWidth, 0, 5004) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oAnnotationShape.LineEndWidth = $iEndWidth
		$iError = (__LO_IntIsBetween($oAnnotationShape.LineEndWidth(), $iEndWidth - 1, $iEndWidth + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bEndCenter <> Null) Then
		If Not IsBool($bEndCenter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oAnnotationShape.LineEndCenter = $bEndCenter
		$iError = ($oAnnotationShape.LineEndCenter() = $bEndCenter) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentLineArrowStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentLineProperties
; Description ...: Set or Retrieve Comment Line settings.
; Syntax ........: _LOCalc_CommentLineProperties(ByRef $oComment[, $vStyle = Null[, $iColor = Null[, $iWidth = Null[, $iTransparency = Null[, $iCornerStyle = Null[, $iCapStyle = Null]]]]]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $vStyle              - [optional] a variant value (0-31, or String). Default is Null. The Line Style to use. Can be a Custom Line Style name, or one of the constants, $LOC_COMMENT_LINE_STYLE_* as defined in LibreOfficeCalc_Constants.au3. See remarks.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Line color, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iWidth              - [optional] an integer value (0-5004). Default is Null. The line Width, set in Micrometers.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The Line transparency percentage. 100% = fully transparent.
;                  $iCornerStyle        - [optional] an integer value (0,2-4). Default is Null. The Line Corner Style. See Constants $LOC_COMMENT_LINE_JOINT_* as defined in LibreOfficeCalc_Constants.au3
;                  $iCapStyle           - [optional] an integer value (0-2). Default is Null. The Line Cap Style. See Constants $LOC_COMMENT_LINE_CAP_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $vStyle not a String, and not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $vStyle is an Integer, but less than 0, or greater than 31. See constants $LOC_COMMENT_LINE_STYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iWidth not an Integer, less than 0, or greater than 5004.
;                  @Error 1 @Extended 6 Return 0 = $iTransparency not an Integer, less than 0, or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iCornerStyle not an Integer, not equal to 0, equal to 1, not equal to 2 or greater than 4. See Constants $LOC_COMMENT_LINE_JOINT_* as defined in LibreOfficeCalc_Constants.au3
;                  @Error 1 @Extended 8 Return 0 = $iCapStyle is an Integer, but less than 0, or greater than 2. See constants $LOC_COMMENT_LINE_CAP_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Annotation Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to convert Constant to Line Style name.
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
;                  When retrieving the current settings, $vStyle could be either an integer or a String. It will be a String if the current Line Style is a custom Line Style, else an Integer, corresponding to one of the constants, $LOC_COMMENT_LINE_STYLE_* as defined in LibreOfficeCalc_Constants.au3.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOCalc_CommentLineArrowStyles
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentLineProperties(ByRef $oComment, $vStyle = Null, $iColor = Null, $iWidth = Null, $iTransparency = Null, $iCornerStyle = Null, $iCapStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $iError = 0
	Local Const $__LOC_COMMENT_LINE_STYLE_NONE = 0, $__LOC_COMMENT_LINE_STYLE_SOLID = 1, $__LOC_COMMENT_LINE_STYLE_DASH = 2
	Local $avLine[6]
	Local $sStyle
	Local $vReturn

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($vStyle, $iColor, $iWidth, $iTransparency, $iCornerStyle, $iCapStyle) Then
		Switch $oAnnotationShape.LineStyle()
			Case $__LOC_COMMENT_LINE_STYLE_NONE
				$vReturn = $LOC_COMMENT_LINE_STYLE_NONE

			Case $__LOC_COMMENT_LINE_STYLE_SOLID
				$vReturn = $LOC_COMMENT_LINE_STYLE_CONTINUOUS

			Case $__LOC_COMMENT_LINE_STYLE_DASH
				$vReturn = __LOCalc_CommentLineStyleName(Null, $oAnnotationShape.LineDashName())
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndSwitch

		__LO_ArrayFill($avLine, $vReturn, $oAnnotationShape.LineColor(), $oAnnotationShape.LineWidth(), $oAnnotationShape.LineTransparence(), $oAnnotationShape.LineJoint(), $oAnnotationShape.LineCap())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avLine)
	EndIf

	If ($vStyle <> Null) Then
		If Not IsString($vStyle) And Not IsInt($vStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If IsInt($vStyle) Then
			If Not __LO_IntIsBetween($vStyle, $LOC_COMMENT_LINE_STYLE_NONE, $LOC_COMMENT_LINE_STYLE_LINE_WITH_FINE_DOTS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			Switch $vStyle
				Case $LOC_COMMENT_LINE_STYLE_NONE
					$oAnnotationShape.LineStyle = $__LOC_COMMENT_LINE_STYLE_NONE
					$iError = ($oAnnotationShape.LineStyle() = $__LOC_COMMENT_LINE_STYLE_NONE) ? ($iError) : (BitOR($iError, 1))

				Case $LOC_COMMENT_LINE_STYLE_CONTINUOUS
					$oAnnotationShape.LineStyle = $__LOC_COMMENT_LINE_STYLE_SOLID
					$iError = ($oAnnotationShape.LineStyle() = $__LOC_COMMENT_LINE_STYLE_SOLID) ? ($iError) : (BitOR($iError, 1))

				Case Else
					$sStyle = __LOCalc_CommentLineStyleName($vStyle)
					If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

					$oAnnotationShape.LineStyle = $__LOC_COMMENT_LINE_STYLE_DASH
					$oAnnotationShape.LineDashName = $sStyle
					$iError = ($oAnnotationShape.LineDashName() = $sStyle) ? ($iError) : (BitOR($iError, 1))
			EndSwitch

		Else
			$sStyle = $vStyle
			$oAnnotationShape.LineDashName = $sStyle
			$iError = ($oAnnotationShape.LineDashName() = $sStyle) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oAnnotationShape.LineColor = $iColor
		$iError = ($oAnnotationShape.LineColor() = $iColor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 0, 5004) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oAnnotationShape.LineWidth = $iWidth
		$iError = (__LO_IntIsBetween($oAnnotationShape.LineWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oAnnotationShape.LineTransparence = $iTransparency
		$iError = ($oAnnotationShape.LineTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iCornerStyle <> Null) Then
		If Not __LO_IntIsBetween($iCornerStyle, $LOC_COMMENT_LINE_JOINT_NONE, $LOC_COMMENT_LINE_JOINT_ROUND, $LOC_COMMENT_LINE_JOINT_MIDDLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oAnnotationShape.LineJoint = $iCornerStyle
		$iError = ($oAnnotationShape.LineJoint() = $iCornerStyle) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iCapStyle <> Null) Then
		If Not __LO_IntIsBetween($iCapStyle, $LOC_COMMENT_LINE_CAP_FLAT, $LOC_COMMENT_LINE_CAP_SQUARE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oAnnotationShape.LineCap = $iCapStyle
		$iError = ($oAnnotationShape.LineCap() = $iCapStyle) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentLineProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentPosition
; Description ...: Set or Retrieve the Comment's position settings.
; Syntax ........: _LOCalc_CommentPosition(ByRef $oComment[, $iX = Null[, $iY = Null[, $bProtectPos = Null]]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $iX                  - [optional] an integer value. Default is Null. The X position from the insertion point, in Micrometers.
;                  $iY                  - [optional] an integer value. Default is Null. The Y position from the insertion point, in Micrometers.
;                  $bProtectPos         - [optional] a boolean value. Default is Null. If True, the Comment's position is locked.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $bProtectPos not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Annotation Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Comment's Position Structure.
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
; Remarks .......: The X Coordinate seems to be measured from the right hand edge of the Comment Box, and the Y Coordinate seems to be measured from the top of the comment box.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer, _LOCalc_CommentRotate, _LOCalc_CommentSize
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentPosition(ByRef $oComment, $iX = Null, $iY = Null, $bProtectPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $iError = 0
	Local $avPosition[3]
	Local $tPos

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tPos = $oAnnotationShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iX, $iY, $bProtectPos) Then
		__LO_ArrayFill($avPosition, $tPos.X(), $tPos.Y(), $oAnnotationShape.MoveProtect())

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

		$oAnnotationShape.Position = $tPos

		$iError = ($iX = Null) ? ($iError) : ((__LO_IntIsBetween($oAnnotationShape.Position.X(), $iX - 1, $iX + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = ($iY = Null) ? ($iError) : ((__LO_IntIsBetween($oAnnotationShape.Position.Y(), $iY - 1, $iY + 1)) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($bProtectPos <> Null) Then
		If Not IsBool($bProtectPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oAnnotationShape.MoveProtect = $bProtectPos
		$iError = ($oAnnotationShape.MoveProtect() = $bProtectPos) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentRotate
; Description ...: Set or retrieve Rotation settings for a Comment.
; Syntax ........: _LOCalc_CommentRotate(ByRef $oComment[, $nRotate = Null])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $nRotate             - [optional] a general number value (0-359.99). Default is Null. The Degrees to rotate the Comment. See remarks.
; Return values .: Success: 1 or Number.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $nRotate not a Number, less than 0, or greater than 359.99.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Annotation Shape Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $nRotate
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Number = Success. All optional parameters were set to Null, returning current setting as a Number.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function uses the deprecated Libre Office method RotateAngle and may stop working in future Libre Office versions, after 7.3.4.2.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOCalc_CommentPosition, _LOCalc_CommentSize
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentRotate(ByRef $oComment, $nRotate = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($nRotate = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, (($oAnnotationShape.RotateAngle()) / 100)) ; Divide by 100 to match L.O. values.

	If Not __LO_NumIsBetween($nRotate, 0, 359.99) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oAnnotationShape.RotateAngle = ($nRotate * 100)     ; * 100 to match L.O. Values.
	If (($oAnnotationShape.RotateAngle() / 100) <> $nRotate) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CommentRotate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentsGetCount
; Description ...: Retrieve a count of Comments contained in the Sheet.
; Syntax ........: _LOCalc_CommentsGetCount(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotations Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve total count of Comments.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning total number of comments contained in the Sheet as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentsGetCount(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotations
	Local $iCount = 0

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotations = $oSheet.Annotations()
	If Not IsObj($oAnnotations) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oAnnotations.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOCalc_CommentsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentsGetList
; Description ...: Retrieve an array of all comments contained in a Sheet.
; Syntax ........: _LOCalc_CommentsGetList(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotations Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve total count of Comments.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning array of Comment Objects contained in the Sheet. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentsGetList(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotations
	Local $aoAnnotations[0]
	Local $iCount = 0

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotations = $oSheet.Annotations()
	If Not IsObj($oAnnotations) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oAnnotations.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	ReDim $aoAnnotations[$iCount]

	For $i = 0 To $iCount - 1
		$aoAnnotations[$i] = $oAnnotations.getByIndex($i)
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoAnnotations)
EndFunc   ;==>_LOCalc_CommentsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentSize
; Description ...: Set or Retrieve Comment Size related settings.
; Syntax ........: _LOCalc_CommentSize(ByRef $oComment[, $iWidth = Null[, $iHeight = Null[, $bProtectSize = Null]]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Comment, in Micrometers(uM). Min. 51.
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the Comment, in Micrometers(uM). Min. 51.
;                  $bProtectSize        - [optional] a boolean value. Default is Null. If True, Locks the size of the Comment.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer, or less than 51.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer, or less than 51.
;                  @Error 1 @Extended 4 Return 0 = $bProtectSize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Annotation Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Comment's Size Structure.
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
;                  I have skipped "Keep Ratio, as there is no built in setting for it for Comments, so I would have to formulate a custom function for this purpose.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer, _LOCalc_CommentPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentSize(ByRef $oComment, $iWidth = Null, $iHeight = Null, $bProtectSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $iError = 0
	Local $avSize[3]
	Local $tSize

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tSize = $oAnnotationShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iWidth, $iHeight, $bProtectSize) Then
		__LO_ArrayFill($avSize, $tSize.Width(), $tSize.Height(), $oAnnotationShape.SizeProtect())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSize)
	EndIf

	If ($iWidth <> Null) Or ($iHeight <> Null) Then
		If ($iWidth <> Null) Then ; Min 51
			If Not __LO_IntIsBetween($iWidth, 51, $iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

			$tSize.Width = $iWidth
		EndIf

		If ($iHeight <> Null) Then
			If Not __LO_IntIsBetween($iHeight, 51, $iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$tSize.Height = $iHeight
		EndIf

		$oAnnotationShape.Size = $tSize

		$iError = ($iWidth = Null) ? ($iError) : ((__LO_IntIsBetween($oAnnotationShape.Size.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = ($iHeight = Null) ? ($iError) : ((__LO_IntIsBetween($oAnnotationShape.Size.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($bProtectSize <> Null) Then
		If Not IsBool($bProtectSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oAnnotationShape.SizeProtect = $bProtectSize
		$iError = ($oAnnotationShape.SizeProtect() = $bProtectSize) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentText
; Description ...: Set or Retrieve the current text contained in the Comment.
; Syntax ........: _LOCalc_CommentText(ByRef $oComment[, $sText = Null])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $sText               - [optional] a string value. Default is Null. The text to set the Comment to. Will overwrite any previous text.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sText not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current text content of the comment.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sText
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Text were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning current text contained in Comment, as a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current text content of the comment.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentText(ByRef $oComment, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sString

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($sText = Null) Then
		$sString = $oComment.String()
		If Not IsString($sString) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sString)
	EndIf

	If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oComment.String = $sText

	If ($oComment.String() <> $sText) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CommentText

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentTextAnchor
; Description ...: Set or Retrieve Comment Text Anchor settings.
; Syntax ........: _LOCalc_CommentTextAnchor(ByRef $oComment[, $iAnchor = Null[, $bFullWidth = Null]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $iAnchor             - [optional] an integer value (0-8). Default is Null. The Comment Anchor position. See Constants $LOC_COMMENT_ANCHOR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bFullWidth          - [optional] a boolean value. Default is Null. If True, the text will be expanded to cover the full width of the comment.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iAnchor not an Integer, less than 0 or greater than 8. See Constants $LOC_COMMENT_ANCHOR_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $bFullWidth not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotation Shape Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iAnchor
;                  |                               2 = Error setting $bFullWidth
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentTextAnchor(ByRef $oComment, $iAnchor = Null, $bFullWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $iError = 0, $iCurrentAnchor
	Local $avAnchor[2]
	Local Const $__LOC_HORI_ALIGN_LEFT = 0, $__LOC_HORI_ALIGN_CENTER = 1, $__LOC_HORI_ALIGN_RIGHT = 2, $__LOC_HORI_ALIGN_BLOCK = 3 ; com.sun.star.drawing.TextHorizontalAdjust
	Local Const $__LOC_VERT_ALIGN_TOP = 0, $__LOC_VERT_ALIGN_CENTER = 1, $__LOC_VERT_ALIGN_BOTTOM = 2, $__LOC_VERT_ALIGN_BLOCK = 3 ; com.sun.star.drawing.TextVerticalAdjust

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iAnchor, $bFullWidth) Then
		Switch $oAnnotationShape.TextVerticalAdjust()
			Case $__LOC_VERT_ALIGN_TOP
				Switch $oAnnotationShape.TextHorizontalAdjust()
					Case $__LOC_HORI_ALIGN_LEFT
						$iCurrentAnchor = $LOC_COMMENT_ANCHOR_TOP_LEFT

					Case $__LOC_HORI_ALIGN_CENTER, $__LOC_HORI_ALIGN_BLOCK
						$iCurrentAnchor = $LOC_COMMENT_ANCHOR_TOP_CENTER

					Case $__LOC_HORI_ALIGN_RIGHT
						$iCurrentAnchor = $LOC_COMMENT_ANCHOR_TOP_RIGHT
				EndSwitch

			Case $__LOC_VERT_ALIGN_CENTER, $__LOC_VERT_ALIGN_BLOCK
				Switch $oAnnotationShape.TextHorizontalAdjust()
					Case $__LOC_HORI_ALIGN_LEFT
						$iCurrentAnchor = $LOC_COMMENT_ANCHOR_MIDDLE_LEFT

					Case $__LOC_HORI_ALIGN_CENTER, $__LOC_HORI_ALIGN_BLOCK
						$iCurrentAnchor = $LOC_COMMENT_ANCHOR_MIDDLE_CENTER

					Case $__LOC_HORI_ALIGN_RIGHT
						$iCurrentAnchor = $LOC_COMMENT_ANCHOR_MIDDLE_RIGHT
				EndSwitch

			Case $__LOC_VERT_ALIGN_BOTTOM
				Switch $oAnnotationShape.TextHorizontalAdjust()
					Case $__LOC_HORI_ALIGN_LEFT
						$iCurrentAnchor = $LOC_COMMENT_ANCHOR_BOTTOM_LEFT

					Case $__LOC_HORI_ALIGN_CENTER, $__LOC_HORI_ALIGN_BLOCK
						$iCurrentAnchor = $LOC_COMMENT_ANCHOR_BOTTOM_CENTER

					Case $__LOC_HORI_ALIGN_RIGHT
						$iCurrentAnchor = $LOC_COMMENT_ANCHOR_BOTTOM_RIGHT
				EndSwitch
		EndSwitch

		If Not IsInt($iCurrentAnchor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LO_ArrayFill($avAnchor, $iCurrentAnchor, ($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_BLOCK) ? (True) : (False))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avAnchor)
	EndIf

	If ($iAnchor <> Null) Then
		If Not __LO_IntIsBetween($iAnchor, $LOC_COMMENT_ANCHOR_TOP_LEFT, $LOC_COMMENT_ANCHOR_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		Switch $iAnchor
			Case $LOC_COMMENT_ANCHOR_TOP_LEFT
				$oAnnotationShape.TextVerticalAdjust = $__LOC_VERT_ALIGN_TOP
				$oAnnotationShape.TextHorizontalAdjust = $__LOC_HORI_ALIGN_LEFT
				$iError = (($oAnnotationShape.TextVerticalAdjust() = $__LOC_VERT_ALIGN_TOP) And ($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_LEFT)) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_COMMENT_ANCHOR_TOP_CENTER
				$oAnnotationShape.TextVerticalAdjust = $__LOC_VERT_ALIGN_TOP
				$oAnnotationShape.TextHorizontalAdjust = ($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_BLOCK) ? ($__LOC_HORI_ALIGN_BLOCK) : ($__LOC_HORI_ALIGN_CENTER)
				$iError = (($oAnnotationShape.TextVerticalAdjust() = $__LOC_VERT_ALIGN_TOP) And _
						($oAnnotationShape.TextHorizontalAdjust() = (($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_BLOCK) ? ($__LOC_HORI_ALIGN_BLOCK) : ($__LOC_HORI_ALIGN_CENTER)))) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_COMMENT_ANCHOR_TOP_RIGHT
				$oAnnotationShape.TextVerticalAdjust = $__LOC_VERT_ALIGN_TOP
				$oAnnotationShape.TextHorizontalAdjust = $__LOC_HORI_ALIGN_RIGHT
				$iError = (($oAnnotationShape.TextVerticalAdjust() = $__LOC_VERT_ALIGN_TOP) And _
						($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_RIGHT)) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_COMMENT_ANCHOR_MIDDLE_LEFT
				$oAnnotationShape.TextVerticalAdjust = $__LOC_VERT_ALIGN_CENTER
				$oAnnotationShape.TextHorizontalAdjust = $__LOC_HORI_ALIGN_LEFT
				$iError = (($oAnnotationShape.TextVerticalAdjust() = $__LOC_VERT_ALIGN_CENTER) And _
						($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_LEFT)) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_COMMENT_ANCHOR_MIDDLE_CENTER
				$oAnnotationShape.TextVerticalAdjust = $__LOC_VERT_ALIGN_CENTER
				$oAnnotationShape.TextHorizontalAdjust = ($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_BLOCK) ? ($__LOC_HORI_ALIGN_BLOCK) : ($__LOC_HORI_ALIGN_CENTER)
				$iError = (($oAnnotationShape.TextVerticalAdjust() = $__LOC_VERT_ALIGN_CENTER) And _
						($oAnnotationShape.TextHorizontalAdjust() = (($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_BLOCK) ? ($__LOC_HORI_ALIGN_BLOCK) : ($__LOC_HORI_ALIGN_CENTER)))) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_COMMENT_ANCHOR_MIDDLE_RIGHT
				$oAnnotationShape.TextVerticalAdjust = $__LOC_VERT_ALIGN_CENTER
				$oAnnotationShape.TextHorizontalAdjust = $__LOC_HORI_ALIGN_RIGHT
				$iError = (($oAnnotationShape.TextVerticalAdjust() = $__LOC_VERT_ALIGN_CENTER) And _
						($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_RIGHT)) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_COMMENT_ANCHOR_BOTTOM_LEFT
				$oAnnotationShape.TextVerticalAdjust = $__LOC_VERT_ALIGN_BOTTOM
				$oAnnotationShape.TextHorizontalAdjust = $__LOC_HORI_ALIGN_LEFT
				$iError = (($oAnnotationShape.TextVerticalAdjust() = $__LOC_VERT_ALIGN_BOTTOM) And _
						($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_LEFT)) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_COMMENT_ANCHOR_BOTTOM_CENTER
				$oAnnotationShape.TextVerticalAdjust = $__LOC_VERT_ALIGN_BOTTOM
				$oAnnotationShape.TextHorizontalAdjust = ($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_BLOCK) ? ($__LOC_HORI_ALIGN_BLOCK) : ($__LOC_HORI_ALIGN_CENTER)
				$iError = (($oAnnotationShape.TextVerticalAdjust() = $__LOC_VERT_ALIGN_BOTTOM) And _
						($oAnnotationShape.TextHorizontalAdjust() = (($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_BLOCK) ? ($__LOC_HORI_ALIGN_BLOCK) : ($__LOC_HORI_ALIGN_CENTER)))) ? ($iError) : (BitOR($iError, 1))

			Case $LOC_COMMENT_ANCHOR_BOTTOM_RIGHT
				$oAnnotationShape.TextVerticalAdjust = $__LOC_VERT_ALIGN_BOTTOM
				$oAnnotationShape.TextHorizontalAdjust = $__LOC_HORI_ALIGN_RIGHT
				$iError = (($oAnnotationShape.TextVerticalAdjust() = $__LOC_VERT_ALIGN_BOTTOM) And _
						($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_RIGHT)) ? ($iError) : (BitOR($iError, 1))
		EndSwitch
	EndIf

	If ($bFullWidth <> Null) Then
		If Not IsBool($bFullWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If $bFullWidth Then
			$oAnnotationShape.TextHorizontalAdjust = $__LOC_HORI_ALIGN_BLOCK
			$iError = ($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_BLOCK) ? ($iError) : (BitOR($iError, 2))

		Else
			If ($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_BLOCK) Then
				$oAnnotationShape.TextHorizontalAdjust = $__LOC_HORI_ALIGN_CENTER
				$iError = ($oAnnotationShape.TextHorizontalAdjust() = $__LOC_HORI_ALIGN_CENTER) ? ($iError) : (BitOR($iError, 2))
			EndIf
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentTextAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentTextAnimation
; Description ...: Set or Retrieve Comment Text Animation settings.
; Syntax ........: _LOCalc_CommentTextAnimation(ByRef $oComment[, $iAnimation = Null[, $iDirection = Null[, $bBeginInside = Null[, $bVisibleOnExit = Null[, $iCycles = Null[, $iPixelIncrement = Null[, $iIncrement = Null[, $iDelay = Null]]]]]]]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $iAnimation          - [optional] an integer value (0-4). Default is Null. The Animation type. See Constants $LOC_COMMENT_ANIMATION_KIND_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iDirection          - [optional] an integer value (0-3). Default is Null. The Animation direction. See Constants $LOC_COMMENT_ANIMATION_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bBeginInside        - [optional] a boolean value. Default is Null. If True, the text is visible inside the Comment when the animation begins.
;                  $bVisibleOnExit      - [optional] a boolean value. Default is Null. If True, the text is visible inside the Comment when the animation ends.
;                  $iCycles             - [optional] an integer value (0-100). Default is Null. The number of times to repeat the animation. 0 = continuous.
;                  $iPixelIncrement     - [optional] an integer value (0-100). Default is Null. The increment value measured in pixels. See remarks.
;                  $iIncrement          - [optional] an integer value (0-25400). Default is Null. The increment value measured in Micrometers. See remarks.
;                  $iDelay              - [optional] an integer value (0-30000). Default is Null. The delay between animation repeats, in milliseconds. 0 = Automatic.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iAnimation not an Integer, less than 0 or greater than 4. See Constants $LOC_COMMENT_ANIMATION_KIND_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iDirection not an Integer, less than 0 or greater than 3. See Constants $LOC_COMMENT_ANIMATION_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $bBeginInside not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bVisibleOnExit not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iCycles not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iPixelIncrement not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 8 Return 0 = $iIncrement not an Integer, less than 0 or greater than 25,400.
;                  @Error 1 @Extended 9 Return 0 = $iDelay not an Integer, less than 0 or greater than 30,000.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotation Shape Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iAnimation
;                  |                               2 = Error setting $iDirection
;                  |                               4 = Error setting $bBeginInside
;                  |                               8 = Error setting $bVisibleOnExit
;                  |                               16 = Error setting $iCycles
;                  |                               32 = Error setting $iPixelIncrement
;                  |                               64 = Error setting $iIncrement
;                  |                               128 = Error setting $iDelay
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Increment value can be set in either Pixels or Micrometers, but internally it uses the same property. To make setting this easier, I have made two separate parameters to set, depending on if you wish to set increment using pixels ($iPixelIncrement) or Micrometers ($iIncrement). Only one should be used, otherwise $iIncrement will overwrite $iPixelIncrement.
;                  When either $iPixelIncrement or $iIncrement is not set, meaning if you are setting the Increment value in pixels, or not, one or the other value will return 0 when retrieving the current settings.
;                  $iIncrement in the L.O. UI allows for 10" max, however this produces an erroneous value internally, and switches back to using pixels, even in the UI, if you set $iIncrement to the max value, it will most likely cause a property setting error.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentTextAnimation(ByRef $oComment, $iAnimation = Null, $iDirection = Null, $bBeginInside = Null, $bVisibleOnExit = Null, $iCycles = Null, $iPixelIncrement = Null, $iIncrement = Null, $iDelay = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $iError = 0
	Local $avAnim[8]

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iAnimation, $iDirection, $bBeginInside, $bVisibleOnExit, $iCycles, $iPixelIncrement, $iIncrement, $iDelay) Then
		__LO_ArrayFill($avAnim, $oAnnotationShape.TextAnimationKind(), $oAnnotationShape.TextAnimationDirection(), $oAnnotationShape.TextAnimationStartInside(), _
				$oAnnotationShape.TextAnimationStopInside(), $oAnnotationShape.TextAnimationCount(), _
				($oAnnotationShape.TextAnimationAmount() < 0) ? ($oAnnotationShape.TextAnimationAmount() * -1) : (0), _ ; Convert from negative to positive for return.
				($oAnnotationShape.TextAnimationAmount() < 0) ? (0) : ($oAnnotationShape.TextAnimationAmount()), _
				$oAnnotationShape.TextAnimationDelay())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avAnim)
	EndIf

	If ($iAnimation <> Null) Then
		If Not __LO_IntIsBetween($iAnimation, $LOC_COMMENT_ANIMATION_KIND_NONE, $LOC_COMMENT_ANIMATION_KIND_SCROLL_IN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oAnnotationShape.TextAnimationKind = $iAnimation
		$iError = ($oAnnotationShape.TextAnimationKind() = $iAnimation) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iDirection <> Null) Then
		If Not __LO_IntIsBetween($iDirection, $LOC_COMMENT_ANIMATION_DIR_LEFT, $LOC_COMMENT_ANIMATION_DIR_DOWN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oAnnotationShape.TextAnimationDirection = $iDirection
		$iError = ($oAnnotationShape.TextAnimationDirection() = $iDirection) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bBeginInside <> Null) Then
		If Not IsBool($bBeginInside) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oAnnotationShape.TextAnimationStartInside = $bBeginInside
		$iError = ($oAnnotationShape.TextAnimationStartInside() = $bBeginInside) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bVisibleOnExit <> Null) Then
		If Not IsBool($bVisibleOnExit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oAnnotationShape.TextAnimationStopInside = $bVisibleOnExit
		$iError = ($oAnnotationShape.TextAnimationStopInside() = $bVisibleOnExit) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iCycles <> Null) Then
		If Not __LO_IntIsBetween($iCycles, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oAnnotationShape.TextAnimationCount = $iCycles
		$iError = ($oAnnotationShape.TextAnimationCount() = $iCycles) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iPixelIncrement <> Null) Then
		If Not __LO_IntIsBetween($iPixelIncrement, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oAnnotationShape.TextAnimationAmount = ($iPixelIncrement * -1) ; Pixel Increment set in negative numbers.
		$iError = ($oAnnotationShape.TextAnimationAmount() = ($iPixelIncrement * -1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 0, 25400) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oAnnotationShape.TextAnimationAmount = $iIncrement
		$iError = ($oAnnotationShape.TextAnimationAmount() = $iIncrement) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 30000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oAnnotationShape.TextAnimationDelay = $iDelay
		$iError = ($oAnnotationShape.TextAnimationDelay() = $iDelay) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentTextAnimation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentTextColumns
; Description ...: Set or Retrieve Comment Text Column settings.
; Syntax ........: _LOCalc_CommentTextColumns(ByRef $oDoc, ByRef $oComment[, $iColumns = Null[, $iSpacing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $iColumns            - [optional] an integer value (1-16). Default is Null. The number of columns to break the text area into.
;                  $iSpacing            - [optional] an integer value. Default is Null. The amount of spacing between the columns, in Micrometers.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iColumns not an Integer, less than 0 or greater than 16.
;                  @Error 1 @Extended 4 Return 0 = $iSpacing not an Integer or less than 0.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextColumns" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotation Shape Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Text Columns Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColumns
;                  |                               2 = Error setting $iSpacing
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentTextColumns(ByRef $oDoc, ByRef $oComment, $iColumns = Null, $iSpacing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextColumns, $oAnnotationShape
	Local $aiColumn[2]
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTextColumns = $oAnnotationShape.TextColumns()

	If ($oTextColumns = "") Then ; If Columns haven't been set before, Text Columns is Null/empty string, I have to create a Text Columns Object and then insert it.
		$oTextColumns = $oDoc.createInstance("com.sun.star.text.TextColumns")
		If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oTextColumns.ColumnCount = 1
		$oAnnotationShape.TextColumns = $oTextColumns
		$oTextColumns = $oAnnotationShape.TextColumns()
	EndIf

	If Not IsObj($oTextColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iColumns, $iSpacing) Then
		__LO_ArrayFill($aiColumn, $oTextColumns.ColumnCount(), $oTextColumns.AutomaticDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiColumn)
	EndIf

	If ($iColumns <> Null) Then
		If Not __LO_IntIsBetween($iColumns, 1, 16) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTextColumns.ColumnCount = $iColumns
	EndIf

	If ($iSpacing <> Null) Then
		If Not __LO_IntIsBetween($iSpacing, 0, $iSpacing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTextColumns.AutomaticDistance = $iSpacing
	EndIf

	$oAnnotationShape.TextColumns = $oTextColumns

	$iError = ($iColumns = Null) ? ($iError) : (($oAnnotationShape.TextColumns.ColumnCount() = $iColumns) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iSpacing = Null) ? ($iError) : ((__LO_IntIsBetween($oAnnotationShape.TextColumns.AutomaticDistance(), $iSpacing - 1, $iSpacing + 1)) ? ($iError) : (BitOR($iError, 1)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentTextColumns

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentTextSettings
; Description ...: Set or Retrieve Comment Text settings.
; Syntax ........: _LOCalc_CommentTextSettings(ByRef $oComment[, $bFitWidth = Null[, $bFitHeight = Null[, $bFitToFrame = Null[, $iSpacingAll = Null[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]]]]]])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $bFitWidth           - [optional] a boolean value. Default is Null. If True, the comment box width will expand to fit text content.
;                  $bFitHeight          - [optional] a boolean value. Default is Null. If True, the comment box height will expand to fit text content.
;                  $bFitToFrame         - [optional] a boolean value. Default is Null. If True text will be resized to fit the frame.
;                  $iSpacingAll         - [optional] an integer value (-100000-100000). Default is Null. Set the spacing around the text between the text and the Comment borders, in Micrometers.
;                  $iLeft               - [optional] an integer value (-100000-100000). Default is Null. Set the spacing on the Left side of the text between the text and the Comment border, in Micrometers.
;                  $iRight              - [optional] an integer value (-100000-100000). Default is Null. Set the spacing on the Right side of the text between the text and the Comment border, in Micrometers.
;                  $iTop                - [optional] an integer value (-100000-100000). Default is Null. Set the spacing on the Top side of the text between the text and the Comment border, in Micrometers.
;                  $iBottom             - [optional] an integer value (-100000-100000). Default is Null. Set the spacing on the Bottom side of the text between the text and the Comment border, in Micrometers.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFitWidth not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bFitHeight not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bFitToFrame not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iSpacingAll not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 6 Return 0 = $iLeft not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 7 Return 0 = $iRight not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 8 Return 0 = $iTop not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 9 Return 0 = $iBottom not an Integer, less than -100,000 or greater than 100,000.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotation Shape Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFitWidth
;                  |                               2 = Error setting $bFitHeight
;                  |                               4 = Error setting $bFitToFrame
;                  |                               8 = Error setting $iSpacingAll
;                  |                               16 = Error setting $iLeft
;                  |                               32 = Error setting $iRight
;                  |                               64 = Error setting $iTop
;                  |                               128 = Error setting $iBottom
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If either $bFitWidth or $bFitHeight is set to True, $bFitToFrame cannot be set to True, and vice versa.
;                  If spacing values on all sides do not match, $iSpacingAll will return 0.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentTextSettings(ByRef $oComment, $bFitWidth = Null, $bFitHeight = Null, $bFitToFrame = Null, $iSpacingAll = Null, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnnotationShape
	Local $iError = 0
	Local $avText[8]
	; TextFitToSize = None, AUTOFIT or proportional, Starts off as None, once set to proportional and switched off, becomes AutoFit
	Local Const $__LOC_COMMENT_TEXT_FIT_FRAME_PROP = 1, $__LOC_COMMENT_TEXT_FIT_FRAME_AUTO_FIT = 3 ; $__LOC_COMMENT_TEXT_FIT_FRAME_NONE = 0,

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oAnnotationShape = $oComment.AnnotationShape()
	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bFitWidth, $bFitHeight, $bFitToFrame, $iSpacingAll, $iLeft, $iRight, $iTop, $iBottom) Then
		__LO_ArrayFill($avText, $oAnnotationShape.TextAutoGrowWidth(), $oAnnotationShape.TextAutoGrowHeight(), _
				($oAnnotationShape.TextFitToSize() = $__LOC_COMMENT_TEXT_FIT_FRAME_PROP) ? (True) : (False), _
				(($oAnnotationShape.TextLeftDistance() = $oAnnotationShape.TextRightDistance()) And _
				($oAnnotationShape.TextRightDistance() = $oAnnotationShape.TextUpperDistance()) And _
				($oAnnotationShape.TextUpperDistance() = $oAnnotationShape.TextLowerDistance())) ? ($oAnnotationShape.TextLeftDistance()) : (0), _
				$oAnnotationShape.TextLeftDistance(), $oAnnotationShape.TextRightDistance(), $oAnnotationShape.TextUpperDistance(), _
				$oAnnotationShape.TextLowerDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avText)
	EndIf

	If ($bFitWidth <> Null) Then
		If Not IsBool($bFitWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oAnnotationShape.TextAutoGrowWidth = $bFitWidth
		$iError = ($oAnnotationShape.TextAutoGrowWidth() = $bFitWidth) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bFitHeight <> Null) Then
		If Not IsBool($bFitHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oAnnotationShape.TextAutoGrowHeight = $bFitHeight
		$iError = ($oAnnotationShape.TextAutoGrowHeight() = $bFitHeight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bFitToFrame <> Null) Then
		If Not IsBool($bFitToFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oAnnotationShape.TextFitToSize = ($bFitToFrame) ? ($__LOC_COMMENT_TEXT_FIT_FRAME_PROP) : ($__LOC_COMMENT_TEXT_FIT_FRAME_AUTO_FIT)
		If $bFitToFrame Then
			$iError = ($oAnnotationShape.TextFitToSize() = $__LOC_COMMENT_TEXT_FIT_FRAME_PROP) ? ($iError) : (BitOR($iError, 4))

		Else
			$iError = ($oAnnotationShape.TextFitToSize() <> $__LOC_COMMENT_TEXT_FIT_FRAME_PROP) ? ($iError) : (BitOR($iError, 4))
		EndIf
	EndIf

	If ($iSpacingAll <> Null) Then
		If Not __LO_IntIsBetween($iSpacingAll, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oAnnotationShape.TextLeftDistance = $iSpacingAll
		$oAnnotationShape.TextRightDistance = $iSpacingAll
		$oAnnotationShape.TextUpperDistance = $iSpacingAll
		$oAnnotationShape.TextLowerDistance = $iSpacingAll
		$iError = (($oAnnotationShape.TextLeftDistance() = $oAnnotationShape.TextRightDistance()) And _
				($oAnnotationShape.TextRightDistance() = $oAnnotationShape.TextUpperDistance()) And _
				($oAnnotationShape.TextUpperDistance() = $oAnnotationShape.TextLowerDistance())) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oAnnotationShape.TextLeftDistance = $iLeft
		$iError = ($oAnnotationShape.TextLeftDistance() = $iLeft) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oAnnotationShape.TextRightDistance = $iRight
		$iError = ($oAnnotationShape.TextRightDistance() = $iRight) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oAnnotationShape.TextUpperDistance = $iTop
		$iError = ($oAnnotationShape.TextUpperDistance() = $iTop) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oAnnotationShape.TextLowerDistance = $iBottom
		$iError = ($oAnnotationShape.TextLowerDistance() = $iBottom) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_CommentTextSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CommentVisible
; Description ...: Set or Retrieve the Comment's visibility settings.
; Syntax ........: _LOCalc_CommentVisible(ByRef $oComment[, $bVisible = Null])
; Parameters ....: $oComment            - [in/out] an object. A Comment object returned by a previous _LOCalc_CommentsGetList, _LOCalc_CommentGetObjByCell, or _LOCalc_CommentGetObjByIndex function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the comment will be always visible.
; Return values .: Success: 1 or Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComment not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were set to Null, returning current setting as a Boolean value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current setting.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CommentVisible(ByRef $oComment, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oComment) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oComment.IsVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oComment.IsVisible = $bVisible

	If ($oComment.IsVisible() <> $bVisible) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CommentVisible
