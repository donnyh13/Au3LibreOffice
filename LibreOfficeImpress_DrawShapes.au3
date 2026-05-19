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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Deleting, etc. Impress Drawing Shapes, such as lines and rectangles etc.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_DrawShapeAltText
; _LOImpress_DrawShapeConnectorModify
; _LOImpress_DrawShapeConnectorSettings
; _LOImpress_DrawShapeDimensionSettings
; _LOImpress_DrawShapeGetType
; _LOImpress_DrawShapeInsert
; _LOImpress_DrawShapePointsAdd
; _LOImpress_DrawShapePointsGetCount
; _LOImpress_DrawShapePointsModify
; _LOImpress_DrawShapePointsRemove
; _LOImpress_DrawShapeText
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeAltText
; Description ...: Set or Retrieve Draw Shape Alternate text settings.
; Syntax ........: _LOImpress_DrawShapeAltText(ByRef $oDrawShape[, $sText = Null[, $sAltText = Null[, $bDecorative = Null]]])
; Parameters ....: $oDrawShape          - A Drawing Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $sText               - [optional] Default is Null. Enter alternative text to display when the image isn't available.
;                  $sAltText            - [optional] Default is Null. Detailed alternative text of the Image.
;                  $bDecorative         - [optional] Default is Null. If True, the image is considered decorative and is ignored by assistive readers. L.O. 7.6+.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDrawShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = Shape called in $oDrawShape not a drawing shape.
;                  @Error 1 @Extended 3 Return 0 = $sText not a string.
;                  @Error 1 @Extended 4 Return 0 = $sAltText not a string.
;                  @Error 1 @Extended 5 Return 0 = $bDecorative not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sText
;                  |                               2 = Error setting $sAltText
;                  |                               4 = Error setting $bDecorative
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current LibreOffice version less than 7.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters. If The current LibreOffice version is below 7.6 the $bDecorative parameter will return a Null value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  These properties are only available for shapes other than lines (e.g. squares, stars, etc.).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeAltText(ByRef $oDrawShape, $sText = Null, $sAltText = Null, $bDecorative = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asAltTxt[3]

	If Not IsObj($oDrawShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDrawShape.supportsService("com.sun.star.drawing.Shape") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sText, $sAltText, $bDecorative) Then
		If __LO_VersionCheck(7.6) Then
			__LO_ArrayFill($asAltTxt, $oDrawShape.Title(), $oDrawShape.Description(), $oDrawShape.Decorative())

		Else
			__LO_ArrayFill($asAltTxt, $oDrawShape.Title(), $oDrawShape.Description(), Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $asAltTxt)
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDrawShape.Title = $sText
		$iError = ($oDrawShape.Title() = $sText) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sAltText <> Null) Then
		If Not IsString($sAltText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDrawShape.Description = $sAltText
		$iError = ($oDrawShape.Description() = $sAltText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bDecorative <> Null) Then
		If Not IsBool($bDecorative) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oDrawShape.Decorative = $bDecorative
		$iError = ($oDrawShape.Decorative() = $bDecorative) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeAltText

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeConnectorModify
; Description ...: Set or Retrieve Connector line connections or position.
; Syntax ........: _LOImpress_DrawShapeConnectorModify(ByRef $oShape[, $iStartX = Null[, $iStartY = Null[, $oStartShape = Null[, $iStartGluePoint = Null[, $iEndX = Null[, $iEndY = Null[, $oEndShape = Null[, $iEndGluePoint = Null]]]]]]]])
; Parameters ....: $oShape              - A Connector Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iStartX             - [optional] Default is Null. The X position from the insertion point of the Start of the line, in Hundredths of a Millimeter (HMM).
;                  $iStartY             - [optional] Default is Null. The Y position from the insertion point of the Start of the line, in Hundredths of a Millimeter (HMM).
;                  $oStartShape         - [optional] Default is Null. The Shape to attach the Start of the line to.
;                  $iStartGluePoint     - [optional] Default is Null. If the Start of the line is connected to a Shape, this is the Glue point it is attached to. 0 Based. See remarks.
;                  $iEndX               - [optional] Default is Null. The X position from the insertion point of the End of the line, in Hundredths of a Millimeter (HMM).
;                  $iEndY               - [optional] Default is Null. The Y position from the insertion point of the End of the line, in Hundredths of a Millimeter (HMM).
;                  $oEndShape           - [optional] Default is Null. The Shape to attach the End of the line to.
;                  $iEndGluePoint       - [optional] Default is Null. If the End of the line is connected to a Shape, this is the Glue point it is attached to. 0 Based. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iStartX not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iStartY not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $oStartShape not an Object.
;                  @Error 1 @Extended 5 Return 0 = $iStartGluePoint not an Integer, or less than -1.
;                  @Error 1 @Extended 6 Return 0 = $iEndX not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iEndY not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $oEndShape not an Object.
;                  @Error 1 @Extended 9 Return 0 = $iEndGluePoint not an Integer, or less than -1.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Start Position.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve End Position.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iStartX
;                  |                               2 = Error setting $iStartY
;                  |                               4 = Error setting $oStartShape
;                  |                               8 = Error setting $iStartGluePoint
;                  |                               16 = Error setting $iEndX
;                  |                               32 = Error setting $iEndY
;                  |                               64 = Error setting $oEndShape
;                  |                               128 = Error setting $iEndGluePoint
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If no shape is set for either the Start or the End, Null is returned when retrieving the current settings.
;                  If no shape is set for either the Start or the End, the GluePoint return will be -1.
;                  When a Start shape or an End shape is currently active, you can retrieve, but NOT set, Start or End position respectively.
;                  When setting a Start or End shape, if a GluePoint isn't called, LibreOffice chooses one automatically.
;                  Currently, it seems to be not possible to disconnect a shape from the Start or End programatically.
;                  Both $iStartGluePoint and $iEndGluePoint do not check if the value called too high, i.e., a higher GluePoint index than present. They also accept -1, but I see nothing noticeable that it does.
;                  The index of the default GluePoints are 0 (top), 1 (right), 2 (bottom), and 3 (left). You also can add new glue points to a shape’s default GluePoints.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeConnectorModify(ByRef $oShape, $iStartX = Null, $iStartY = Null, $oStartShape = Null, $iStartGluePoint = Null, $iEndX = Null, $iEndY = Null, $oEndShape = Null, $iEndGluePoint = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oS_Shape, $oE_Shape
	Local $tPos
	Local $avConnector[8]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iStartX, $iStartY, $oStartShape, $iStartGluePoint, $iEndX, $iEndY, $oEndShape, $iEndGluePoint) Then
		$oS_Shape = $oShape.StartShape()
		$oE_Shape = $oShape.EndShape()

		__LO_ArrayFill($avConnector, $oShape.StartPosition.X(), $oShape.StartPosition.Y(), (IsObj($oS_Shape)) ? ($oS_Shape) : (Null), _ ; If Start shape isn't set, return Null.
				$oShape.StartGluePointIndex(), $oShape.EndPosition.X(), $oShape.EndPosition.Y(), (IsObj($oE_Shape)) ? ($oE_Shape) : (Null), _ ; If End shape isn't set, return Null.
				$oShape.EndGluePointIndex())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avConnector)
	EndIf

	If ($iStartX <> Null) Or ($iStartY <> Null) Then
		$tPos = $oShape.StartPosition()
		If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If ($iStartX <> Null) Then
			If Not IsInt($iStartX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

			$tPos.X = $iStartX
		EndIf

		If ($iStartY <> Null) Then
			If Not IsInt($iStartY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$tPos.Y = $iStartY
		EndIf

		$oShape.StartPosition = $tPos
		$iError = (__LO_VarsAreNull($iStartX)) ? ($iError) : ((__LO_IntIsBetween($oShape.StartPosition.X(), $iStartX - 1, $iStartX + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = (__LO_VarsAreNull($iStartY)) ? ($iError) : ((__LO_IntIsBetween($oShape.StartPosition.Y(), $iStartY - 1, $iStartY + 1)) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($oStartShape <> Null) Then
		If Not IsObj($oStartShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.StartShape = $oStartShape
		$iError = ($oShape.StartShape() = $oStartShape) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iStartGluePoint <> Null) Then
		If Not __LO_IntIsBetween($iStartGluePoint, -1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oShape.StartGluePointIndex = $iStartGluePoint
		$iError = ($oShape.StartGluePointIndex() = $iStartGluePoint) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iEndX <> Null) Or ($iEndY <> Null) Then
		$tPos = $oShape.EndPosition()
		If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		If ($iEndX <> Null) Then
			If Not IsInt($iEndX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$tPos.X = $iEndX
		EndIf

		If ($iEndY <> Null) Then
			If Not IsInt($iEndY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$tPos.Y = $iEndY
		EndIf

		$oShape.EndPosition = $tPos
		$iError = (__LO_VarsAreNull($iEndX)) ? ($iError) : ((__LO_IntIsBetween($oShape.EndPosition.X(), $iEndX - 1, $iEndX + 1)) ? ($iError) : (BitOR($iError, 16)))
		$iError = (__LO_VarsAreNull($iEndY)) ? ($iError) : ((__LO_IntIsBetween($oShape.EndPosition.Y(), $iEndY - 1, $iEndY + 1)) ? ($iError) : (BitOR($iError, 32)))
	EndIf

	If ($oEndShape <> Null) Then
		If Not IsObj($oEndShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oShape.EndShape = $oEndShape
		$iError = ($oShape.EndShape() = $oEndShape) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iEndGluePoint <> Null) Then
		If Not __LO_IntIsBetween($iEndGluePoint, -1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oShape.EndGluePointIndex = $iEndGluePoint
		$iError = ($oShape.EndGluePointIndex() = $iEndGluePoint) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeConnectorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeConnectorSettings
; Description ...: Set or Retrieve Connector line settings.
; Syntax ........: _LOImpress_DrawShapeConnectorSettings(ByRef $oShape[, $iType = Null[, $iL1Skew = Null[, $iL2Skew = Null[, $iL3Skew = Null[, $iHoriBeg = Null[, $iHoriEnd = Null[, $iVertBeg = Null[, $iVertEnd = Null]]]]]]]])
; Parameters ....: $oShape              - A Connector Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iType               - [optional] (0-3) Default is Null. The connector line type. See Constants, $LOI_DRAWSHAPE_CONNECTOR_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iL1Skew             - [optional] (-100000-100000) Default is Null. The skew amount of line 1, in Hundredths of a Millimeter (HMM).
;                  $iL2Skew             - [optional] (-100000-100000) Default is Null. The skew amount of line 2, in Hundredths of a Millimeter (HMM).
;                  $iL3Skew             - [optional] (-100000-100000) Default is Null. The skew amount of line 3, in Hundredths of a Millimeter (HMM).
;                  $iHoriBeg            - [optional] (0-10008) Default is Null. The amount of horizontal spacing, in Hundredths of a Millimeter (HMM), at the beginning of the connector.
;                  $iHoriEnd            - [optional] (0-10008) Default is Null. The amount of horizontal spacing, in Hundredths of a Millimeter (HMM), at the end of the connector.
;                  $iVertBeg            - [optional] (0-10008) Default is Null. The amount of vertical spacing, in Hundredths of a Millimeter (HMM), at the beginning of the connector.
;                  $iVertEnd            - [optional] (0-10008) Default is Null. The amount of vertical spacing, in Hundredths of a Millimeter (HMM), at the end of the connector.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 0 or greater than 3. See Constants, $LOI_DRAWSHAPE_CONNECTOR_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iL1Skew not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 4 Return 0 = $iL2Skew not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 5 Return 0 = $iL3Skew not an Integer, less than -100,000 or greater than 100,000.
;                  @Error 1 @Extended 6 Return 0 = $iHoriBeg not an Integer, less than 0 or greater than 10,008.
;                  @Error 1 @Extended 7 Return 0 = $iHoriEnd not an Integer, less than 0 or greater than 10,008.
;                  @Error 1 @Extended 8 Return 0 = $iVertBeg not an Integer, less than 0 or greater than 10,008.
;                  @Error 1 @Extended 9 Return 0 = $iVertEnd not an Integer, less than 0 or greater than 10,008.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iType
;                  |                               2 = Error setting $iL1Skew
;                  |                               4 = Error setting $iL2Skew
;                  |                               8 = Error setting $iL3Skew
;                  |                               16 = Error setting $iHoriBeg
;                  |                               32 = Error setting $iHoriEnd
;                  |                               64 = Error setting $iVertBeg
;                  |                               128 = Error setting $iVertEnd
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeConnectorSettings(ByRef $oShape, $iType = Null, $iL1Skew = Null, $iL2Skew = Null, $iL3Skew = Null, $iHoriBeg = Null, $iHoriEnd = Null, $iVertBeg = Null, $iVertEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avConnector[8]

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iType, $iL1Skew, $iL2Skew, $iL3Skew, $iHoriBeg, $iHoriEnd, $iVertBeg, $iVertEnd) Then
		__LO_ArrayFill($avConnector, $oShape.EdgeKind(), $oShape.EdgeLine1Delta(), $oShape.EdgeLine2Delta(), $oShape.EdgeLine3Delta(), $oShape.EdgeNode1HorzDist(), _
				$oShape.EdgeNode2HorzDist(), $oShape.EdgeNode1VertDist(), $oShape.EdgeNode2VertDist())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avConnector)
	EndIf

	If ($iType <> Null) Then
		If Not __LO_IntIsBetween($iType, $LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD, $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oShape.EdgeKind = $iType
		$iError = ($oShape.EdgeKind() = $iType) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iL1Skew <> Null) Then
		If Not __LO_IntIsBetween($iL1Skew, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oShape.EdgeLine1Delta = $iL1Skew
		$iError = (__LO_IntIsBetween($oShape.EdgeLine1Delta(), $iL1Skew - 1, $iL1Skew + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iL2Skew <> Null) Then
		If Not __LO_IntIsBetween($iL2Skew, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oShape.EdgeLine2Delta = $iL2Skew
		$iError = (__LO_IntIsBetween($oShape.EdgeLine2Delta(), $iL2Skew - 1, $iL2Skew + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iL3Skew <> Null) Then
		If Not __LO_IntIsBetween($iL3Skew, -100000, 100000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oShape.EdgeLine3Delta = $iL3Skew
		$iError = (__LO_IntIsBetween($oShape.EdgeLine3Delta(), $iL3Skew - 1, $iL3Skew + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iHoriBeg <> Null) Then
		If Not __LO_IntIsBetween($iHoriBeg, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oShape.EdgeNode1HorzDist = $iHoriBeg
		$iError = (__LO_IntIsBetween($oShape.EdgeNode1HorzDist(), $iHoriBeg - 1, $iHoriBeg + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iHoriEnd <> Null) Then
		If Not __LO_IntIsBetween($iHoriEnd, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oShape.EdgeNode2HorzDist = $iHoriEnd
		$iError = (__LO_IntIsBetween($oShape.EdgeNode2HorzDist(), $iHoriEnd - 1, $iHoriEnd + 1)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iVertBeg <> Null) Then
		If Not __LO_IntIsBetween($iVertBeg, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oShape.EdgeNode1VertDist = $iVertBeg
		$iError = (__LO_IntIsBetween($oShape.EdgeNode1VertDist(), $iVertBeg - 1, $iVertBeg + 1)) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iVertEnd <> Null) Then
		If Not __LO_IntIsBetween($iVertEnd, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oShape.EdgeNode2VertDist = $iVertEnd
		$iError = (__LO_IntIsBetween($oShape.EdgeNode2VertDist(), $iVertEnd - 1, $iVertEnd + 1)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeConnectorSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeDimensionSettings
; Description ...: Set or Retrieve Dimension line settings.
; Syntax ........: _LOImpress_DrawShapeDimensionSettings(ByRef $oShape[, $iDistance = Null[, $iGuideOverhang = Null[, $iGuideDistance = Null[, $iLGuide = Null[, $iRGuide = Null[, $bBelow = Null[, $iDecimal = Null[, $iVertPos = Null[, $iHoriPos = Null[, $bParallel = Null[, $iUnitType = Null]]]]]]]]]]])
; Parameters ....: $oShape              - A Dimension Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
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
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iDistance not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 3 Return 0 = $iGuideOverhang not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 4 Return 0 = $iGuideDistance not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 5 Return 0 = $iLGuide not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 6 Return 0 = $iRGuide not an Integer, less than -10,008 or greater than 10,008.
;                  @Error 1 @Extended 7 Return 0 = $bBelow not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iDecimal not an Integer, less than 0 or greater than 99.
;                  @Error 1 @Extended 9 Return 0 = $iVertPos not an Integer, less than 0 or greater than 4. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $iHoriPos not an Integer, less than 0 or greater than 3. See Constants, $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $bParallel not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iUnitType not an Integer, less than -1 or greater than 15. See Constants, $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
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
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeDimensionSettings(ByRef $oShape, $iDistance = Null, $iGuideOverhang = Null, $iGuideDistance = Null, $iLGuide = Null, $iRGuide = Null, $bBelow = Null, $iDecimal = Null, $iVertPos = Null, $iHoriPos = Null, $bParallel = Null, $iUnitType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_DimensionSettings($oShape, $iDistance, $iGuideOverhang, $iGuideDistance, $iLGuide, $iRGuide, $bBelow, $iDecimal, $iVertPos, $iHoriPos, $bParallel, $iUnitType)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_DrawShapeDimensionSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeGetType
; Description ...: Return the Drawing Shape's Type corresponding to the constants $LOI_DRAWSHAPE_TYPE_*
; Syntax ........: _LOImpress_DrawShapeGetType(ByRef $oShape)
; Parameters ....: $oShape              - A Drawing Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
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
;                  When the arrowhead type "Arrow" is set in the LO UI, or upon creation of a line with arrows, the internal name of the arrowhead is set to an incrementing name of "Arrowheads x", where x is an Integer value. Since I have no way to determine if the head is a custom arrowhead or supposed to be the "Arrow" type, I cannot necessarily identify the right connector or Line.
;                  When setting an Arrowhead to be $LOI_SHAPE_LINE_ARROW_TYPE_ARROW, the head is set correctly, but the LibreOffice UI will show "None". The return for Arrowhead type will be the correct name however, and will allow me to identify the shape.
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
; Syntax ........: _LOImpress_DrawShapeInsert(ByRef $oSlide, $iShapeType, $iWidth, $iHeight[, $iX = 0[, $iY = 0]])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iShapeType          - (0-187) The Type of shape to create. See remarks. See $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $iWidth              - The Shape's Width in Hundredths of a Millimeter (HMM). Note, for Lines, Width is the length of the line.
;                  $iHeight             - The Shape's Height in Hundredths of a Millimeter (HMM). Note, for Lines, Height is the amount the line goes below the point of insertion.
;                  $iX                  - [optional] Default is 0. The X position from the top-left of the page, in Hundredths of a Millimeter (HMM).
;                  $iY                  - [optional] Default is 0. The Y position from the top-left of the page, in Hundredths of a Millimeter (HMM).
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iShapeType not an Integer, less than 0 or greater than 187. See $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 3 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iY not an Integer.
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
; Related .......: _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeInsert(ByRef $oSlide, $iShapeType, $iWidth, $iHeight, $iX = 0, $iY = 0)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iShapeType, $LOI_DRAWSHAPE_TYPE_3D_CONE, $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_3D_CONE To $LOI_DRAWSHAPE_TYPE_3D_TORUS ; Can't create a 3D shape.

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY To $LOI_DRAWSHAPE_TYPE_ARROWS_PENTAGON ; Create an Arrow Shape.
			$oShape = __LOImpress_DrawShape_CreateArrow($oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_BASIC_ARC To $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT ; Create a Basic Shape.
			$oShape = __LOImpress_DrawShape_CreateBasic($oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_CLOUD To $LOI_DRAWSHAPE_TYPE_CALLOUT_ROUND ; Create a Callout Shape.
			$oShape = __LOImpress_DrawShape_CreateCallout($oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_CONNECTOR To $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW ; Create a Connector.
			$oShape = __LOImpress_DrawShape_CreateLine($oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_CARD To $LOI_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR ; Create a Flowchart Shape.
			$oShape = __LOImpress_DrawShape_CreateFlowchart($oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL To $LOI_DRAWSHAPE_TYPE_FONTWORK_TRICOLORE ; Can't create Fontwork.

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS To $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED ; Create a Line Shape.
			$oShape = __LOImpress_DrawShape_CreateLine($oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_STARS_4_POINT To $LOI_DRAWSHAPE_TYPE_STARS_SIGNET ; Create a Star or Banner Shape.
			$oShape = __LOImpress_DrawShape_CreateStars($oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND To $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE ; Create a Symbol Shape.
			$oShape = __LOImpress_DrawShape_CreateSymbol($oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>_LOImpress_DrawShapeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapePointsAdd
; Description ...: Add a Position Point to Shape.
; Syntax ........: _LOImpress_DrawShapePointsAdd(ByRef $oShape, $iPoint, $iX, $iY[, $iPointType = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL[, $bIsCurve = False]])
; Parameters ....: $oShape              - A Drawing Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function. See remarks.
;                  $iPoint              - The Point to insert the new point AFTER. 0 means insert at the beginning.
;                  $iX                  - The X coordinate value, set in Hundredths of a Millimeter (HMM).
;                  $iY                  - The Y coordinate value, set in Hundredths of a Millimeter (HMM).
;                  $iPointType          - [optional] (0, 1, 3) Default is $LOI_DRAWSHAPE_POINT_TYPE_NORMAL. The Type of Point this new Point is. See Remarks. See constants $LOI_DRAWSHAPE_POINT_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $bIsCurve            - [optional] Default is False. If True, the Normal Point is a Curve. See remarks.
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
;                  This is a homemade function as LibreOffice doesn't offer an easy way for adding points to a shape. Consequently this will not produce similar results as when working with LibreOffice manually, and may wreck your shape's shape. Use with caution.
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
; Parameters ....: $oShape              - A Drawing Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oShape does not have property "PolyPolygonBezier", and consequently does not have Position Points that can be modified.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Retrieve Array of Point Type Flags.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning total number of points present in a shape.
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
; Parameters ....: $oShape              - A Drawing Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function. See remarks.
;                  $iPoint              - The Point to modify, starting at 1.
;                  $iX                  - [optional] Default is Null. The X coordinate value, set in Hundredths of a Millimeter (HMM).
;                  $iY                  - [optional] Default is Null. The Y coordinate value, set in Hundredths of a Millimeter (HMM).
;                  $iPointType          - [optional] (0, 1, 3) Default is Null. The Type of Point to change the called point to. See Remarks. See constants $LOI_DRAWSHAPE_POINT_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $bIsCurve            - [optional] Default is Null. If True, the Normal Point is a Curve. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oShape does not have property "PolyPolygonBezier", and consequently does not have Position Points that can be modified.
;                  @Error 1 @Extended 3 Return 0 = $iPoint not an Integer, less than 1 or greater than number of points in the shape.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer
;                  @Error 1 @Extended 6 Return 0 = $PointType not an Integer, less than 0 or greater than 3, or equal to 2.
;                  @Error 1 @Extended 7 Return 0 = $PointType called with other than Normal while $iPoint is referencing first or last point.
;                  @Error 1 @Extended 8 Return 0 = $bIsCurve not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bIsCurve cannot be set for last point in a shape.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Retrieve Array of Point Type Flags.
;                  @Error 3 @Extended 2 Return 0 = Failed to Retrieve Array of Points.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify the requested Array element.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve current settings for requested point.
;                  @Error 3 @Extended 5 Return 0 = Failed to modify the requested point.
;                  @Error 3 @Extended 6 Return 0 = Failed to Retrieve PolyPolygonBezier Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings for the Array Element called in $iArrayElement.
;                  Call any optional parameter with Null keyword to skip it.
;                  Only $LOI_DRAWSHAPE_TYPE_LINE_* type shapes have Points that can be added to, removed, or modified.
;                  This is a homemade function as LibreOffice doesn't offer an easy way for modifying points in a shape. Consequently this will not produce similar results as when working with LibreOffice manually, and may wreck your shape's shape. Use with caution.
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
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$tPolyCoords = $oShape.PolyPolygonBezier()
	If Not IsObj($tPolyCoords) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

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
; Parameters ....: $oShape              - A Drawing Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $iPoint              - The Point to in the Shape to delete, beginning at 1.
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
;                  This is a homemade function as LibreOffice doesn't offer an easy way for removing points in a shape. Consequently this will not produce similar results as when working with LibreOffice manually, and may wreck your shape's shape. Use with caution.
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
; Name ..........: _LOImpress_DrawShapeText
; Description ...: Set or Retrieve the current text displayed in a shape's text box.
; Syntax ........: _LOImpress_DrawShapeText(ByRef $oShape[, $sText = Null])
; Parameters ....: $oShape              - A Drawing Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_ShapesGetList function.
;                  $sText               - [optional] Default is Null. The text to display in the Shape's text box.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an object.
;                  @Error 1 @Extended 2 Return 0 = $sText not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve shape's current text.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sText
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning shape's current text content.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When setting the text of a Shape, any previous text will be overwritten.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeText(ByRef $oShape, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurrText

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sText) Then
		$sCurrText = $oShape.String()
		If Not IsString($sCurrText) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrText)
	EndIf

	If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oShape.String = $sText
	; Strip @CR and @LF, otherwise errors result when comparing.
	$iError = (StringRegExpReplace($oShape.String(), @CR & "|" & @LF, "") = StringRegExpReplace($sText, @CR & "|" & @LF, "")) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeText
