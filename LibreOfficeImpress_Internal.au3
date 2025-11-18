#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Impress
#include "LibreOfficeImpress_Constants.au3"
#include "LibreOfficeImpress_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Various functions for internal data processing, data retrieval, retrieving and applying settings for LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LOImpress_ColorRemoveAlpha
; __LOImpress_CreatePoint
; __LOImpress_DrawShape_CreateArrow
; __LOImpress_DrawShape_CreateBasic
; __LOImpress_DrawShape_CreateCallout
; __LOImpress_DrawShape_CreateFlowchart
; __LOImpress_DrawShape_CreateLine
; __LOImpress_DrawShape_CreateStars
; __LOImpress_DrawShape_CreateSymbol
; __LOImpress_DrawShape_GetCustomType
; __LOImpress_DrawShapeArrowStyleName
; __LOImpress_DrawShapeLineStyleName
; __LOImpress_DrawShapePointGetSettings
; __LOImpress_DrawShapePointModify
; __LOImpress_FilterNameGet
; __LOImpress_GetShapeName
; __LOImpress_GradientNameInsert
; __LOImpress_GradientPresets
; __LOImpress_InternalComErrorHandler
; __LOImpress_ShapeGetType
; __LOImpress_TransparencyGradientConvert
; __LOImpress_TransparencyGradientNameInsert
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ColorRemoveAlpha
; Description ...: Remove the Alpha value from a RGB Color Integer.
; Syntax ........: __LOImpress_ColorRemoveAlpha($iColor)
; Parameters ....: $iColor              - an integer value. A RGB Color Integer to remove Alpha from.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return $iColor = $iColor not an Integer. Returning $iColor to be sure not to lose the value.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Color already has no Alpha value, returning same color.
;                  @Error 0 @Extended 1 Return Integer = Success. Removed Alpha value from RGB Color Integer, returning new Color value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In functions which return the current color value, generally background colors, if Transparency (alpha) is set, the background color value is not the literal color set, but also includes the transparency value added to it. This functions removes that value for simpler color values.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ColorRemoveAlpha($iColor)
	Local $iRed, $iGreen, $iBlue, $iLong

	If Not IsInt($iColor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, $iColor)

	If __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_SUCCESS, 0, $iColor) ; If Color value is not greater than White(16777215) or less than -1, then there is no alpha to remove.

	; Obtain individual color values.
	$iRed = BitAND(BitShift($iColor, 16), 0xff)
	$iGreen = BitAND(BitShift($iColor, 8), 0xff)
	$iBlue = BitAND($iColor, 0xff)
	$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

	Return SetError($__LO_STATUS_SUCCESS, 1, $iLong)
EndFunc   ;==>__LOImpress_ColorRemoveAlpha

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CreatePoint
; Description ...: Creates a Position structure.
; Syntax ........: __LOImpress_CreatePoint($iX, $iY)
; Parameters ....: $iX                  - an integer value. The X position, in Hundredths of a Millimeter (100th MM).
;                  $iY                  - an integer value. The Y position, in Hundredths of a Millimeter (100th MM).
; Return values .: Success: Structure
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 2 Return 0 = $iY not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a Position Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Structure = Success. Returning created Position Structure using $iX and $iY values.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Modified from A. Pitonyak, Listing 493. in OOME 3.0
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CreatePoint($iX, $iY)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tPoint

	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tPoint = __LO_CreateStruct("com.sun.star.awt.Point")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tPoint.X = $iX
	$tPoint.Y = $iY

	Return SetError($__LO_STATUS_SUCCESS, 0, $tPoint)
EndFunc   ;==>__LOImpress_CreatePoint

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateArrow
; Description ...: Create an Arrow type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateArrow(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (100th MM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (100th MM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iShapeType          - an integer value (0-25). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_ARROWS_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" or "com.sun.star.drawing.EllipseShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  @Error 2 @Extended 3 Return 0 = Failed to create "MirroredX" property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT,
;                  $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateArrow(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tProp2, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY
			$tProp.Value = "quad-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY
			$tProp.Value = "quad-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN
			$tProp.Value = "down-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT
			$tProp.Value = "left-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT
			$tProp.Value = "left-right-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT
			$tProp.Value = "right-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP
			$tProp.Value = "up-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN
			$tProp.Value = "up-down-arrow-callout"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
			$tProp.Value = "mso-spt100"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CIRCULAR
			$tProp.Value = "circular-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT
			$tProp.Value = "corner-right-arrow" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_DOWN
			$tProp.Value = "down-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT
			$tProp.Value = "left-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT
			$tProp.Value = "left-right-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT
			$tProp.Value = "notched-right-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT
			$tProp.Value = "right-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT
			$tProp.Value = "split-arrow" ; "non-primitive"??

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED
			$tProp.Value = "s-sharped-arrow" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT
			$tProp.Value = "split-arrow" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT
			$tProp.Value = "striped-right-arrow" ; "mso-spt100"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP
			$tProp.Value = "up-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_DOWN
			$tProp.Value = "up-down-arrow"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT
			$tProp.Value = "up-right-arrow-callout" ; "mso-spt89"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN
			$tProp.Value = "up-right-down-arrow" ; "mso-spt100"

			$tProp2 = __LO_SetPropertyValue("MirroredX", True) ; Shape is an up and left arrow without this Property.
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

			ReDim $atCusShapeGeo[2]
			$atCusShapeGeo[1] = $tProp2

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_CHEVRON
			$tProp.Value = "chevron"

		Case $LOI_DRAWSHAPE_TYPE_ARROWS_PENTAGON
			$tProp.Value = "pentagon-right"
	EndSwitch

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateArrow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateBasic
; Description ...: Create a Basic type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateBasic(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (100th MM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (100th MM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iShapeType          - an integer value (26-49). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_BASIC_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" or "com.sun.star.drawing.EllipseShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, $LOI_DRAWSHAPE_TYPE_BASIC_FRAME
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateBasic(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]
	Local $iCircleKind_CUT = 2 ; a circle with a cut connected by a line.
	Local $iCircleKind_ARC = 3 ; a circle with an open cut.

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iShapeType = $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT) Or ($iShapeType = $LOI_DRAWSHAPE_TYPE_BASIC_ARC) Then ; These two shapes need special procedures.
		$oShape = $oDoc.createInstance("com.sun.star.drawing.EllipseShape")
		If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Switch $iShapeType
			Case $LOI_DRAWSHAPE_TYPE_BASIC_ARC
				$oShape.Name = __LOImpress_GetShapeName($oSlide, "Elliptical arc ")
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Case $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT
				$oShape.Name = __LOImpress_GetShapeName($oSlide, "Ellipse Segment ")
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndSwitch

		$oSlide.add($oShape)

	Else
		$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
		If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tProp = __LO_SetPropertyValue("Type", "")
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oSlide.add($oShape)
	EndIf

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_BASIC_ARC
			$oShape.FillColor = $LO_COLOR_OFF

			$oShape.CircleKind = $iCircleKind_ARC
			$oShape.CircleStartAngle = 0
			$oShape.CircleEndAngle = 25000

		Case $LOI_DRAWSHAPE_TYPE_BASIC_ARC_BLOCK
			$tProp.Value = "block-arc"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE
			$tProp.Value = "circle-pie" ; "mso-spt100"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT
			$oShape.CircleKind = $iCircleKind_CUT
			$oShape.CircleStartAngle = 0
			$oShape.CircleEndAngle = 25000

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CROSS
			$tProp.Value = "cross"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CUBE
			$tProp.Value = "cube"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_CYLINDER
			$tProp.Value = "can"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_DIAMOND
			$tProp.Value = "diamond"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE
			$tProp.Value = "ellipse"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_FOLDED_CORNER
			$tProp.Value = "paper"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_FRAME
			$tProp.Value = "frame" ; Not working

		Case $LOI_DRAWSHAPE_TYPE_BASIC_HEXAGON
			$tProp.Value = "hexagon"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_OCTAGON
			$tProp.Value = "octagon"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_PARALLELOGRAM
			$tProp.Value = "parallelogram"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE
			$tProp.Value = "rectangle"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED
			$tProp.Value = "round-rectangle"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_REGULAR_PENTAGON
			$tProp.Value = "pentagon"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_RING
			$tProp.Value = "ring"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_TRAPEZOID
			$tProp.Value = "trapezoid"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES
			$tProp.Value = "isosceles-triangle"

		Case $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT
			$tProp.Value = "right-triangle"
	EndSwitch

	If ($iShapeType <> $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT) And ($iShapeType <> $LOI_DRAWSHAPE_TYPE_BASIC_ARC) Then
		$atCusShapeGeo[0] = $tProp
		$oShape.CustomShapeGeometry = $atCusShapeGeo

		; Settings for TextBox use.
		$oShape.TextMinimumFrameWidth = $iWidth
		$oShape.TextMinimumFrameHeight = $iHeight
		$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
		$oShape.TextAutoGrowHeight = False
		$oShape.TextAutoGrowWidth = False
	EndIf

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateBasic

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateCallout
; Description ...: Create a Callout type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateCallout(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (100th MM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (100th MM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iShapeType          - an integer value (50-56). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_CALLOUT_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateCallout(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_CLOUD
			$tProp.Value = "cloud-callout"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_1
			$tProp.Value = "line-callout-1"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_2
			$tProp.Value = "line-callout-2"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_3
			$tProp.Value = "line-callout-3"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR
			$tProp.Value = "rectangular-callout"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED
			$tProp.Value = "round-rectangular-callout"

		Case $LOI_DRAWSHAPE_TYPE_CALLOUT_ROUND
			$tProp.Value = "round-callout"
	EndSwitch

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateCallout

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateFlowchart
; Description ...: Create a FlowChart type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateFlowchart(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (100th MM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (100th MM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iShapeType          - an integer value (57-84). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_FLOWCHART_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateFlowchart(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_CARD
			$tProp.Value = "flowchart-card"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_COLLATE
			$tProp.Value = "flowchart-collate"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR
			$tProp.Value = "flowchart-connector"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE
			$tProp.Value = "flowchart-off-page-connector"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DATA
			$tProp.Value = "flowchart-data"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DECISION
			$tProp.Value = "flowchart-decision"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DELAY
			$tProp.Value = "flowchart-delay"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE
			$tProp.Value = "flowchart-direct-access-storage"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DISPLAY
			$tProp.Value = "flowchart-display"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_DOCUMENT
			$tProp.Value = "flowchart-document"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_EXTRACT
			$tProp.Value = "flowchart-extract"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE
			$tProp.Value = "flowchart-internal-storage"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MAGNETIC_DISC
			$tProp.Value = "flowchart-magnetic-disk"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_INPUT
			$tProp.Value = "flowchart-manual-input"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_OPERATION
			$tProp.Value = "flowchart-manual-operation"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MERGE
			$tProp.Value = "flowchart-merge"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_MULTIDOCUMENT
			$tProp.Value = "flowchart-multidocument"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_OR
			$tProp.Value = "flowchart-or"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PREPARATION
			$tProp.Value = "flowchart-preparation"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS
			$tProp.Value = "flowchart-process"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE
			$tProp.Value = "flowchart-alternate-process"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED
			$tProp.Value = "flowchart-predefined-process"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_PUNCHED_TAPE
			$tProp.Value = "flowchart-punched-tape"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS
			$tProp.Value = "flowchart-sequential-access"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_SORT
			$tProp.Value = "flowchart-sort"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_STORED_DATA
			$tProp.Value = "flowchart-stored-data"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION
			$tProp.Value = "flowchart-summing-junction"

		Case $LOI_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR
			$tProp.Value = "flowchart-terminator"
	EndSwitch

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateFlowchart

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateLine
; Description ...: Create a Line type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateLine(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (100th MM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (100th MM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iShapeType          - an integer value (85-92). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_LINE_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create the requested Line type Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a Position structure.
;                  @Error 2 @Extended 3 Return 0 = Failed to create "com.sun.star.drawing.PolyPolygonBezierCoords" Structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateLine(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tSize, $tPos, $tPolyCoords, $tStart, $tEnd
	Local $atPoint[0], $aiFlags[0]
	Local $avArray[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iShapeType <> $LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIMENSION_LINE) Then
		$tPolyCoords = __LO_CreateStruct("com.sun.star.drawing.PolyPolygonBezierCoords")
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)
	EndIf

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS To $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_STARTS_ARROW
			$oShape = $oDoc.createInstance("com.sun.star.drawing.LineShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Line ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[2]
			ReDim $aiFlags[2]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			Switch $iShapeType
				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_CIRCLE
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Circle"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_SQUARE
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Square"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_CIRCLE_ARROW
					$oShape.LineStartName = "Circle"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ENDS_ARROW
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_SQUARE_ARROW
					$oShape.LineStartName = "Square"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_STARTS_ARROW
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
			EndSwitch

		Case $LOI_DRAWSHAPE_TYPE_CONNECTOR To $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW
			$oShape = $oDoc.createInstance("com.sun.star.drawing.ConnectorShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Connector ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			$tStart = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($tStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$tEnd = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($tEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			Switch $iShapeType
				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_ARROWS
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_CURVE

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ARROWS
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_CURVE
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ENDS_ARROW
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_CURVE
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_ENDS_ARROW
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ARROWS
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ENDS_ARROW
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STRAIGHT

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ARROWS
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STRAIGHT
					$oShape.LineStartName = "Arrow"
					$oShape.LineStartWidth = 300
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300

				Case $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW
					$oShape.EdgeKind = $LOI_DRAWSHAPE_CONNECTOR_TYPE_STRAIGHT
					$oShape.LineEndName = "Arrow"
					$oShape.LineEndWidth = 300
			EndSwitch

		Case $LOI_DRAWSHAPE_TYPE_LINE_CURVE
			$oShape = $oDoc.createInstance("com.sun.star.drawing.OpenBezierShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[4]
			ReDim $aiFlags[4]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight / 2))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint(Int($iX + $iWidth), $iY)
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED
			$oShape = $oDoc.createInstance("com.sun.star.drawing.ClosedBezierShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[4]
			ReDim $aiFlags[4]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight / 2))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint(Int($iX + $iWidth), $iY)
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = 7512015 ; Light blue

		Case $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION
			$oShape = $oDoc.createInstance("com.sun.star.drawing.MeasureShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oSlide.add($oShape)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Dimension Line ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tStart = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($tStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$tEnd = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($tEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		Case $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE
			$oShape = $oDoc.createInstance("com.sun.star.drawing.OpenFreeHandShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[3]
			ReDim $aiFlags[3]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth / 2), Int($iY + $iHeight / 2))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

		Case $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED
			$oShape = $oDoc.createInstance("com.sun.star.drawing.ClosedFreeHandShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[4]
			ReDim $aiFlags[4]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth) + Int(($iX + $iWidth / 8)), Int(($iY + $iHeight / 2)))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = 7512015 ; Light blue

		Case $LOI_DRAWSHAPE_TYPE_LINE_LINE, $LOI_DRAWSHAPE_TYPE_LINE_LINE_45
			$oShape = $oDoc.createInstance("com.sun.star.drawing.LineShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Line ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[2]
			ReDim $aiFlags[2]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

		Case $LOI_DRAWSHAPE_TYPE_LINE_POLYGON, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45
			$oShape = $oDoc.createInstance("com.sun.star.drawing.PolyLineShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Polygon 4 corners ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[5]
			ReDim $aiFlags[5]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth), $iY)
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint($iX, Int($iY + $iHeight))
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[4] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[4]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[4] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED
			$oShape = $oDoc.createInstance("com.sun.star.drawing.PolyPolygonShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oShape.Name = __LOImpress_GetShapeName($oSlide, "Polygon 4 corners ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oSlide.add($oShape)

			ReDim $atPoint[5]
			ReDim $aiFlags[5]

			$atPoint[0] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[0]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOImpress_CreatePoint(Int($iX + $iWidth), $iY)
			If Not IsObj($atPoint[1]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOImpress_CreatePoint(Int($iX + $iWidth), Int($iY + $iHeight))
			If Not IsObj($atPoint[2]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOImpress_CreatePoint($iX, Int($iY + $iHeight))
			If Not IsObj($atPoint[3]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[4] = __LOImpress_CreatePoint($iX, $iY)
			If Not IsObj($atPoint[4]) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[2] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[3] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL
			$aiFlags[4] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

			$oShape.FillColor = 7512015 ; Light blue
	EndSwitch

	If __LO_IntIsBetween($iShapeType, $LOI_DRAWSHAPE_TYPE_CONNECTOR, $LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW, "", $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION) Then
		$oShape.StartPosition = $tStart
		$oShape.EndPosition = $tEnd

	Else
		$avArray[0] = $atPoint
		$tPolyCoords.Coordinates = $avArray

		$avArray[0] = $aiFlags
		$tPolyCoords.Flags = $avArray

		$oShape.PolyPolygonBezier = $tPolyCoords
	EndIf

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateStars
; Description ...: Create a Star or Banner type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateStars(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (100th MM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (100th MM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iShapeType          - an integer value (93-104). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_STARS_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOI_DRAWSHAPE_TYPE_STARS_6_POINT, $LOI_DRAWSHAPE_TYPE_STARS_12_POINT, $LOI_DRAWSHAPE_TYPE_STARS_SIGNET, $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateStars(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_STARS_4_POINT
			$tProp.Value = "star4"

		Case $LOI_DRAWSHAPE_TYPE_STARS_5_POINT
			$tProp.Value = "star5"

		Case $LOI_DRAWSHAPE_TYPE_STARS_6_POINT
			$tProp.Value = "star6" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE
			$tProp.Value = "concave-star6" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_STARS_8_POINT
			$tProp.Value = "star8"

		Case $LOI_DRAWSHAPE_TYPE_STARS_12_POINT
			$tProp.Value = "star12" ; "non-primitive"

		Case $LOI_DRAWSHAPE_TYPE_STARS_24_POINT
			$tProp.Value = "star24"

		Case $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE
			$tProp.Value = "mso-spt21" ; "doorplate"

		Case $LOI_DRAWSHAPE_TYPE_STARS_EXPLOSION
			$tProp.Value = "bang"

		Case $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_HORIZONTAL
			$tProp.Value = "horizontal-scroll"

		Case $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_VERTICAL
			$tProp.Value = "vertical-scroll"

		Case $LOI_DRAWSHAPE_TYPE_STARS_SIGNET
			$tProp.Value = "signet" ; "non-primitive"
	EndSwitch

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateStars

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_CreateSymbol
; Description ...: Create a Symbol type Shape.
; Syntax ........: __LOImpress_DrawShape_CreateSymbol(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
; Parameters ....: $oSlide              - [in/out] an object. A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetByIndex, or _LOImpress_SlideCopy function.
;                  $iWidth              - an integer value. The Shape's Width in Hundredths of a Millimeter (100th MM).
;                  $iHeight             - an integer value. The Shape's Height in Hundredths of a Millimeter (100th MM).
;                  $iX                  - an integer value. The X position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iY                  - an integer value. The Y position from the insertion point, in Hundredths of a Millimeter (100th MM).
;                  $iShapeType          - an integer value (105-122). The Type of shape to create. See $LOI_DRAWSHAPE_TYPE_SYMBOL_* as defined in LibreOfficeImpress_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iShapeType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD, $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER, $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
;                  The following shape is visually different from the manually inserted one in L.O. 7.3.4.2:
;                  $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_CreateSymbol(ByRef $oSlide, $iWidth, $iHeight, $iX, $iY, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oShape.Name = __LOImpress_GetShapeName($oSlide, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oSlide.add($oShape)

	Switch $iShapeType
		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
			$tProp.Value = "col-502ad400"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON
			$tProp.Value = "col-60da8460"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_SQUARE
			$tProp.Value = "quad-bevel"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_DOUBLE
			$tProp.Value = "brace-pair"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_LEFT
			$tProp.Value = "left-brace"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_RIGHT
			$tProp.Value = "right-brace"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_DOUBLE
			$tProp.Value = "bracket-pair"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_LEFT
			$tProp.Value = "left-bracket"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_RIGHT
			$tProp.Value = "right-bracket"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "cloud"
			$tProp.Value = "cloud"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "flower"
			$tProp.Value = "flower"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_HEART
			$tProp.Value = "heart"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "lightning"
			$tProp.Value = "lightning"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_MOON
			$tProp.Value = "moon"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_SMILEY
			$tProp.Value = "smiley"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_SUN
			$tProp.Value = "sun"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_PROHIBITED
			$tProp.Value = "forbidden"

		Case $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE
			$tProp.Value = "puzzle"
	EndSwitch

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	; Settings for TextBox use.
	$oShape.TextMinimumFrameWidth = $iWidth
	$oShape.TextMinimumFrameHeight = $iHeight
	$oShape.TextVerticalAdjust = $LOI_ALIGN_VERT_MIDDLE
	$oShape.TextAutoGrowHeight = False
	$oShape.TextAutoGrowWidth = False

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOImpress_DrawShape_CreateSymbol

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_GetCustomType
; Description ...: Return the Shape Type Constant corresponding to the Custom Shape Type string.
; Syntax ........: __LOImpress_DrawShape_GetCustomType($sCusShapeType)
; Parameters ....: $sCusShapeType       - a string value. The Returned Custom Shape Type Value from CustomShapeGeometry Array of properties.
; Return values .: Success: Integer or -1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sCusShapeType not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Custom Shape Type was successfully identified. Returning the Constant value of the Shape, see Constants $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 0 @Extended 0 Return -1 = Success. Custom Shape is of an unimplemented type that has an ambiguous name, and cannot be identified. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Some shapes are not implemented, or not fully implemented into LibreOffice for automation, consequently they do not have appropriate type names as of yet. Many have simply ambiguous names, such as "non-primitive".
;                  #1 Because of this the following shape types cannot be identified, and this function will return -1:
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
;                  #3 The following Shapes have strange names that may change in the future, but currently are able to be identified:
;                  - $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE, known as, "mso-spt21", should be "doorplate"
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, known as, "col-502ad400", should be ??
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, known as, "col-60da8460", should be ??
;                  #4 The following Shapes are customizable one to another, and are consequently indistinguishable:
;                  - $LOI_DRAWSHAPE_TYPE_FONTWORK_* (The Value of $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL is returned for any of these.)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_GetCustomType($sCusShapeType)
	If Not IsString($sCusShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $sCusShapeType
		Case "quad-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY)

		Case "quad-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY)

		Case "down-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN)

		Case "left-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT)

		Case "left-right-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT)

		Case "right-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT)

		Case "up-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP)

		Case "up-down-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN)

;~ 	Case "mso-spt100" ; Can't include this one as other shapes return mso-spt100 also
;~ Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT)

		Case "circular-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CIRCULAR)

		Case "corner-right-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT)

		Case "down-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_DOWN)

		Case "left-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT)

		Case "left-right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT)

		Case "notched-right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT)

		Case "right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT)

		Case "right-left-arrow" ; "non-primitive"??

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT)

		Case "s-sharped-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED)

		Case "split-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT)

		Case "striped-right-arrow" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT)

		Case "up-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP)

		Case "up-down-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_DOWN)

		Case "up-right-arrow-callout", "mso-spt89" ; "mso-spt89"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT)

		Case "up-right-down-arrow" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN)

		Case "chevron"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_CHEVRON)

		Case "pentagon-right"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_PENTAGON)

		Case "block-arc"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_ARC_BLOCK)

		Case "circle-pie" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE)

		Case "cross"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CROSS)

		Case "cube"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CUBE)

		Case "can"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CYLINDER)

		Case "diamond"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_DIAMOND)

		Case "ellipse"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE)
;~ $LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE

		Case "paper"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_FOLDED_CORNER)

		Case "frame" ; Not working

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_FRAME)

		Case "hexagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_HEXAGON)

		Case "octagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_OCTAGON)

		Case "parallelogram"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_PARALLELOGRAM)

		Case "rectangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE)
;~ $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE

		Case "round-rectangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED)
;~ $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED

		Case "pentagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_REGULAR_PENTAGON)

		Case "ring"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_RING)

		Case "trapezoid"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_TRAPEZOID)

		Case "isosceles-triangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES)

		Case "right-triangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT)

		Case "cloud-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_CLOUD)

		Case "line-callout-1"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_1)

		Case "line-callout-2"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_2)

		Case "line-callout-3"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_3)

		Case "rectangular-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR)

		Case "round-rectangular-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED)

		Case "round-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_ROUND)

		Case "flowchart-card"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_CARD)

		Case "flowchart-collate"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_COLLATE)

		Case "flowchart-connector"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR)

		Case "flowchart-off-page-connector"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE)

		Case "flowchart-data"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DATA)

		Case "flowchart-decision"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DECISION)

		Case "flowchart-delay"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DELAY)

		Case "flowchart-direct-access-storage"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE)

		Case "flowchart-display"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DISPLAY)

		Case "flowchart-document"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DOCUMENT)

		Case "flowchart-extract"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_EXTRACT)

		Case "flowchart-internal-storage"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE)

		Case "flowchart-magnetic-disk"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MAGNETIC_DISC)

		Case "flowchart-manual-input"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_INPUT)

		Case "flowchart-manual-operation"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_OPERATION)

		Case "flowchart-merge"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MERGE)

		Case "flowchart-multidocument"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MULTIDOCUMENT)

		Case "flowchart-or"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_OR)

		Case "flowchart-preparation"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PREPARATION)

		Case "flowchart-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS)

		Case "flowchart-alternate-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE)

		Case "flowchart-predefined-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED)

		Case "flowchart-punched-tape"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PUNCHED_TAPE)

		Case "flowchart-sequential-access"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS)

		Case "flowchart-sort"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_SORT)

		Case "flowchart-stored-data"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_STORED_DATA)

		Case "flowchart-summing-junction"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION)

		Case "flowchart-terminator"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR)

		Case "fontwork-arch-down-pour", "fontwork-arch-left-pour", "fontwork-arch-right-pour", "fontwork-arch-up-pour", "fontwork-arch-down-curve", _
				"fontwork-arch-left-curve", "fontwork-arch-right-curve", "fontwork-arch-up-curve", "fontwork-chevron-down", "fontwork-chevron-up", _
				"fontwork-circle-curve", "fontwork-circle-pour", "fontwork-curve-down", "fontwork-curve-up", "fontwork-fade-down", "fontwork-fade-left", _
				"fontwork-fade-right", "fontwork-fade-up", "fontwork-fade-up-and-left", "fontwork-fade-up-and-right", "fontwork-inflate", "fontwork-slant-down", _
				"fontwork-slant-up", "fontwork-stop", "fontwork-triangle-up", "fontwork-triangle-down", "fontwork-open-circle-curve", "fontwork-open-circle-pour", _
				"fontwork-plain-text", "fontwork-wave"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL) ; Can't differentiate reliably

		Case "star4"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_4_POINT)

		Case "star5"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_5_POINT)

		Case "star6" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_6_POINT)

		Case "concave-star6" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE)

		Case "star8"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_8_POINT)

		Case "star12" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_12_POINT)

		Case "star24"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_24_POINT)

		Case "mso-spt21", "doorplate" ; "doorplate"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE)

		Case "bang"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_EXPLOSION)

		Case "horizontal-scroll"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_HORIZONTAL)

		Case "vertical-scroll"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_VERTICAL)

		Case "signet" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_SIGNET)

		Case "col-502ad400"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND)

		Case "col-60da8460"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON)

		Case "quad-bevel"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_SQUARE)

		Case "brace-pair"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_DOUBLE)

		Case "left-brace"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_LEFT)

		Case "right-brace"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_RIGHT)

		Case "bracket-pair"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_DOUBLE)

		Case "left-bracket"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_LEFT)

		Case "right-bracket"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_RIGHT)

		Case "cloud"
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "cloud"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD)

		Case "flower"
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "flower"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER)

		Case "heart"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_HEART)

		Case "lightning"
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "lightning"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING)

		Case "moon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_MOON)

		Case "smiley"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_SMILEY)

		Case "sun"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_SUN)

		Case "forbidden"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_PROHIBITED)

		Case "puzzle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE)

		Case Else

			Return SetError($__LO_STATUS_SUCCESS, 0, -1)
	EndSwitch
EndFunc   ;==>__LOImpress_DrawShape_GetCustomType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShapeArrowStyleName
; Description ...: Convert a Arrow head Constant to the corresponding name or reverse.
; Syntax ........: __LOImpress_DrawShapeArrowStyleName([$iArrowStyle = Null[, $sArrowStyle = Null]])
; Parameters ....: $iArrowStyle         - [optional] an integer value (0-32). Default is Null. The Arrow Style Constant to convert to its corresponding name. See $LOI_DRAWSHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $sArrowStyle         - [optional] a string value. Default is Null. The Arrow Style Name to convert to the corresponding constant if found.
; Return values .: Success: String or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iArrowStyle not an Integer, less than 0 or greater than Arrow type constants. See $LOI_DRAWSHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 2 Return 0 = $sArrowStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = Both $iArrowStyle and $sArrowStyle called with Null.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Constant called in $iArrowStyle was successfully converted to its corresponding Arrow Type Name.
;                  @Error 0 @Extended 1 Return Integer = Success. Arrow Type Name called in $sArrowStyle was successfully converted to its corresponding Constant value.
;                  @Error 0 @Extended 2 Return String = Success. Arrow Type Name called in $sArrowStyle was not matched to an existing Constant value, returning called name. Possibly a custom value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShapeArrowStyleName($iArrowStyle = Null, $sArrowStyle = Null)
	Local $asArrowStyles[33]

	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_NONE] = ""
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW_SHORT] = "Arrow short"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CONCAVE_SHORT] = "Concave short"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW] = "Arrow"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_TRIANGLE] = "Triangle"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CONCAVE] = "Concave"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW_LARGE] = "Arrow large"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CIRCLE] = "Circle"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_SQUARE] = "Square"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_SQUARE_45] = "Square 45"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIAMOND] = "Diamond"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_HALF_CIRCLE] = "Half Circle"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINES] = "Dimension Lines"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW] = "Dimension Line Arrow"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIMENSION_LINE] = "Dimension Line"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_LINE_SHORT] = "Line short"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_LINE] = "Line"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_TRIANGLE_UNFILLED] = "Triangle unfilled"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIAMOND_UNFILLED] = "Diamond unfilled"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CIRCLE_UNFILLED] = "Circle unfilled"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED] = "Square 45 unfilled"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_SQUARE_UNFILLED] = "Square unfilled"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED] = "Half Circle unfilled"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_HALF_ARROW_LEFT] = "Half Arrow left"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_HALF_ARROW_RIGHT] = "Half Arrow right"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_REVERSED_ARROW] = "Reversed Arrow"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DOUBLE_ARROW] = "Double Arrow"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ONE] = "CF One"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ONLY_ONE] = "CF Only One"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_MANY] = "CF Many"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_MANY_ONE] = "CF Many One"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ZERO_ONE] = "CF Zero One"
	$asArrowStyles[$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY] = "CF Zero Many"

	If ($iArrowStyle <> Null) Then
		If Not __LO_IntIsBetween($iArrowStyle, 0, UBound($asArrowStyles) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $asArrowStyles[$iArrowStyle]) ; Return the requested Arrow Style name.

	ElseIf ($sArrowStyle <> Null) Then
		If Not IsString($sArrowStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To UBound($asArrowStyles) - 1
			If ($asArrowStyles[$i] = $sArrowStyle) Then Return SetError($__LO_STATUS_SUCCESS, 1, $i) ; Return the array element where the matching Arrow Style was found.

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 2, $sArrowStyle) ; If no matches, just return the name, as it could be a custom value.

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; No values called.
	EndIf
EndFunc   ;==>__LOImpress_DrawShapeArrowStyleName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShapeLineStyleName
; Description ...: Convert a Line Style Constant to the corresponding name or reverse.
; Syntax ........: __LOImpress_DrawShapeLineStyleName([$iLineStyle = Null[, $sLineStyle = Null]])
; Parameters ....: $iLineStyle          - [optional] an integer value (0-31). Default is Null. The Line Style Constant to convert to its corresponding name. See $LOI_DRAWSHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3
;                  $sLineStyle          - [optional] a string value. Default is Null. The Line Style Name to convert to the corresponding constant if found.
; Return values .: Success: String or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iLineStyle not an Integer, less than 0 or greater than Line Style constants. See $LOI_DRAWSHAPE_LINE_STYLE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 1 @Extended 2 Return 0 = $sLineStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = Both $iLineStyle and $sLineStyle called with Null.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Constant called in $iLineStyle was successfully converted to its corresponding Line Style Name.
;                  @Error 0 @Extended 1 Return Integer = Success. Line Style Name called in $sLineStyle was successfully converted to its corresponding Constant value.
;                  @Error 0 @Extended 2 Return String = Success. Line Style Name called in $sLineStyle was not matched to an existing Constant value, returning called name. Possibly a custom value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShapeLineStyleName($iLineStyle = Null, $sLineStyle = Null)
	Local $asLineStyles[32]

	; $LOI_DRAWSHAPE_LINE_STYLE_NONE, $LOI_DRAWSHAPE_LINE_STYLE_CONTINUOUS, don't have a name, so to keep things symmetrical I created my own, but those two won't be used.
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_NONE] = "NONE"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_CONTINUOUS] = "CONTINUOUS"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DOT] = "Dot"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DOT_ROUNDED] = "Dot (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_LONG_DOT] = "Long Dot"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_LONG_DOT_ROUNDED] = "Long Dot (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DASH] = "Dash"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DASH_ROUNDED] = "Dash (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_LONG_DASH] = "Long Dash"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_LONG_DASH_ROUNDED] = "Long Dash (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH] = "Double Dash"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_ROUNDED] = "Double Dash (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DASH_DOT] = "Dash Dot"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DASH_DOT_ROUNDED] = "Dash Dot (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_LONG_DASH_DOT] = "Long Dash Dot"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_LONG_DASH_DOT_ROUNDED] = "Long Dash Dot (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_DOT] = "Double Dash Dot"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED] = "Double Dash Dot (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DASH_DOT_DOT] = "Dash Dot Dot"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DASH_DOT_DOT_ROUNDED] = "Dash Dot Dot (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT] = "Double Dash Dot Dot"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED] = "Double Dash Dot Dot (Rounded)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_ULTRAFINE_DOTTED] = "Ultrafine Dotted (var)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_FINE_DOTTED] = "Fine Dotted"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_ULTRAFINE_DASHED] = "Ultrafine Dashed"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_FINE_DASHED] = "Fine Dashed"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_DASHED] = "Dashed (var)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_SPARSE_DASH] = "Sparse Dash"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_3_DASHES_3_DOTS] = "3 Dashes 3 Dots (var)"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES] = "Ultrafine 2 Dots 3 Dashes"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_2_DOTS_1_DASH] = "2 Dots 1 Dash"
	$asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS] = "Line with Fine Dots"

	If Not __LO_VersionCheck(24.2) Then $asLineStyles[$LOI_DRAWSHAPE_LINE_STYLE_SPARSE_DASH] = "Line Style 9"

	If ($iLineStyle <> Null) Then
		If Not __LO_IntIsBetween($iLineStyle, 0, UBound($asLineStyles) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $asLineStyles[$iLineStyle]) ; Return the requested Line Style name.

	ElseIf ($sLineStyle <> Null) Then
		If Not IsString($sLineStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To UBound($asLineStyles) - 1
			If ($asLineStyles[$i] = $sLineStyle) Then Return SetError($__LO_STATUS_SUCCESS, 1, $i) ; Return the array element where the matching Line Style was found.

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 2, $sLineStyle) ; If no matches, just return the name, as it could be a custom value.

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; No values called.
	EndIf
EndFunc   ;==>__LOImpress_DrawShapeLineStyleName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShapePointGetSettings
; Description ...: Retrieve the current settings for a particular point in a shape.
; Syntax ........: __LOImpress_DrawShapePointGetSettings(ByRef $avArray, ByRef $aiFlags, ByRef $atPoints, $iArrayElement)
; Parameters ....: $avArray             - [in/out] an array of variants. An array to fill with settings. Array will be directly modified.
;                  $aiFlags             - [in/out] an array of integers. An Array of Point Type Flags returned from the Shape.
;                  $atPoints            - [in/out] an array of dll structs. An Array of Points returned from the Shape.
;                  $iArrayElement       - an integer value. The Array element that contains the point to retrieve the settings for.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $avArray is not an Array.
;                  @Error 1 @Extended 2 Return 0 = $aiFlags is not an Array.
;                  @Error 1 @Extended 3 Return 0 = $atPoints is not an Array.
;                  @Error 1 @Extended 4 Return 0 = $iArrayElement not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the current X coordinate.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the current Y coordinate.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the current Point Type Flag.
;                  @Error 3 @Extended 4 Return 0 = Failed to determine if the Point is a Curve or not.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Current Settings were successfully retrieved, $avArray has been filled with the current settings.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShapePointGetSettings(ByRef $avArray, ByRef $aiFlags, ByRef $atPoints, $iArrayElement)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iX, $iY, $iPointType
	Local $bIsCurve

	If Not IsArray($avArray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsArray($atPoints) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iArrayElement) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iX = $atPoints[$iArrayElement].X()
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$avArray[0] = $iX

	$iY = $atPoints[$iArrayElement].Y()
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$avArray[1] = $iY

	$iPointType = $aiFlags[$iArrayElement]
	If Not IsInt($iPointType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$avArray[2] = $iPointType

	If ($iPointType = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then
		If ($iArrayElement <> (UBound($atPoints) - 1)) Then ; Requested point is not at the end of the array of points.

			If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Point after requested point is a Control Point.
				; If a Point and the following Control point have the same coordinates, the point is not a curve.
				$bIsCurve = (($atPoints[$iArrayElement].X() = $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() = $atPoints[$iArrayElement + 1].Y())) ? (False) : (True)

			Else ; Next point after requested point is not a control type point.
				$bIsCurve = False
			EndIf

		Else ; Point is the last point, cant be a curve.
			$bIsCurve = False
		EndIf

	Else ; Point is a Smooth, or Symmetrical Point type, point is a curve regardless.
		$bIsCurve = True
	EndIf

	If Not IsBool($bIsCurve) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$avArray[3] = $bIsCurve

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_DrawShapePointGetSettings

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShapePointModify
; Description ...: Internal function for modifying A Shape's Points.
; Syntax ........: __LOImpress_DrawShapePointModify(ByRef $aiFlags, ByRef $atPoints, ByRef $iArrayElement[, $iX = Null[, $iY = Null[, $iPointType = Null[, $bIsCurve = Null]]]])
; Parameters ....: $aiFlags             - [in/out] an array of integers. An Array of Point Type Flags returned from the Shape. Array will be directly modified.
;                  $atPoints            - [in/out] an array of dll structs. An Array of Points returned from the Shape. Array will be directly modified.
;                  $iArrayElement       - [in/out] an integer value. The Array element that contains the point to modify. This may be directly modified, depending on the settings.
;                  $iX                  - [optional] an integer value. Default is Null. The X coordinate value, set in Hundredths of a Millimeter (100th MM).
;                  $iY                  - [optional] an integer value. Default is Null. The Y coordinate value, set in Hundredths of a Millimeter (100th MM).
;                  $iPointType          - [optional] an integer value (0,1,3). Default is Null. The Type of Point to change the called point to. See Remarks. See constants $LOI_DRAWSHAPE_POINT_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  $bIsCurve            - [optional] a boolean value. Default is Null. If True, the Normal Point is a Curve. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $aiFlags not an Array.
;                  @Error 1 @Extended 2 Return 0 = $atPoints not an Array.
;                  @Error 1 @Extended 3 Return 0 = $iArrayElement not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a new Position Point Structure for the First Control Point.
;                  @Error 2 @Extended 2 Return 0 = Failed to Create a new Position Point Structure for the Second Control Point.
;                  @Error 2 @Extended 3 Return 0 = Failed to Create a new Position Point Structure for the Third Control Point.
;                  @Error 2 @Extended 4 Return 0 = Failed to Create a new Position Point Structure for the Fourth Control Point.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify the next position point in the shape.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify the previous position point in the shape.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;                  Only $LOI_DRAWSHAPE_TYPE_LINE_* type shapes have Points that can be added to, removed, or modified.
;                  This is a homemade function as LibreOffice doesn't offer an easy way for modifying points in a shape. Consequently this will not produce similar results as when working with Libre office manually, and may wreck your shape's shape. Use with caution.
;                  For an unknown reason, I am unable to insert "SMOOTH" Points, and consequently, any smooth Points are reverted back to "Normal" points, but still having their Smooth control points upon insertion that were already present in the shape. If you modify a point to "SMOOTH" type, it will be, for now, replaced with "Symmetrical".
;                  The first and last points in a shape can only be a "Normal" Point Type. The last point cannot be Curved, but the first can be.
;                  Calling and Smooth or Symmetrical point types with $bIsCurve = True, will be ignored, as they are already a curve.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShapePointModify(ByRef $aiFlags, ByRef $atPoints, ByRef $iArrayElement, $iX = Null, $iY = Null, $iPointType = Null, $bIsCurve = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iNextArrayElement, $iPreviousArrayElement, $iSymmetricalPointXValue, $iSymmetricalPointYValue, $iOffset, $iForOffset, $iReDimCount
	Local $tControlPoint1, $tControlPoint2, $tControlPoint3, $tControlPoint4
	Local $avArray[0], $avArray2[0]

	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($atPoints) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iArrayElement) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Error if point called is not between 0 or number of points.

	If ($iArrayElement <> UBound($atPoints) - 1) Then ; If The requested point to be modified is not at the end of the Array of points, find the next regular point.

		For $i = ($iArrayElement + 1) To UBound($aiFlags) - 1 ; Locate the next non-Control Point in the Array for later use.
			If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$iNextArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		If Not IsInt($iNextArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else
		$iNextArrayElement = -1
	EndIf

	If ($iArrayElement > 0) Then ; If Point requested is not the first point, find the previous Point's position.

		For $i = ($iArrayElement - 1) To 0 Step -1 ; Locate the previous non-Control Point in the Array for later use.
			If ($aiFlags[$i] <> $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then
				$iPreviousArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		Next

		If Not IsInt($iPreviousArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Else
		$iPreviousArrayElement = -1
	EndIf

	If ($iX <> Null) Then
		If ($iArrayElement < UBound($atPoints) - 1) And ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Next point is a control point, check if this point is a curve.

			If ($atPoints[$iArrayElement].X() = $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() = $atPoints[$iArrayElement + 1].Y()) Then ; Update the coordinates, because the point is not a curve.
				$atPoints[$iArrayElement + 1].X = $iX
			EndIf
		EndIf

		$atPoints[$iArrayElement].X = $iX
	EndIf

	If ($iY <> Null) Then
		If ($iArrayElement < UBound($atPoints) - 1) And ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Next point is a control point, check if this point is a curve.

			If ($atPoints[$iArrayElement].X() = $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() = $atPoints[$iArrayElement + 1].Y()) Then ; Update the coordinates, because the point is not a curve.
				$atPoints[$iArrayElement + 1].Y = $iY
			EndIf
		EndIf

		$atPoints[$iArrayElement].Y = $iY
	EndIf

	If ($iPointType <> Null) Then
		If ($iPointType <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; New point type is a curve.

			If ($aiFlags[$iArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; Converting point from Normal to a curve.

				; Pick the lowest X and Y value difference between previous point and current point and Next point and current Point.
				$iSymmetricalPointXValue = ((($atPoints[$iArrayElement].X() - $atPoints[$iPreviousArrayElement].X()) * .5) < (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)) ? Int((($atPoints[$iArrayElement].X() - $atPoints[$iPreviousArrayElement].X()) * .5)) : Int((($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5))
				$iSymmetricalPointYValue = ((($atPoints[$iArrayElement].Y() - $atPoints[$iPreviousArrayElement].Y()) * .5) < (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)) ? Int((($atPoints[$iArrayElement].Y() - $atPoints[$iPreviousArrayElement].Y()) * .5)) : Int((($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5))

				If ($aiFlags[$iArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; previous point is a control Point, might just need to modify it.

					If (($iArrayElement - 2 > $iPreviousArrayElement) And $aiFlags[$iArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; there are two control points before this point, I can just modify the first point before.
						$tControlPoint1 = $atPoints[$iArrayElement - 2]

						$tControlPoint2 = $atPoints[$iArrayElement - 1]
						$tControlPoint2.X = ($atPoints[$iArrayElement].X() - $iSymmetricalPointXValue)
						$tControlPoint2.Y = ($atPoints[$iArrayElement].Y() - $iSymmetricalPointYValue)

					Else ; There is only one control point, I need to create a new one.
						$tControlPoint1 = $atPoints[$iArrayElement - 1]

						$tControlPoint2 = __LOImpress_CreatePoint($atPoints[$iArrayElement].X() - $iSymmetricalPointXValue, $atPoints[$iArrayElement].Y() - $iSymmetricalPointYValue)
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
					EndIf

				Else ; Previous point is a normal point, need to create new control points.
					$tControlPoint1 = __LOImpress_CreatePoint($atPoints[$iPreviousArrayElement].X(), $atPoints[$iPreviousArrayElement].Y())
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tControlPoint2 = __LOImpress_CreatePoint($atPoints[$iArrayElement].X() - $iSymmetricalPointXValue, $atPoints[$iArrayElement].Y() - $iSymmetricalPointYValue)
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
				EndIf

				If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Next point is a control Point, might just need to modify it.

					If (($iArrayElement + 2 < $iNextArrayElement) And $aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; there are two control points after this point, I can just modify the first point after.
						$tControlPoint4 = $atPoints[$iArrayElement + 2]

						$tControlPoint3 = $atPoints[$iArrayElement + 1]
						$tControlPoint3.X = ($atPoints[$iArrayElement].X() + $iSymmetricalPointXValue)
						$tControlPoint3.Y = ($atPoints[$iArrayElement].Y() + $iSymmetricalPointYValue)

					Else ; There is only one control point, I need to create a new one and modify the other.
						$tControlPoint3 = __LOImpress_CreatePoint($atPoints[$iArrayElement].X() + $iSymmetricalPointXValue, $atPoints[$iArrayElement].Y() + $iSymmetricalPointYValue)
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

						$tControlPoint4 = $atPoints[$iArrayElement + 1] ; Modify the Control Point.
						$tControlPoint4.X = ($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5))
						$tControlPoint4.Y = ($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5))
					EndIf

				Else ; Next point is a normal point, need to create new control points.
					$tControlPoint3 = __LOImpress_CreatePoint(($atPoints[$iArrayElement].X() + $iSymmetricalPointXValue), ($atPoints[$iArrayElement].Y() + $iSymmetricalPointYValue))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

					$tControlPoint4 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 4
				; Check if there already was 4 control point present around this point I am modifying.
				$iReDimCount -= ($aiFlags[$iArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0)
				$iReDimCount -= (($iArrayElement - 2 > $iPreviousArrayElement) And ($aiFlags[$iArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)
				$iReDimCount -= ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0)
				$iReDimCount -= (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) + $iReDimCount]
				ReDim $avArray2[UBound($aiFlags) + $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the points to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iPreviousArrayElement) Then ; Insert the new or modified control points.

						$avArray[$i + 1] = $tControlPoint1
						$avArray2[$i + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$avArray[$i + 2] = $tControlPoint2
						$avArray2[$i + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$avArray[$i + 3] = $atPoints[$iArrayElement]
						$avArray2[$i + 3] = $iPointType
						$avArray[$i + 4] = $tControlPoint3
						$avArray2[$i + 4] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL
						$avArray[$i + 5] = $tControlPoint4
						$avArray2[$i + 5] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						$iOffset = 1 ; Add one to offset to skip the point I am modifying.
						$iOffset += ($aiFlags[$iArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am modifying has a control point before it, I need to skip them in the PointsArray.
						$iOffset += (($iArrayElement - 2 > $iPreviousArrayElement) And ($aiFlags[$iArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0) ; If the point I am modifying has two control points before it, I need to skip them in the PointsArray.
						$iOffset += ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am modifying has a control point after it, I need to skip them in the PointsArray.
						$iOffset += (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0) ; If the point I am modifying has two control points after it, I need to skip them in the PointsArray.

						$iForOffset += 5 ; Add to $i to skip the elements I manually added.
					EndIf

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
				Next

				; Update the ArrayElement value to its new position.
				$iArrayElement += ($aiFlags[$iArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (0) : (1) ; If the point I am modifying has a control point before it, don't add one to array element, because I didn't have to create and insert a new control point.
				$iArrayElement += (($iArrayElement - 2 > $iPreviousArrayElement) And ($aiFlags[$iArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (0) : (1) ; If the point I am modifying has two control points before it, don't add one to array element, because I didn't have to create and insert a new control point.

				$atPoints = $avArray
				$aiFlags = $avArray2

			Else ; Point is already a curve.
				; Do nothing?
			EndIf

		Else ; New Point is a Normal Point.
			If ($aiFlags[$iArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; Point being modified is not a normal type of point.

				If ($aiFlags[$iPreviousArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; If previous point is a normal point, see if I need to delete control points or not.

					If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Point after previous point is a control point, see if previous point is a curved point.

						If ($atPoints[$iPreviousArrayElement].X() <> $atPoints[$iPreviousArrayElement + 1].X()) And ($atPoints[$iPreviousArrayElement].Y() <> $atPoints[$iPreviousArrayElement + 1].Y()) Then
							; Previous Point is a Curved normal point, copy the control points present.

							$tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]

							If ($iPreviousArrayElement + 2 < $iArrayElement) And ($atPoints[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2] ; If two control points are present, copy them.
						EndIf
					EndIf

				Else ; Previous point is not a normal point.
					; Copy Control Points present.

					If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]

					If ($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2] ; If two control points are present, copy them.
				EndIf

				If ($aiFlags[$iNextArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then
					; Next point is a curve of some form, copy the control points.

					If ($aiFlags[$iNextArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint4 = $atPoints[$iNextArrayElement - 1]

					If ($iNextArrayElement - 2 > $iArrayElement) And ($aiFlags[$iNextArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint3 = $atPoints[$iNextArrayElement - 2] ; If two control points are present, copy them.
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 4
				; Check how many control points I am keeping.
				$iReDimCount -= (IsObj($tControlPoint1)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint2)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint3)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint4)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) - $iReDimCount]
				ReDim $avArray2[UBound($aiFlags) - $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i + $iOffset] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i + $iOffset] ; Add the rest of the points to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iPreviousArrayElement) Then ; Insert the old control points or remove them.

						If IsObj($tControlPoint1) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint1
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						If IsObj($tControlPoint2) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint2
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If (($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						$iForOffset += 1
						$iOffset += 1
						$avArray[$i + $iForOffset] = $atPoints[$iArrayElement]
						$avArray2[$i + $iForOffset] = $iPointType

						If IsObj($tControlPoint3) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint3
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iNextArrayElement - 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1
						EndIf

						If IsObj($tControlPoint3) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint4
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If (($iNextArrayElement - 2 > $iArrayElement) And ($aiFlags[$iNextArrayElement - 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1
						EndIf
					EndIf

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
				Next

				; Update the ArrayElement value to its new position.
				$iArrayElement -= (IsObj($tControlPoint1)) ? (0) : (1) ; If ControlPoint 1 is a object, it means I copied it, meaning I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.
				$iArrayElement -= (IsObj($tControlPoint2)) ? (0) : (1) ; If ControlPoint 2 is a object, it means I copied it, meaning I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.

				$atPoints = $avArray
				$aiFlags = $avArray2

			Else ; Point being modified is a normal point already.
				; Do nothing?
			EndIf
		EndIf
	EndIf

	If ($bIsCurve <> Null) Then
		If ($aiFlags[$iArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; If Point to modify is a normal point, then proceed, else point is a curve already.

			If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then ; Point after point to modify is a control point, just modify it.
				$tControlPoint3 = $atPoints[$iArrayElement + 1]

				If ($bIsCurve = True) Then
					$tControlPoint3.X = ($atPoints[$iArrayElement].X() + (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5))
					$tControlPoint3.Y = ($atPoints[$iArrayElement].Y() + (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5))

					If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then
						$tControlPoint4 = $atPoints[$iArrayElement + 2] ; Copy the second control point.

					Else ; Create a new control point.
						$tControlPoint4 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
					EndIf

				ElseIf ($bIsCurve = False) And ($aiFlags[$iNextArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; Next point is a curve, so just modify the control point.
					$tControlPoint3.X = $atPoints[$iArrayElement].X() ; When the control point after a point has the same coordinates, it means it is not a curve.
					$tControlPoint3.Y = $atPoints[$iArrayElement].Y()
					; Copy the second control point if it exists.
					If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $tControlPoint4 = $atPoints[$iArrayElement + 2]

				Else ; IsCurve = False, and next point is normal. delete control points.
					$tControlPoint3 = Null
				EndIf

			Else ; Need to create new control points if IsCurve = True.
				If ($bIsCurve = True) Then
					$tControlPoint3 = __LOImpress_CreatePoint(Int($atPoints[$iArrayElement].X() + (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iArrayElement].Y() + (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

					$tControlPoint4 = __LOImpress_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
				EndIf
			EndIf

			$iOffset = 0
			$iForOffset = 0
			$iReDimCount = 0
			; Check how many control points I am keeping vs creating.
			$iReDimCount += (IsObj($tControlPoint3)) ? (1) : (0)
			$iReDimCount += (IsObj($tControlPoint4)) ? (1) : (0)
			$iReDimCount -= ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If a control point already existed, minus one from ReDim as it is either not new, or I am deleting it.
			$iReDimCount -= (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)

			ReDim $avArray[UBound($atPoints) + $iReDimCount]
			ReDim $avArray2[UBound($aiFlags) + $iReDimCount]

			$iReDimCount = 0

			For $i = 0 To UBound($atPoints) - 1
				If ($iOffset = 0) Then
					$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
					$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the points to the array.

				Else
					$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
					$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
				EndIf

				If ($i = $iArrayElement) Then ; Insert the new or modified control points.

					If IsObj($tControlPoint3) Then
						$iForOffset += 1
						If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.

						$avArray[$i + $iForOffset] = $tControlPoint3
						$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

					Else
						If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
					EndIf

					If IsObj($tControlPoint4) Then
						$iForOffset += 1
						If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1

						$avArray[$i + $iForOffset] = $tControlPoint4
						$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

					Else
						If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
					EndIf
				EndIf

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$atPoints = $avArray
			$aiFlags = $avArray2

		Else ; Point is a Curve, see if bIsCurve = False.
			If ($bIsCurve = False) Then ; If bIsCurve = True, I can just skip it, as there is nothing to do when the point is a curve already.

				If ($aiFlags[$iNextArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; Next point is a curve, need to keep the control points.
					If ($aiFlags[$iArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint3 = $atPoints[$iArrayElement + 1]
					If ($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint4 = $atPoints[$iArrayElement + 2]
				EndIf

				If ($iPreviousArrayElement <> -1) And ($aiFlags[$iPreviousArrayElement] <> $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; There is a previous point, and it is a curve, I need to keep the control points.
					If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]
					If ($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2]

				ElseIf ($iPreviousArrayElement <> -1) And ($aiFlags[$iPreviousArrayElement] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL) Then ; There is a previous point, and it is a normal point.
					; See if it is curved.

					If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) And _
							(($atPoints[$iPreviousArrayElement].X() <> $atPoints[$iPreviousArrayElement + 1].X()) And _
							($atPoints[$iPreviousArrayElement].Y() <> $atPoints[$iPreviousArrayElement + 1].Y())) Then ; Previous Point is a curve, need to keep the control points.
						$tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]

						If ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2]
					EndIf
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 4
				; Check how many control points I am keeping vs deleting.
				$iReDimCount -= (IsObj($tControlPoint1)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint2)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint3)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint4)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) - $iReDimCount]
				ReDim $avArray2[UBound($aiFlags) - $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the points to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iPreviousArrayElement) Then
						If IsObj($tControlPoint1) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint1
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						If IsObj($tControlPoint2) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint2
							$avArray2[$i + $iForOffset] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If (($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

					ElseIf ($i = $iArrayElement) Then ; Insert or skip Control Points as necessary.
						$avArray[$i] = $atPoints[$iArrayElement]
						$avArray2[$i] = $LOI_DRAWSHAPE_POINT_TYPE_NORMAL

						If IsObj($tControlPoint3) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + 1] = $tControlPoint3
							$avArray2[$i + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						If IsObj($tControlPoint4) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + 2] = $tControlPoint4
							$avArray2[$i + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL

						Else
							If (($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf
					EndIf

					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
				Next

				; Update the ArrayElement value to its new position.
				If ($iPreviousArrayElement <> -1) Then $iArrayElement -= ((IsObj($tControlPoint2) And ($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL))) ? (0) : (1) ; If ControlPoint 2 is a object, it means I copied it, meaining I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.
				If ($iPreviousArrayElement <> -1) Then $iArrayElement -= ((IsObj($tControlPoint1) And ($aiFlags[$iPreviousArrayElement + 1] = $LOI_DRAWSHAPE_POINT_TYPE_CONTROL))) ? (0) : (1) ; If ControlPoint 1 is a object, it means I copied it, meaning I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.

				$atPoints = $avArray
				$aiFlags = $avArray2
			EndIf
		EndIf
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_DrawShapePointModify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_FilterNameGet
; Description ...: Retrieves the correct L.O. Filter name for use in SaveAs and Export.
; Syntax ........: __LOImpress_FilterNameGet(ByRef $sDocSavePath[, $bExportFilters = False])
; Parameters ....: $sDocSavePath        - [in/out] a string value. Full path with extension.
;                  $bExportFilters      - [optional] a boolean value. Default is False. If True, includes the Filter Names that can be used to Export only, in the search.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sDocSavePath is not a string.
;                  @Error 1 @Extended 2 Return 0 = $bExportFilters not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sDocSavePath is not a correct path or URL.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Success. Returning required filter name from "SaveAs" Filter Names.
;                  @Error 0 @Extended 2 Return String = Success. Returning required filter name from "Export" Filter Names.
;                  @Error 0 @Extended 3 Return String = Filter Name not found for given file extension, defaulting to .odp file format and updating save path accordingly.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Searches a predefined list of extensions stored in an array. Not all FilterNames are listed.
;                  For finding your own Filter Names, see convertfilters.html found in L.O. Install Folder: LibreOffice\help\en-US\text\shared\guide
;                  Or See: "OOME_3_0", "OpenOffice.org Macros Explained OOME Third Edition" by Andrew D. Pitonyak, which has a handy Macro for listing all Filter Names, found on page 284 of the above book in the ODT format.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_FilterNameGet(ByRef $sDocSavePath, $bExportFilters = False)
	Local $iLength, $iSlashLocation, $iDotLocation
	Local Const $STR_NOCASESENSE = 0, $STR_STRIPALL = 8
	Local $sFileExtension, $sFilterName
	Local $msSaveAsFilters[], $msExportFilters[]

	If Not IsString($sDocSavePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExportFilters) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iLength = StringLen($sDocSavePath)

	$msSaveAsFilters[".fodp"] = "OpenDocument Presentation Flat XML"
	$msSaveAsFilters[".pot"] = "PowerPoint 3"
	$msSaveAsFilters[".potx"] = "Impress MS PowerPoint 2007 XML Template" ; Note these have a XML version too.
	$msSaveAsFilters[".pps"] = "MS PowerPoint 97 AutoPlay"
	$msSaveAsFilters[".ppt"] = "PowerPoint 3"
	$msSaveAsFilters[".ppsx"] = "Impress MS PowerPoint 2007 XML AutoPlay" ; Note these have a XML version too.
	$msSaveAsFilters[".pptm"] = "Impress MS PowerPoint 2007 XML VBA"
	$msSaveAsFilters[".pptx"] = "Impress MS PowerPoint 2007 XML" ; Note these have a XML version too.
	$msSaveAsFilters[".odg"] = "impress8_draw"
	$msSaveAsFilters[".odp"] = "impress8"
	$msSaveAsFilters[".otp"] = "impress8_template"
	$msSaveAsFilters[".uop"] = "UOF presentation"

	If $bExportFilters Then
		$msExportFilters[".apng"] = "impress_png_Export"
		$msExportFilters[".bmp"] = "impress_bmp_Export"
		$msExportFilters[".emf"] = "impress_emf_Export"
		$msExportFilters[".eps"] = "impress_eps_Export"
		$msExportFilters[".gif"] = "impress_gif_Export"
		$msExportFilters[".htm"] = "impress_html_Export"
		$msExportFilters[".html"] = "impress_html_Export"
		$msExportFilters[".jfif"] = "impress_jpg_Export"
		$msExportFilters[".jif"] = "impress_jpg_Export"
		$msExportFilters[".jpg"] = "impress_jpg_Export"
		$msExportFilters[".jpeg"] = "impress_jpg_Export"
		$msExportFilters[".svg"] = "impress_svg_Export"
		$msExportFilters[".pdf"] = "impress_pdf_Export"
		$msExportFilters[".png"] = "impress_png_Export"
		$msExportFilters[".tif"] = "impress_tif_Export"
		$msExportFilters[".tiff"] = "impress_tif_Export"
		$msExportFilters[".webp"] = "writer_webp_Export"
		$msExportFilters[".wmf"] = "impress_wmf_Export"
		$msExportFilters[".xhtml"] = "XHTML Impress File"
	EndIf

	If StringInStr($sDocSavePath, "file:///") Then ;  If L.O. URl Then
		$iSlashLocation = StringInStr($sDocSavePath, "/", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, ($iLength - $iDotLocation) + 1)

	ElseIf StringInStr($sDocSavePath, "\") Then ;  Else if PC Path Then
		$iSlashLocation = StringInStr($sDocSavePath, "\", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, $iLength - $iDotLocation + 1)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	If $sFileExtension = $sDocSavePath Then ;  If no file extension identified, append .ods extension and return.
		$sDocSavePath = $sDocSavePath & ".odp"

		Return SetError($__LO_STATUS_SUCCESS, 3, "impress8")

	Else
		$sFileExtension = StringLower(StringStripWS($sFileExtension, $STR_STRIPALL))
	EndIf

	$sFilterName = $msSaveAsFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilterName)

	If $bExportFilters Then $sFilterName = $msExportFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilterName)

	$sDocSavePath = StringReplace($sDocSavePath, $sFileExtension, ".odp") ; If No results, replace with ODS extension.

	Return SetError($__LO_STATUS_SUCCESS, 3, "impress8")
EndFunc   ;==>__LOImpress_FilterNameGet

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_GetShapeName
; Description ...: Create a Shape Name that hasn't been used yet in the slide.
; Syntax ........: __LOImpress_GetShapeName(ByRef $oSlide, $sShapeName)
; Parameters ....: $oSlide              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
;                  $sShapeName          - a string value. The Shape name to begin with.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSlide not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sShapeName not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Slide contained no shapes, returning the Shape name with a "1" appended.
;                  @Error 0 @Extended 1 Return String = Success. Returning the unique Shape name to use.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function adds a digit after the shape name, incrementing it until that name hasn't been used yet in L.O.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_GetShapeName(ByRef $oSlide, $sShapeName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oSlide.hasElements() Then
		Do ; Cycle through until I find a unique name.
			$iCount += 1
			For $i = 0 To $oSlide.getCount() - 1
				; Impress doesn't set the Shape name on new shapes. It has names in the UI that would correspond to the order of the shapes inserted, i.e. Shape 1, Shape 2. Etc.
				If ($oSlide.getByIndex($i).Name() = $sShapeName & $iCount) Or (($oSlide.getByIndex($i).Name() = "") And (("Shape " & ($i + 1)) = $sShapeName & $iCount)) Then ExitLoop

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
			Next
		Until $i = $oSlide.getCount()

	Else

		Return SetError($__LO_STATUS_SUCCESS, 0, $sShapeName & "1") ; If Doc has no shapes, just return the name with a "1" appended.
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 1, $sShapeName & $iCount)
EndFunc   ;==>__LOImpress_GetShapeName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_GradientNameInsert
; Description ...: Create and insert a new Gradient name.
; Syntax ........: __LOImpress_GradientNameInsert(ByRef $oDoc, $tGradient[, $sGradientName = "Gradient "])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $tGradient           - a dll struct value. A Gradient Structure to copy settings from.
;                  $sGradientName       - [optional] a string value. Default is "Gradient ". The Gradient name to create.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tGradient not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sGradientName not a string.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.GradientTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" or "com.sun.star.awt.Gradient2" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error creating Gradient Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. A new Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Gradient name is blank, I need to create a new name and apply it. I think I could re-use an old one without problems, but I'm not sure, so to be safe, I will create a new one.
;                  If there are no names that have been already created, then I need to create and apply one before the gradient will be displayed.
;                  Else if a preset Gradient is called, I need to create its name before it can be used.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_GradientNameInsert(ByRef $oDoc, $tGradient, $sGradientName = "Gradient ")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tNewGradient
	Local $oGradTable
	Local $iCount = 1
	Local $sGradient = "com.sun.star.awt.Gradient2"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If Not __LO_VersionCheck(7.6) Then $sGradient = "com.sun.star.awt.Gradient"

	$oGradTable = $oDoc.createInstance("com.sun.star.drawing.GradientTable")
	If Not IsObj($oGradTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sGradientName = "Gradient ") Then
		While $oGradTable.hasByName($sGradientName & $iCount)
			$iCount += 1
			Sleep((IsInt($iCount / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		WEnd
		$sGradientName = $sGradientName & $iCount
	EndIf

	$tNewGradient = __LO_CreateStruct($sGradient)
	If Not IsObj($tNewGradient) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; Copy the settings over from the input Style Gradient to my new one. This may not be necessary? But just in case.
	With $tNewGradient
		.Style = $tGradient.Style()
		.XOffset = $tGradient.XOffset()
		.YOffset = $tGradient.YOffset()
		.Angle = $tGradient.Angle()
		.Border = $tGradient.Border()
		.StartColor = $tGradient.StartColor()
		.EndColor = $tGradient.EndColor()
		.StartIntensity = $tGradient.StartIntensity()
		.EndIntensity = $tGradient.EndIntensity()

		If __LO_VersionCheck(7.6) Then .ColorStops = $tGradient.ColorStops()

	EndWith

	If Not $oGradTable.hasByName($sGradientName) Then
		$oGradTable.insertByName($sGradientName, $tNewGradient)
		If Not ($oGradTable.hasByName($sGradientName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $sGradientName)
EndFunc   ;==>__LOImpress_GradientNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_GradientPresets
; Description ...: Set Page background Gradient to preset settings.
; Syntax ........: __LOImpress_GradientPresets(ByRef $oDoc, ByRef $oObject, ByRef $tGradient, $sGradientName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oObject             - [in/out] an object. The Object to modify the Gradient settings for.
;                  $tGradient           - [in/out] an object. The Fill Gradient Object to modify the Gradient settings for.
;                  $sGradientName       - a string value. The Gradient Preset name to apply.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.awt.ColorStop" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to create Gradient name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. The Style Gradient settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_GradientPresets(ByRef $oDoc, ByRef $oObject, ByRef $tGradient, $sGradientName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tColorStop, $tStopColor
	Local $atColorStop[2]

	If __LO_VersionCheck(7.6) Then
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$atColorStop[0] = $tColorStop

		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$atColorStop[1] = $tColorStop
	EndIf

	Switch $sGradientName
		Case $LOI_GRAD_NAME_PASTEL_BOUQUET
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 14543051
				.EndColor = 16766935
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_PASTEL_DREAM
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RECT
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 450
				.Border = 0
				.StartColor = 16766935
				.EndColor = 11847644
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_BLUE_TOUCH
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 10
				.Border = 0
				.StartColor = 11847644
				.EndColor = 14608111
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_BLANK_W_GRAY
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 900
				.Border = 75
				.StartColor = $LO_COLOR_WHITE
				.EndColor = 14540253
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_LONDON_MIST
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 13421772
				.EndColor = 6710886
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUBMARINE
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 14543051
				.EndColor = 11847644
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_MIDNIGHT
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = $LO_COLOR_BLACK
				.EndColor = 2777241
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_DEEP_OCEAN
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 0
				.StartColor = $LO_COLOR_BLACK
				.EndColor = 7512015
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_MAHOGANY
			With $tGradient
				.Style = $LOI_GRAD_TYPE_SQUARE
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 450
				.Border = 0
				.StartColor = $LO_COLOR_BLACK
				.EndColor = 9250846
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_GREEN_GRASS
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 16776960
				.EndColor = 8508442
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_NEON_LIGHT
			With $tGradient
				.Style = $LOI_GRAD_TYPE_ELLIPTICAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 15
				.StartColor = 1209890
				.EndColor = $LO_COLOR_WHITE
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNSHINE
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 66
				.YOffset = 33
				.Angle = 0
				.Border = 33
				.StartColor = 16760576
				.EndColor = 16776960
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_RAINBOW
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 100
				.Angle = 0
				.Border = 0
				.StartColor = $LO_COLOR_WHITE
				.EndColor = $LO_COLOR_WHITE
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					ReDim $atColorStop[7]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0.2
					$atColorStop[0] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.2

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 0
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.4

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 1
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0
					$tStopColor.Green = 1
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[3] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.65

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0
					$tStopColor.Green = 1
					$tStopColor.Blue = 1
					$tColorStop.StopColor = $tStopColor

					$atColorStop[4] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.8

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 0
					$tStopColor.Blue = 1
					$tColorStop.StopColor = $tStopColor

					$atColorStop[5] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.8
					$atColorStop[6] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNRISE
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 3713206
				.EndColor = 14065797
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					ReDim $atColorStop[4]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.505882352941176
					$tStopColor.Green = 0.784313725490196
					$tStopColor.Blue = 0.768627450980392
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.75

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.717647058823529
					$tStopColor.Green = 0.807843137254902
					$tStopColor.Blue = 0.698039215686275
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 1
					$atColorStop[3] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNDOWN
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 985943
				.EndColor = 16759664
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					ReDim $atColorStop[5]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.3

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.392156862745098
					$tStopColor.Green = 0.305882352941177
					$tStopColor.Blue = 0.690196078431373
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.827450980392157
					$tStopColor.Green = 0.572549019607843
					$tStopColor.Blue = 0.83921568627451
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.75

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.996078431372549
					$tStopColor.Green = 0.733333333333333
					$tStopColor.Blue = 0.76078431372549
					$tColorStop.StopColor = $tStopColor

					$atColorStop[3] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 1
					$atColorStop[4] = $tColorStop
				EndIf
			EndWith

		Case Else ; Custom Gradient Name
			__LOImpress_GradientNameInsert($oDoc, $tGradient, $sGradientName)
			If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oObject.FillGradientName = $sGradientName

			Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndSwitch

	If __LO_VersionCheck(7.6) Then
		$tColorStop = $atColorStop[0] ; "Start Color" Value.

		$tStopColor = $tColorStop.StopColor()

		$tStopColor.Red = (BitAND(BitShift($tGradient.StartColor(), 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($tGradient.StartColor(), 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($tGradient.StartColor(), 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStop[0] = $tColorStop

		$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; Last element is "End Color" Value.

		$tStopColor = $tColorStop.StopColor()

		$tStopColor.Red = (BitAND(BitShift($tGradient.EndColor(), 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($tGradient.EndColor(), 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($tGradient.EndColor(), 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStop[UBound($atColorStop) - 1] = $tColorStop

		$tGradient.ColorStops = $atColorStop
	EndIf

	__LOImpress_GradientNameInsert($oDoc, $tGradient, $sGradientName)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oObject.FillGradient = $tGradient
	$oObject.FillGradientName = $sGradientName
	$oObject.FillGradientStepCount = $tGradient.StepCount()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_GradientPresets

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LOImpress_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LOImpress_ComError_UserFunction(Default)
	Local $vUserFunction, $avUserParams[2] = ["CallArgArray", $oComError]

	If IsArray($avUserFunction) Then
		$vUserFunction = $avUserFunction[0]
		ReDim $avUserParams[UBound($avUserFunction) + 1]
		For $i = 1 To UBound($avUserFunction) - 1
			$avUserParams[$i + 1] = $avUserFunction[$i]
		Next

	Else
		$vUserFunction = $avUserFunction
	EndIf
	If IsFunc($vUserFunction) Then
		Switch $vUserFunction
			Case ConsoleWrite
				ConsoleWrite("!--COM Error-Begin--" & @CRLF & _
						"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline & @CRLF & _
						"!--COM-Error-End--" & @CRLF)

			Case MsgBox
				MsgBox(0, "COM Error", "Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline)

			Case Else
				Call($vUserFunction, $avUserParams)
		EndSwitch
	EndIf
EndFunc   ;==>__LOImpress_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ShapeGetType
; Description ...: Identify a Shape's type.
; Syntax ........: __LOImpress_ShapeGetType(ByRef $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_DrawShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape's type.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify Shape.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning the Shape's Type, corresponding to one of the Constants $LOI_SHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ShapeGetType(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avShapeTypes[20][2] = [[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.Shape3DSceneObject"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.CustomShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.MeasureShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.EllipseShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.ClosedBezierShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.OpenBezierShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.PolyPolygonShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.PolyLineShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.LineShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.ConnectorShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.OpenFreeHandShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.ClosedFreeHandShape"], [$LOI_SHAPE_TYPE_FORM_CONTROL, "com.sun.star.drawing.ControlShape"], _
			[$LOI_SHAPE_TYPE_IMAGE, "com.sun.star.drawing.GraphicObjectShape"], [$LOI_SHAPE_TYPE_MEDIA, "com.sun.star.drawing.MediaShape"], _
			[$LOI_SHAPE_TYPE_OLE2, "com.sun.star.drawing.OLE2Shape"], [$LOI_SHAPE_TYPE_TABLE, "com.sun.star.drawing.TableShape"], _
			[$LOI_SHAPE_TYPE_TEXTBOX, "com.sun.star.drawing.TextShape"], [$LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE, "com.sun.star.presentation.SubtitleShape"], _
			[$LOI_SHAPE_TYPE_TEXTBOX_TITLE, "com.sun.star.presentation.TitleTextShape"]]
	Local $sShapeType

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sShapeType = $oShape.ShapeType()
	If Not IsString($sShapeType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($avShapeTypes) - 1
		If ($sShapeType = $avShapeTypes[$i][1]) Then Return SetError($__LO_STATUS_SUCCESS, 0, $avShapeTypes[$i][0])
	Next

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
EndFunc   ;==>__LOImpress_ShapeGetType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_TransparencyGradientConvert
; Description ...: Convert a Transparency Gradient percentage value to a color value or from a color value to a percentage.
; Syntax ........: __LOImpress_TransparencyGradientConvert([$iPercentToLong = Null[, $iLongToPercent = Null]])
; Parameters ....: $iPercentToLong      - [optional] an integer value. Default is Null. The percentage to convert to a RGB Color Integer.
;                  $iLongToPercent      - [optional] an integer value. Default is Null. The RGB Color Integer to convert to percentage.
; Return values .: Success: Integer.
;                  Failure: Null and sets the @Error and @Extended flags to non-zero.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return Null = No values called in parameters.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. The requested Integer value converted from percentage to a RGB Color Integer.
;                  @Error 0 @Extended 1 Return Integer = Success. The requested Integer value from a RGB Color Integer to percentage.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_TransparencyGradientConvert($iPercentToLong = Null, $iLongToPercent = Null)
	Local $iReturn

	If ($iPercentToLong <> Null) Then
		$iReturn = ((255 * ($iPercentToLong / 100)) + .50) ; Change percentage to decimal and times by White color (255 RGB) Add . 50 to round up if applicable.
		$iReturn = _LO_ConvertColorToLong(Int($iReturn), Int($iReturn), Int($iReturn))

		Return SetError($__LO_STATUS_SUCCESS, 0, $iReturn)

	ElseIf ($iLongToPercent <> Null) Then
		$iReturn = _LO_ConvertColorFromLong(Null, $iLongToPercent)
		$iReturn = Int((($iReturn[0] / 255) * 100) + .50) ; All return color values will be the same, so use only one. Add . 50 to round up if applicable.

		Return SetError($__LO_STATUS_SUCCESS, 1, $iReturn)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, Null)
	EndIf
EndFunc   ;==>__LOImpress_TransparencyGradientConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_TransparencyGradientNameInsert
; Description ...: Create and insert a new Transparency Gradient name.
; Syntax ........: __LOImpress_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $tTGradient          - a dll struct value. A Gradient Structure to copy settings from.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tTGradient not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.TransparencyGradientTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error creating Transparency Gradient Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. A new transparency Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Transparency Gradient name is blank, I need to create a new name and apply it. I think I could re-use an old one without problems, but I'm not sure, so to be safe, I will create a new one.
;                  If there are no names that have been already created, then I need to create and apply one before the transparency gradient will be displayed.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tNewTGradient
	Local $oTGradTable
	Local $iCount = 1
	Local $sGradient = "com.sun.star.awt.Gradient2"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tTGradient) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not __LO_VersionCheck(7.6) Then $sGradient = "com.sun.star.awt.Gradient"

	$oTGradTable = $oDoc.createInstance("com.sun.star.drawing.TransparencyGradientTable")
	If Not IsObj($oTGradTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oTGradTable.hasByName("Transparency " & $iCount)
		$iCount += 1
		Sleep((IsInt($iCount / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
	WEnd

	$tNewTGradient = __LO_CreateStruct($sGradient)
	If Not IsObj($tNewTGradient) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; Copy the settings over from the input Style Gradient to my new one. This may not be necessary? But just in case.
	With $tNewTGradient
		.Style = $tTGradient.Style()
		.XOffset = $tTGradient.XOffset()
		.YOffset = $tTGradient.YOffset()
		.Angle = $tTGradient.Angle()
		.Border = $tTGradient.Border()
		.StartColor = $tTGradient.StartColor()
		.EndColor = $tTGradient.EndColor()

		If __LO_VersionCheck(7.6) Then .ColorStops = $tTGradient.ColorStops()
	EndWith

	$oTGradTable.insertByName("Transparency " & $iCount, $tNewTGradient)
	If Not ($oTGradTable.hasByName("Transparency " & $iCount)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, "Transparency " & $iCount)
EndFunc   ;==>__LOImpress_TransparencyGradientNameInsert
