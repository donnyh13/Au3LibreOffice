#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $iHMM, $iHMM2
	Local $avSettings, $avShapes

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_TITLE_ONLY
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_TITLE_ONLY)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of Textboxes in the slide.
	$avShapes = _LOImpress_SlideShapesGetList($oSlide, BitOR($LOI_SHAPE_TYPE_TEXTBOX, $LOI_SHAPE_TYPE_TEXTBOX_TITLE, $LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE))
	If @error Or (@extended = 0) Then _ERROR($oDoc, "Failed to retrieve Shapes, or no Shapes present in Slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert .5" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 4.5" to Hundredths of a Millimeter (HMM)
	$iHMM2 = _LO_UnitConvert(4.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Title box's Position to .5" over, 4.5" down.
	_LOImpress_ShapePosition($avShapes[0][0], $iHMM, $iHMM2)
	If @error Then _ERROR($oDoc, "Failed to modify Shape's position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000, 2000, 3500)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(1, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2" to Hundredths of a Millimeter (HMM)
	$iHMM2 = _LO_UnitConvert(2, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape's Position to 1" over, 2" down, and set Protect Position to True.
	_LOImpress_ShapePosition($oShape, $iHMM, $iHMM2, True)
	If @error Then _ERROR($oDoc, "Failed to modify Shape's position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_ShapePosition($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Position settings are as follows: " & @CRLF & _
			"The Shape's X coordinate is, in Hundredths of a Millimeter (HMM): " & $avSettings[0] & @CRLF & _
			"The Shape's Y coordinate is, in Hundredths of a Millimeter (HMM): " & $avSettings[1] & @CRLF & _
			"Is the Shape's position protected against changes? True/False: " & $avSettings[2])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
