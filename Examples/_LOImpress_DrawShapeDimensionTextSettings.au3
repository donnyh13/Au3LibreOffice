#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Dimension Shape into the Slide, 12000 Wide by 100 High, 5000X and 5400Y.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION, 12000, 100, 5000, 5400)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Don't fit the text to the Text Box's size, don't contour the text to the shape, 175 distance left and right, 148 distance at the top, 130 distance at the bottom,
	; Position the text at the top-left, and do not fit the text to the full width.
	_LOImpress_DrawShapeDimensionTextSettings($oShape, False, False, 175, 175, 148, 130, $LOI_TEXT_ANCHOR_TOP_LEFT, False)
	If @error Then _ERROR($oDoc, "Failed to modify Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeDimensionTextSettings($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Connector Shape's settings are as follows: " & @CRLF & _
			"Is the Text adjusted to fit the Text Box size? True/False: " & $avSettings[0] & @CRLF & _
			"Is the text flow contoured to fit the Shape's contour? True/False: " & $avSettings[1] & @CRLF & _
			"The amount of space between the left edge of the text and the left edge of the shape is (in Hundredths of a Millimeter (HMM)): " & $avSettings[2] & @CRLF & _
			"The amount of space between the right edge of the text and the right edge of the shape is (in Hundredths of a Millimeter (HMM)): " & $avSettings[3] & @CRLF & _
			"The amount of space between the top edge of the text and the top edge of the shape is (in Hundredths of a Millimeter (HMM)): " & $avSettings[4] & @CRLF & _
			"The amount of space between the bottom edge of the text and the bottom edge of the shape is (in Hundredths of a Millimeter (HMM)): " & $avSettings[5] & @CRLF & _
			"The position the text is anchored is (See UDF Constants): " & $avSettings[6] & @CRLF & _
			"Will the text be fit to the shape's full width? True/False" & $avSettings[7])

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
