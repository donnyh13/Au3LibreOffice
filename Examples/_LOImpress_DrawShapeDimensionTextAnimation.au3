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

	; Set animation type to Scroll alternate, Begin on the left and go towards the right, Don't begin with the text visible, leave the text visible when the animation finishes,
	; Repeat 5 times, increment the text by 23 units, interpret the increment in pixels, and delay each animation cycle by 250 ms.
	_LOImpress_DrawShapeDimensionTextAnimation($oShape, $LOI_ANIMATION_TYPE_SCROLL_ALTERNATE, $LOI_ANIMATION_DIR_RIGHT, False, True, 5, 23, True, 250)
	If @error Then _ERROR($oDoc, "Failed to modify Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeDimensionTextAnimation($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Dimension Shape's settings are as follows: " & @CRLF & _
			"The Animation type is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Animation direction is (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"Is the Text visible and inside shape when the effect is applied? True/False: " & $avSettings[2] & @CRLF & _
			"Is the Text visible and inside shape when the effect is finished? True/False: " & $avSettings[3] & @CRLF & _
			"How many time will the animation repeat?: " & $avSettings[4] & @CRLF & _
			"How much is the text incremented by?: " & $avSettings[5] & @CRLF & _
			"Is the increment measured in Pixels? True/False: " & $avSettings[6] & @CRLF & _
			"How much delay is between each animation repeat? (In Milliseconds): " & $avSettings[7])

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
