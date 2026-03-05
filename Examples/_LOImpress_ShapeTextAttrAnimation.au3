#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide
	Local $avSettings, $avShapes

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_TITLE_CONTENT
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_TITLE_CONTENT)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of Textboxes in the slide.
	$avShapes = _LOImpress_SlideShapesGetList($oSlide, BitOR($LOI_SHAPE_TYPE_TEXTBOX, $LOI_SHAPE_TYPE_TEXTBOX_TITLE, $LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE, $LOI_SHAPE_TYPE_TEXTBOX_OUTLINER))
	If @error Or (@extended = 0) Then _ERROR($oDoc, "Failed to retrieve Shapes, or no Shapes present in Slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set animation type to Scroll alternate, Begin on the left and go towards the right, Don't begin with the text visible, leave the text visible when the animation finishes,
	; Repeat 5 times, increment the text by 23 units, interpret the increment in pixels, and delay each animation cycle by 250 ms.
	_LOImpress_ShapeTextAttrAnimation($avShapes[1][0], $LOI_ANIMATION_TYPE_SCROLL_ALTERNATE, $LOI_ANIMATION_DIR_RIGHT, False, True, 5, 23, True, 250)
	If @error Then _ERROR($oDoc, "Failed to modify Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_ShapeTextAttrAnimation($avShapes[1][0])
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Text Box's settings are as follows: " & @CRLF & _
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
