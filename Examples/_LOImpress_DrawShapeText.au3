#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $sText, $sCurrentText

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Dimension Shape into the Slide, 12000 Wide by 100 High, 5000X and 5400Y.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION, 12000, 100, 5000, 5400)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Shape's current string.
	$sCurrentText = _LOImpress_DrawShapeText($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape's current string. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's current text is: " & @CRLF & _
			$sCurrentText & @CRLF & @CRLF & _
			"Press ok to enter a custom string.")

	; Prompt the user for some text.
	$sText = InputBox("Text Entry", "Please enter some text to set the shape to.", "Some AutoIt Text", " M")

	; Set the Shape's text to the input string.
	_LOImpress_DrawShapeText($oShape, $sText)
	If @error Then _ERROR($oDoc, "Failed to modify Shape text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Shape's current string.
	$sCurrentText = _LOImpress_DrawShapeText($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape's current string. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's current text is: " & @CRLF & _
			$sCurrentText)

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
