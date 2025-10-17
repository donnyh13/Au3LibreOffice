#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Slide, 3000 Wide by 6000 High.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape Background Color settings. Background color = $LO_COLOR_TEAL.
	_LOImpress_DrawShapeAreaColor($oShape, $LO_COLOR_TEAL)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape Transparency Gradient settings to: Gradient Type = $LOI_GRAD_TYPE_ELLIPTICAL, XCenter to 75%, YCenter to 45%, Angle to 180 degrees
	; Border to 16%, Start transparency to 10%, End Transparency to 62%
	_LOImpress_DrawShapeAreaTransparencyGradient($oShape, $LOI_GRAD_TYPE_ELLIPTICAL, 75, 45, 180, 16, 10, 62)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeAreaTransparencyGradient($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Transparency Gradient settings are as follows: " & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avSettings[1] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avSettings[2] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avSettings[3] & @CRLF & _
			"The percentage of area not covered by the transparency is: " & $avSettings[4] & @CRLF & _
			"The starting transparency percentage is: " & $avSettings[5] & @CRLF & _
			"The ending transparency percentage is: " & $avSettings[6])

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
