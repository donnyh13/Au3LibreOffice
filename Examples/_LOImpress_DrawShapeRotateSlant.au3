#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Shape.
	_LOImpress_DrawShapePosition($oShape, 12000, 4300)
	If @error Then _ERROR($oDoc, "Failed to move a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Rotate the Shape 45 degrees.
	_LOImpress_DrawShapeRotateSlant($oShape, 45)
	If @error Then _ERROR($oDoc, "Failed to Rotate the Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeRotateSlant($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Rotation and Slant settings are as follows: " & @CRLF & _
			"The Shape is rotated " & $avSettings[0] & " degrees" & @CRLF & _
			"The Shape has a " & $avSettings[1] & " degree slant applied to it")

	; Rotate the Shape back to 0 degrees and apply a 23 degree slant to it.
	_LOImpress_DrawShapeRotateSlant($oShape, 0, 23)
	If @error Then _ERROR($oDoc, "Failed to Rotate the Shape and apply a slant. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeRotateSlant($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Rotation and Slant settings are as follows: " & @CRLF & _
			"The Shape is rotated " & $avSettings[0] & " degrees" & @CRLF & _
			"The Shape has a " & $avSettings[1] & " degree slant applied to it")

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
