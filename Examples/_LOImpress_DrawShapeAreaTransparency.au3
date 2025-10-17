#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $iTransparency

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

	; Modify the Shape Transparency settings to 55% transparent
	_LOImpress_DrawShapeAreaTransparency($oShape, 55)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape Transparency. Return will be an Integer.
	$iTransparency = _LOImpress_DrawShapeAreaTransparency($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's current Transparency percentage is: " & $iTransparency)

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
