#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $iMicrometers, $iMicrometers2
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

	; Convert 1" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(1)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2" to Micrometers
	$iMicrometers2 = _LO_ConvertToMicrometer(2)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape's Position to 1" over, 2" down, and set Protect Position to True.
	_LOImpress_DrawShapePosition($oShape, $iMicrometers, $iMicrometers2, True)
	If @error Then _ERROR($oDoc, "Failed to modify Shape's position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapePosition($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Position settings are as follows: " & @CRLF & _
			"The Shape's X coordinate is, in Micrometers: " & $avSettings[0] & @CRLF & _
			"The Shape's Y coordinate is, in Micrometers: " & $avSettings[1] & @CRLF & _
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
