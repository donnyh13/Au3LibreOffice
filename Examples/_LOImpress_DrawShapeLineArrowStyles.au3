#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $iMicrometers
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Line Shape into the document, 4000 Wide by 0 High, 12000X, 4300Y.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_LINE_LINE, 4000, 0, 12000, 4300)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape Arrow Style settings to: Set the Start Arrowhead to $LOI_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, Start width = 1/4",
	; Start Center = True, Synchronize Start and End = True.
	_LOImpress_DrawShapeLineArrowStyles($oShape, $LOI_DRAWSHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, $iMicrometers, True, True)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeLineArrowStyles($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Arrow Style settings are as follows: " & @CRLF & _
			"The Start Arrowhead Style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Start Arrowhead Width is, in Micrometers: " & $avSettings[1] & @CRLF & _
			"Is the Start Arrowhead centered on the line end? True/False: " & $avSettings[2] & @CRLF & _
			"Is the Starting and Ending Arrowhead settings synchronized? True/False: " & $avSettings[3] & @CRLF & _
			"The End Arrowhead Style is (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"The End Arrowhead Width is, in Micrometers: " & $avSettings[5] & @CRLF & _
			"Is the Start Arrowhead centered on the line end? True/False: " & $avSettings[6])

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
