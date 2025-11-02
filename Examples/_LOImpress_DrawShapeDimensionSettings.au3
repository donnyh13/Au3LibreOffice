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

	; 150 Distance between the baseline, 65 Guide overhang, 76 Guide distance, 55 Left Guide length, 76 Right guide length, Measure below, 1 decimal places,
	; Position the measurement line below the text and centered, Display the measurement parallel to the line, and use the Pica unit.
	_LOImpress_DrawShapeDimensionSettings($oShape, 150, 65, 76, 55, 76, True, 1, $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_BOTTOM, $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_CENTER, True, $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_PICA)
	If @error Then _ERROR($oDoc, "Failed to modify Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeDimensionSettings($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Dimension Shape's settings are as follows: " & @CRLF & _
			"The distance of the Dimension line from the base line is (in 1/100th mm): " & $avSettings[0] & @CRLF & _
			"The Guide overhang is (in 1/100th mm): " & $avSettings[1] & @CRLF & _
			"The Guide distance is (in 1/100th mm): " & $avSettings[2] & @CRLF & _
			"The Left Guide length is (in 1/100th mm): " & $avSettings[3] & @CRLF & _
			"The Right Guide length is (in 1/100th mm): " & $avSettings[4] & @CRLF & _
			"Measure below the shape? True/False: " & $avSettings[5] & @CRLF & _
			"The number of decimal places is: " & $avSettings[6] & @CRLF & _
			"The Vertical alignment of the Dimension line to the text is (See UDF Constants): " & $avSettings[7] & @CRLF & _
			"The Horizontal alignment of the Dimension line text is (See UDF Constants): " & $avSettings[8] & @CRLF & _
			"Display the measurement value parallel to the line? True/False: " & $avSettings[9] & @CRLF & _
			"The Unit type to display the measurement in is (See UDF Constants): " & $avSettings[10])

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
