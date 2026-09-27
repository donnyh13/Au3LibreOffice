#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oStyle
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_BLANK
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_BLANK)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the "Default Drawing Style" Shape Style.
	$oStyle = _LOImpress_ShapeStyleGetObjByname($oDoc, "standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; 150 Distance between the baseline, 65 Guide overhang, 76 Guide distance, 55 Left Guide length, 76 Right guide length, Measure below, 1 decimal places,
	; Position the measurement line below the text and centered, Display the measurement parallel to the line, and use the Pica unit.
	_LOImpress_ShapeStyleDimensionSettings($oStyle, 150, 65, 76, 55, 76, True, 1, $LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_BOTTOM, $LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_CENTER, True, $LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_PICA)
	If @error Then _ERROR($oDoc, "Failed to modify Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_ShapeStyleDimensionSettings($oStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Style's Dimension settings are as follows: " & @CRLF & _
			"The distance of the Dimension line from the base line is (in Hundredths of a Millimeter (HMM)): " & $avSettings[0] & @CRLF & _
			"The Guide overhang is (in Hundredths of a Millimeter (HMM)): " & $avSettings[1] & @CRLF & _
			"The Guide distance is (in Hundredths of a Millimeter (HMM)): " & $avSettings[2] & @CRLF & _
			"The Left Guide length is (in Hundredths of a Millimeter (HMM)): " & $avSettings[3] & @CRLF & _
			"The Right Guide length is (in Hundredths of a Millimeter (HMM)): " & $avSettings[4] & @CRLF & _
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

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
