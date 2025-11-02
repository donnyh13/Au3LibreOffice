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

	; Insert a Connector Shape into the Slide, 8000 Wide by 1000 High, 5000X and 5400Y.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ENDS_ARROW, 8000, 1000, 5000, 5400)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Connector line style to be a Line type, with 500 Line 1 Skew, 50 Line 2 Skew, 0 Line 3 Skew, 300 horizontal spacing at the beginning,
	; 550  horizontal spacing at the end, 700 vertical spacing at the beginning, and 1500 vertical spacing at the end.
	_LOImpress_DrawShapeConnectorSettings($oShape, $LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE, 500, 50, 0, 300, 550, 700, 1500)
	If @error Then _ERROR($oDoc, "Failed to modify Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeConnectorSettings($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Connector Shape's settings are as follows: " & @CRLF & _
			"The Connector line style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The skew value of Line 1 is: " & $avSettings[1] & @CRLF & _
			"The skew value of Line 2 is: " & $avSettings[2] & @CRLF & _
			"The skew value of Line 3 is: " & $avSettings[3] & @CRLF & _
			"The horizontal spacing at the beginning of the connector is (in 1/100th mm): " & $avSettings[4] & @CRLF & _
			"The horizontal spacing at the end of the connector is (in 1/100th mm): " & $avSettings[5] & @CRLF & _
			"The vertical spacing at the beginning of the connector is (in 1/100th mm): " & $avSettings[6] & @CRLF & _
			"The vertical spacing at the end of the connector is (in 1/100th mm): " & $avSettings[7])

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
