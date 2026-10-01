#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oStyle
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the "Default Drawing Style" Shape Style.
	$oStyle = _LODraw_ShapeStyleGetObjByname($oDoc, "standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Connector line style to be a Line type, 300 horizontal spacing at the beginning,
	; 550 horizontal spacing at the end, 700 vertical spacing at the beginning, and 1500 vertical spacing at the end.
	_LODraw_ShapeStyleConnectorSettings($oStyle, $LOD_DRAWSHAPE_CONNECTOR_TYPE_LINE, 300, 550, 700, 1500)
	If @error Then _ERROR($oDoc, "Failed to modify Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_ShapeStyleConnectorSettings($oStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Style's Connector settings are as follows: " & @CRLF & _
			"The Connector line style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The horizontal spacing at the beginning of the connector is (in Hundredths of a Millimeter (HMM)): " & $avSettings[1] & @CRLF & _
			"The horizontal spacing at the end of the connector is (in Hundredths of a Millimeter (HMM)): " & $avSettings[2] & @CRLF & _
			"The vertical spacing at the beginning of the connector is (in Hundredths of a Millimeter (HMM)): " & $avSettings[3] & @CRLF & _
			"The vertical spacing at the end of the connector is (in Hundredths of a Millimeter (HMM)): " & $avSettings[4])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
