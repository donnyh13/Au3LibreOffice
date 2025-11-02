#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape, $oConnectorShape
	Local $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Slide, 3000 Wide by 6000 High, 14500X and 3100Y.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000, 14500, 3100)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Connector Shape into the Slide, 4000 Wide by 1000 High, 500X and 1400Y.
	$oConnectorShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ENDS_ARROW, 4000, 1000, 500, 1400)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the beginning of the connector to 5500X, 2500Y, and connect the end of the connector to the rectangle, at the second GluePoint.
	_LOImpress_DrawShapeConnectorModify($oConnectorShape, 5500, 2500, Null, Null, Null, Null, $oShape, 2)
	If @error Then _ERROR($oDoc, "Failed to modify Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeConnectorModify($oConnectorShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Connector Shape's settings are as follows: " & @CRLF & _
			"The Connector Start X position is: " & $avSettings[0] & @CRLF & _
			"The Connector Start Y position is: " & $avSettings[1] & @CRLF & _
			"The Shape the Start of the Connector is connected to is here (if any), I will check if it is an Object only: " & IsObj($avSettings[2]) & @CRLF & _
			"The GluePoint the Start of the Connector is connected to (if any) is: " & $avSettings[3] & @CRLF & _
			"The Connector End X position is: " & $avSettings[4] & @CRLF & _
			"The Connector End Y position is: " & $avSettings[5] & @CRLF & _
			"The Shape the End of the Connector is connected to is here (if any), I will check if it is an Object only: " & IsObj($avSettings[6]) & @CRLF & _
			"The GluePoint the Start of the Connector is connected to (if any) is: " & $avSettings[7])

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
