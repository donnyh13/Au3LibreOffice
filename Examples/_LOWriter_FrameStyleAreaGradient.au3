#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oFrameStyle, $oViewCursor, $oFrame
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a new Frame Style named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR("Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Frame into the document for demonstration.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	; Set the Frame's style to my created style, "Test Style"
	_LOWriter_FrameStyleSet($oDoc, $oFrame, "Test Style")
	If @error Then _ERROR("Failed to set Frame style. Error:" & @error & " Extended:" & @extended)

	; Modify the Frame Style Gradient settings to: Preset Gradient name = $LOW_GRAD_NAME_TEAL_TO_BLUE
	_LOWriter_FrameStyleAreaGradient($oDoc, $oFrameStyle, $LOW_GRAD_NAME_TEAL_TO_BLUE)
	If @error Then _ERROR("Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleAreaGradient($oDoc, $oFrameStyle)
	If @error Then _ERROR("Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame style's Gradient settings are as follows: " & @CRLF & _
			"The Gradient name is: " & $avSettings[0] & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The number of steps to increment color is: " & $avSettings[2] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avSettings[3] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avSettings[4] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avSettings[5] & @CRLF & _
			"The percentage of area not covered by the ending color is: " & $avSettings[6] & @CRLF & _
			"The starting color is, in Long Color format: " & $avSettings[7] & @CRLF & _
			"The ending color is, in Long Color format: " & $avSettings[8] & @CRLF & _
			"The starting color intensity percentage is: " & $avSettings[9] & @CRLF & _
			"The ending color intensity percentage is: " & $avSettings[10])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
