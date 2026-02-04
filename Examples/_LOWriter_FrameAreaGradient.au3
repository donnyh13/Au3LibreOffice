#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Frame into the document at the ViewCursor position, and 3000x3000 Hundredths of a Millimeter (HMM) wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to create a Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Frame Gradient settings to: Preset Gradient name = $LOW_GRAD_NAME_TEAL_TO_BLUE
	_LOWriter_FrameAreaGradient($oDoc, $oFrame, $LOW_GRAD_NAME_TEAL_TO_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameAreaGradient($oDoc, $oFrame)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame's Gradient settings are as follows: " & @CRLF & _
			"The Gradient name is: " & $avSettings[0] & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The number of steps to increment color is: " & $avSettings[2] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avSettings[3] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avSettings[4] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avSettings[5] & @CRLF & _
			"The percentage of area not covered by the ending color is: " & $avSettings[6] & @CRLF & _
			"The starting color is (as a RGB Color Integer): " & $avSettings[7] & @CRLF & _
			"The ending color is (as a RGB Color Integer): " & $avSettings[8] & @CRLF & _
			"The starting color intensity percentage is: " & $avSettings[9] & @CRLF & _
			"The ending color intensity percentage is: " & $avSettings[10])

	; Modify the Frame Gradient settings to: skip pre-set gradient name, Gradient type = $LOW_GRAD_TYPE_SQUARE, increment steps = 150,
	; horizontal (X) offset = 25%, vertical offset (Y) = 56%, rotational angle = 135 degrees, percentage not covered by "From" color = 50%
	; Starting color = $LO_COLOR_ORANGE, Ending color = $LO_COLOR_TEAL, Starting color intensity = 100%, ending color intensity = 68%
	_LOWriter_FrameAreaGradient($oDoc, $oFrame, Null, $LOW_GRAD_TYPE_SQUARE, 150, 25, 56, 135, 50, $LO_COLOR_ORANGE, $LO_COLOR_TEAL, 100, 68)
	If @error Then _ERROR($oDoc, "Failed to set Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameAreaGradient($oDoc, $oFrame)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame's Gradient settings are as follows: " & @CRLF & _
			"The Gradient name is: " & $avSettings[0] & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The number of steps to increment color is: " & $avSettings[2] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avSettings[3] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avSettings[4] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avSettings[5] & @CRLF & _
			"The percentage of area not covered by the ending color is: " & $avSettings[6] & @CRLF & _
			"The starting color is (as a RGB Color Integer): " & $avSettings[7] & @CRLF & _
			"The ending color is (as a RGB Color Integer): " & $avSettings[8] & @CRLF & _
			"The starting color intensity percentage is: " & $avSettings[9] & @CRLF & _
			"The ending color intensity percentage is: " & $avSettings[10])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
