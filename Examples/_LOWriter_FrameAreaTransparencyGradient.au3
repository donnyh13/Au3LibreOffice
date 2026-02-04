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

	; Modify the Frame  Background Color settings. Background color = $LO_COLOR_TEAL
	_LOWriter_FrameAreaColor($oFrame, $LO_COLOR_TEAL)
	If @error Then _ERROR($oDoc, "Failed to set Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Frame Transparency Gradient settings to: Gradient Type = $LOW_GRAD_TYPE_ELLIPTICAL, XCenter to 75%, YCenter to 45%, Angle to 180 degrees
	; Border to 16%, Start transparency to 10%, End Transparency to 62%
	_LOWriter_FrameAreaTransparencyGradient($oDoc, $oFrame, $LOW_GRAD_TYPE_ELLIPTICAL, 75, 45, 180, 16, 10, 62)
	If @error Then _ERROR($oDoc, "Failed to set Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameAreaTransparencyGradient($oDoc, $oFrame)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame's Transparency Gradient settings are as follows: " & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avSettings[1] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avSettings[2] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avSettings[3] & @CRLF & _
			"The percentage of area not covered by the transparency is: " & $avSettings[4] & @CRLF & _
			"The starting transparency percentage is: " & $avSettings[5] & @CRLF & _
			"The ending transparency percentage is: " & $avSettings[6])

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
