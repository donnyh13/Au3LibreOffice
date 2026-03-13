#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oFrameStyle
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Frame Style named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Frame Style options. Set Protect content to True, Protect Position to True, Protect size to True, Vertical alignment to
	; $LOW_PAR_TXT_ADJ_VERT_CENTER, Edit in Read-Only to True, Print to False, Text direction to $LOW_PAR_TXT_DIR_TB_LR
	_LOWriter_FrameStyleOptions($oFrameStyle, True, True, True, $LOW_PAR_TXT_ADJ_VERT_CENTER, True, False, $LOW_PAR_TXT_DIR_TB_LR)
	If @error Then _ERROR($oDoc, "Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleOptions($oFrameStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame style's option settings are as follows: " & @CRLF & _
			"Protect the Frame's contents from changes? True/False: " & $avSettings[0] & @CRLF & _
			"Protect the Frame's position from changes? True/False: " & $avSettings[1] & @CRLF & _
			"Protect the Frame's Size from changes? True/False: " & $avSettings[2] & @CRLF & _
			"The Vertical alignment of the frame is, (see UDF constants): " & $avSettings[3] & @CRLF & _
			"Allow the Frame's contents to be changed in Read-Only mode? True/False: " & $avSettings[4] & @CRLF & _
			"Print frames with this frame style when the document is printed? True/False: " & $avSettings[5] & @CRLF & _
			"The text direction for this frame style is, (See UDF constants): " & $avSettings[6])

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
