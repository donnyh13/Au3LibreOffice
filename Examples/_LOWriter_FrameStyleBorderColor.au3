#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oFrameStyle, $oViewCursor, $oFrame
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Frame Style named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Frame into the document for demonstration.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to create a Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Frame's style to my created style, "Test Style"
	_LOWriter_FrameStyleCurrent($oDoc, $oFrame, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to set Frame style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Frame style Border Width settings to: $LOW_BORDERWIDTH_MEDIUM on all four sides.
	_LOWriter_FrameStyleBorderWidth($oFrameStyle, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR($oDoc, "Failed to modify Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Frame Style Border Color settings to: Top, $LO_COLOR_ORANGE, Bottom $LO_COLOR_BLUE, Left, $LO_COLOR_LGRAY, Right $LO_COLOR_BLACK
	_LOWriter_FrameStyleBorderColor($oFrameStyle, $LO_COLOR_ORANGE, $LO_COLOR_BLUE, $LO_COLOR_LGRAY, $LO_COLOR_BLACK)
	If @error Then _ERROR($oDoc, "Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleBorderColor($oFrameStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame style's current Border Color settings are as follows: " & @CRLF & _
			"The Top Border Color is (as a RGB Color Integer): " & $avSettings[0] & @CRLF & _
			"The Bottom Border Color is (as a RGB Color Integer): " & $avSettings[1] & @CRLF & _
			"The Left Border Color is (as a RGB Color Integer): " & $avSettings[2] & @CRLF & _
			"The Right Border Color is (as a RGB Color Integer): " & $avSettings[3])

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
