#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly." & @CR & _
			"A second paragraph to demonstrate modifying formatting settings directly." & @CR & _
			"A third paragraph to demonstrate modifying formatting settings directly.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor down one paragraph
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_DOWN, 1)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Paragraph Gradient settings to: Preset Gradient name = $LOW_GRAD_NAME_TEAL_TO_BLUE
	_LOWriter_DirFrmtParAreaGradient($oDoc, $oViewCursor, $LOW_GRAD_NAME_TEAL_TO_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set Paragraph settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Paragraph settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_DirFrmtParAreaGradient($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's Gradient settings are as follows: " & @CRLF & _
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

	; Modify the Paragraph Gradient settings to: skip pre-set gradient name, Gradient type = $LOW_GRAD_TYPE_SQUARE, increment steps = 150,
	; horizontal (X) offset = 25%, vertical offset (Y) = 56%, rotational angle = 135 degrees, percentage not covered by "From" color = 50%
	; Starting color = $LO_COLOR_ORANGE, Ending color = $LO_COLOR_TEAL, Starting color intensity = 100%, ending color intensity = 68%
	_LOWriter_DirFrmtParAreaGradient($oDoc, $oViewCursor, Null, $LOW_GRAD_TYPE_SQUARE, 150, 25, 56, 135, 50, $LO_COLOR_ORANGE, $LO_COLOR_TEAL, 100, 68)
	If @error Then _ERROR($oDoc, "Failed to set Paragraph settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Paragraph settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_DirFrmtParAreaGradient($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's Gradient settings are as follows: " & @CRLF & _
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
