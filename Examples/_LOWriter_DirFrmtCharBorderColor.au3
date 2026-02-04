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
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor right 13 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 13)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select the word "Demonstrate".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 11, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the selected text's Border Width to $LOW_BORDERWIDTH_THICK.
	_LOWriter_DirFrmtCharBorderWidth($oViewCursor, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the selected text's Border color to $LO_COLOR_BRICK
	_LOWriter_DirFrmtCharBorderColor($oViewCursor, $LO_COLOR_BRICK, $LO_COLOR_BRICK, $LO_COLOR_BRICK, $LO_COLOR_BRICK)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_DirFrmtCharBorderColor($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The selected text's current Character Border color settings are as follows: " & @CRLF & _
			"Top color is (as a RGB Color Integer): " & $avSettings[0] & @CRLF & _
			"Bottom color is (as a RGB Color Integer): " & $avSettings[1] & @CRLF & _
			"Left color is (as a RGB Color Integer): " & $avSettings[2] & @CRLF & _
			"Right color is (as a RGB Color Integer): " & $avSettings[3] & @CRLF & @CRLF & _
			"Press ok to remove direct formatting.")

	; Remove direct formatting from selection.
	_LOWriter_DirFrmtCharBorderColor($oViewCursor, Null, Null, Null, Null, True)
	If @error Then _ERROR($oDoc, "Failed to clear the selected text's direct formatting settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
