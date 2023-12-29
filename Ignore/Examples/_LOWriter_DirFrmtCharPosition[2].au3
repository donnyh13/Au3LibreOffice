#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly.")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Move the cursor right 13 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 13)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Select the word "Demonstrate".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 11, True)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Set the selected text's font position to 75% Superscript, and relative size to 50%.
	_LOWriter_DirFrmtCharPosition($oViewCursor, Null, 75, Null, Null, 50)
	If @error Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtCharPosition($oViewCursor)
	If @error Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The selected text's current position settings are as follows: " & @CRLF & _
			"Is Auto-SuperScript? True/False: " & $avSettings[0] & @CRLF & _
			"Current SuperScript percentage (If Auto, then it will be 14000): " & $avSettings[1] & @CRLF & _
			"Is Auto-SubScript? True/False: " & $avSettings[2] & @CRLF & _
			"Current SubScript percentage (If Auto, then it will be -14000): " & $avSettings[3] & @CRLF & _
			"Relative size percentage: " & $avSettings[4] & @CRLF & @CRLF & _
			"Press ok and I will set SubScript next.")

	; Set selected text's font position to 75 Subscript
	_LOWriter_DirFrmtCharPosition($oViewCursor, Null, Null, Null, 75)
	If @error Then _ERROR("Failed to set the selected text's settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtCharPosition($oViewCursor)
	If @error Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The selected text's new position settings are as follows: " & @CRLF & _
			"Is Auto-SuperScript? True/False: " & $avSettings[0] & @CRLF & _
			"Current SuperScript percentage (If Auto, then it will be 14000): " & $avSettings[1] & @CRLF & _
			"Is Auto-SubScript? True/False: " & $avSettings[2] & @CRLF & _
			"Current SubScript percentage (If Auto, then it will be -14000): " & $avSettings[3] & @CRLF & _
			"Relative size percentage: " & $avSettings[4] & @CRLF & @CRLF & _
			"Press ok to remove direct formating.")

	; Remove Direct formatting.
	_LOWriter_DirFrmtCharPosition($oViewCursor, Null, Null, Null, Null, Null, True)
	If @error Then _ERROR("Failed to clear the selected text's direct formatting settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
