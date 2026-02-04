#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iColor

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

	; Set the paragraph at the current cursor's location background color to $LO_COLOR_GRAY.
	_LOWriter_DirFrmtParAreaColor($oViewCursor, $LO_COLOR_GRAY)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an Integer.
	$iColor = _LOWriter_DirFrmtParAreaColor($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Paragraph Background color settings are as follows is (as a RGB Color Integer): " & $iColor & @CRLF & @CRLF & _
			"Press ok to remove direct formatting.")

	; Remove Direct formatting.
	_LOWriter_DirFrmtParAreaColor($oViewCursor, Null, True)
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
