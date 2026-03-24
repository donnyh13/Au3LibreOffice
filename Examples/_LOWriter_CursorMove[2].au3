#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I am going to insert enough paragraphs into the document in order to demonstrate moving the view cursor up or down a page etc.")

	; Insert 150 New lines
	For $i = 1 To 150
		_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Line " & $i & @CR)
		If @error Then _ERROR($oDoc, "Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		Sleep(10)
	Next

	; Return the cursor back to the start.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now move the cursor to the next page.")

	; Move the View Cursor to the next page.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, 1, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next I will move the cursor down one screen space, the same as pressing ""Page Down"".")

	; Move the View Cursor down a screen view space.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_SCREEN_DOWN, 1, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next I will move the cursor to the first page")

	; Move the View Cursor to the first page.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, 1, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next I will move the cursor to the end of the page.")

	; Move the View Cursor to the end of the same page.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, 1, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "And finally I will go to a specific page.")

	; Move the View Cursor to a specific page, page 3.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_JUMP_TO_PAGE, 3, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
