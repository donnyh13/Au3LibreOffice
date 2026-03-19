
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I am going to insert enough paragraphs into the document in order to demonstrate moving the view cursor up or down a page etc.")

	;Insert 150 New lines
	For $i = 1 To 150
		_LOWriter_DocInsertString($oDoc, $oViewCursor, "Line " & $i & @CR)
		If (@error > 0) Then _ERROR("Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended)
		Sleep(10)
	Next

	;Return the cursor back to the start.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now move the cursor to the next page.")

	;Move the View Cursor to the next page.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, 1, False)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Next I will move the cursor down one screen space, the same as pressing ""Page Down"".")

	;Move the View Cursor down a screen view space.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_SCREEN_DOWN, 1, False)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Next I will move the cursor to the first page")

	;Move the View Cursor to the first page.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, 1, False)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Next I will move the cursor to the end of the page.")

	;Move the View Cursor to the end of the same page.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, 1, False)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "And finally I will go to a specific page.")

	;Move the View Cursor to a specific page, page 3.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_JUMP_TO_PAGE, 3, False)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



