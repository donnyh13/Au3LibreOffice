
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTextCursor

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text using the View Cursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text." & @CR & @CR & "Some different text" & @CR & "Another Line.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Create a new Text Cursor.
	$oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended)

	;Insert some text using the Text Cursor
	_LOWriter_DocInsertString($oDoc, $oTextCursor, ">[This is where the Text Cursor currently is.]< ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Move the TextCursor to where the View Cursor is, do not select text on the way.
	_LOWriter_CursorGoToRange($oTextCursor, $oViewCursor, False)
	If (@error > 0) Then _ERROR("Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	;Insert more text using the Text Cursor
	_LOWriter_DocInsertString($oDoc, $oTextCursor, " >[This is where the Text Cursor now is, after moving it to the ViewCursor's position.]<")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



