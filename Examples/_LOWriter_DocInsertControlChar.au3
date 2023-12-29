
#include "LibreOfficeWriter.au3"
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

	;Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate inserting control characters into a document." & @CR & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "This is a $LOW_CON_CHAR_PAR_BREAK --> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a control Character, a Paragraph break.
	_LOWriter_DocInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_PAR_BREAK)
	If (@error > 0) Then _ERROR("Failed to insert Control character. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_LINE_BREAK --> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a control Character, a Line break.
	_LOWriter_DocInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_LINE_BREAK)
	If (@error > 0) Then _ERROR("Failed to insert Control character. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_HARD_HYPHEN --> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a control Character, a hard hyphen.
	_LOWriter_DocInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_HARD_HYPHEN)
	If (@error > 0) Then _ERROR("Failed to insert Control character. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_SOFT_HYPHEN --> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a control Character, a Soft Hyphen.
	_LOWriter_DocInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_SOFT_HYPHEN)
	If (@error > 0) Then _ERROR("Failed to insert Control character. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_HARD_SPACE --> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a control Character, a Hard Space.
	_LOWriter_DocInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_HARD_SPACE)
	If (@error > 0) Then _ERROR("Failed to insert Control character. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_APPEND_PAR --> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	_LOWriter_DocInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_APPEND_PAR)
	If (@error > 0) Then _ERROR("Failed to insert Control character. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

