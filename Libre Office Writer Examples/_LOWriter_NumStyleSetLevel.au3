
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

	;Set the Numbering Style at the View Cursor to Numbering 123.
	_LOWriter_NumStyleSet($oDoc, $oViewCursor, "Numbering 123")
	If (@error > 0) Then _ERROR("Failed to Set the Numbering Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to insert 5 lines and set each to a different level of Numbering.")

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Level 1" & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 2.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 2)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Level 2" & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 3.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 3)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Level 3" & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 4.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 4)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Level 4" & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 5.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 5)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Level 5" & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I could also insert the text first and then set the level.")

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Level 6")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 6.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 6)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "Level 7")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 7.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 7)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "Level 8")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 8.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 8)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "Level 9")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 9.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 9)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "Level 10")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 10.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 10)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "Level 3 again")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 3.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 3)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



