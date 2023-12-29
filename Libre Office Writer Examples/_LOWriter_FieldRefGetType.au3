
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a reference Mark at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Reference Mark at the ViewCursor, named "Ref. 1".
	_LOWriter_FieldRefMarkSet($oDoc, $oViewCursor, "Ref. 1", False)
	If (@error > 0) Then _ERROR("Failed to insert a Reference Mark. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "I have inserted a field at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Reference Field at the View Cursor. Reference Ref. 1 Reference mark, do not overwrite any content selected by the cursor, and refer using
	;$LOW_FIELD_REF_USING_PAGE_NUM_STYLED
	$oField = _LOWriter_FieldRefInsert($oDoc, $oViewCursor, "Ref. 1", False, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED)
	If (@error > 0) Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Reference Field is referencing following type: " & _LOWriter_FieldRefGetType($oField) & @CRLF & @CRLF & _
			"The Possible types are: " & @CRLF & "$LOW_FIELD_REF_TYPE_REF_MARK(0)" & @CRLF & "$LOW_FIELD_REF_TYPE_SEQ_FIELD(1)" & @CRLF & _
			"$LOW_FIELD_REF_TYPE_BOOKMARK(2)" & @CRLF & "$LOW_FIELD_REF_TYPE_FOOTNOTE(3)" & @CRLF & "$LOW_FIELD_REF_TYPE_ENDNOTE(4)")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



