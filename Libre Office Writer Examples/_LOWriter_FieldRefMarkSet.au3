
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asRefs

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
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "I have inserted a reference Mark at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Reference Mark at the ViewCursor, named "Ref. 2".
	_LOWriter_FieldRefMarkSet($oDoc, $oViewCursor, "Ref. 2", False)
	If (@error > 0) Then _ERROR("Failed to insert a Reference Mark. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "I have inserted a reference Mark at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Reference Mark at the ViewCursor, named "Ref. 3".
	_LOWriter_FieldRefMarkSet($oDoc, $oViewCursor, "Ref. 3", False)
	If (@error > 0) Then _ERROR("Failed to insert a Reference Mark. Error:" & @error & " Extended:" & @extended)

	;Insert a new Line.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "The Reference Mark names contained in this document are: " & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Retrieve an Array of Reference Mark Names.
	$asRefs = _LOWriter_FieldRefMarkList($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve an Array of Reference Marks. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To @extended - 1
		;Insert the Reference Mark Names.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & $asRefs[$i])
		If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



