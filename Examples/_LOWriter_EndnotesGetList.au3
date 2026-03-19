
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $aoEndnotes

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Endnote at the end of this line. ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Endnote at the ViewCursor
	_LOWriter_EndnoteInsert($oDoc, $oViewCursor)
	If (@error > 0) Then _ERROR("Failed to insert a Endnote. Error:" & @error & " Extended:" & @extended)

	;Insert some more text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, " More Text.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert another Endnote at the ViewCursor.
	_LOWriter_EndnoteInsert($oDoc, $oViewCursor)
	If (@error > 0) Then _ERROR("Failed to insert a Endnote. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to delete the first Endnote.")

	;Retrieve an array of Endnote Objects.
	$aoEndnotes = _LOWriter_EndnotesGetList($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve an array of Endnote objects. Error:" & @error & " Extended:" & @extended)

	;Delete the first Endnote returned
	_LOWriter_EndnoteDelete($aoEndnotes[0])
	If (@error > 0) Then _ERROR("Failed to delete a Endnote. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



