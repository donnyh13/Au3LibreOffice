
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oBookmark, $oBookAnchor

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Bookmark at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Bookmark at the ViewCursor, named "New Bookmark".
	$oBookmark = _LOWriter_DocBookmarkInsert($oDoc, $oViewCursor, False, "New Bookmark")
	If (@error > 0) Then _ERROR("Failed to insert a BookMark. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will retrieve the anchor for ""New Bookmark"" and insert some text after it.")

	;Retrieve the Bookmark's Anchor (Text Cursor).
	$oBookAnchor = _LOWriter_DocBookmarkGetAnchor($oBookmark)
	If (@error > 0) Then _ERROR("Failed to Retrieve a Reference Mark anchor. Error:" & @error & " Extended:" & @extended)

	;Move the Anchor (Text Cursor)
	_LOWriter_CursorMove($oBookAnchor, $LOW_TEXTCUR_GO_RIGHT, 1, False)
	If (@error > 0) Then _ERROR("Failed to move a cursor. Error:" & @error & " Extended:" & @extended)

	;Insert Some text.
	_LOWriter_DocInsertString($oDoc, $oBookAnchor, " Some new text")
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



