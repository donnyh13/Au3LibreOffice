
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oBookmark
	Local $asBookmarks

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Bookmark at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a BookMark at the ViewCursor, named "New Bookmark".
	_LOWriter_DocBookmarkInsert($oDoc, $oViewCursor, False, "New Bookmark")
	If (@error > 0) Then _ERROR("Failed to insert a BookMark. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Bookmark's Object.
	$oBookmark = _LOWriter_DocBookmarkGetObj($oDoc, "New Bookmark")
	If (@error > 0) Then _ERROR("Failed to retrieve a BookMark Object. Error:" & @error & " Extended:" & @extended)

	;Modify the Bookmark name to "Old Bookmark"
	_LOWriter_DocBookmarkModify($oDoc, $oBookmark, "Old Bookmark")
	If (@error > 0) Then _ERROR("Failed to modify a Bookmark. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "The Bookmark names currently contained in this document are:" & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Retrieve an array of all Bookmarks.
	$asBookmarks = _LOWriter_DocBookmarksList($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve an array of Bookmarks. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asBookmarks) - 1
		;Insert some text.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asBookmarks[$i] & @CR)
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

