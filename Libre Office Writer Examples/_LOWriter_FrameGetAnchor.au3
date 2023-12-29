
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame, $oAnchorTextCursor

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame into the document at the Viewcursor position, Named "AutoitTest", and 3000x3000 Micrometers wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, "AutoitTest", 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	;Create a Text Cursor at the Anchor location.
	$oAnchorTextCursor = _LOWriter_FrameGetAnchor($oFrame)
	If (@error > 0) Then _ERROR("Failed to create a text Cursor. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the Frame anchor location.
	_LOWriter_DocInsertString($oDoc, $oAnchorTextCursor, "(NEW TEXT INSERTED USING THE ANCHOR CURSOR.) ")
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



