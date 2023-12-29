
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the View Cursor for later use
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve ViewCursor object. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the View Cursor's position.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I inserted text where the ViewCursor was.")
	If (@error > 0) Then _ERROR("Failed to insert text at the ViewCursor's position. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



