
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $bReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame into the document.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, "AutoitTest", 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	;Check if the document has a Frame by the name of "AutoitTest"
	$bReturn = _LOWriter_DocHasFrameName($oDoc, "AutoitTest")
	If (@error > 0) Then _ERROR("Failed to look for Text Frame name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does this document contain a Frame named ""AutoitTest""? True/ False. " & $bReturn)

	;Delete the Frame.
	_LOWriter_FrameDelete($oDoc, $oFrame)
	If (@error > 0) Then _ERROR("Failed to delete Text Frame. Error:" & @error & " Extended:" & @extended)

	;Check again, if the document has a Frame by the name of "AutoitTest"
	$bReturn = _LOWriter_DocHasFrameName($oDoc, "AutoitTest")
	If (@error > 0) Then _ERROR("Failed to look for Text Frame name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Now does this document contain a Frame named ""AutoitTest""? True/ False. " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



