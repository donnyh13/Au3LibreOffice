
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $bReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to Maximize the document.")

	;Maximize the document.
	_LOWriter_DocMaximize($oDoc, True)
	If (@error > 0) Then _ERROR("Failed to Maximize Document. Error:" & @error & " Extended:" & @extended)

	;Test If document is currently maximized.
	$bReturn = _LOWriter_DocMaximize($oDoc)
	If (@error > 0) Then _ERROR("Failed to query Document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the document currently maximized? True/False: " & $bReturn & @CRLF & _
			"Press Ok to restore the document to it's previous size and position.")

	;Restore the document to it's previous size.
	_LOWriter_DocMaximize($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to restore Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



