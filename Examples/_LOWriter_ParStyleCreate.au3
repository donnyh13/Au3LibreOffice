
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $bExists

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new Paragraph Style to use for demonstration.
	_LOWriter_ParStyleCreate($oDoc, "NewParStyle")
	If (@error > 0) Then _ERROR("Failed to Create a new Paragraph Style. Error:" & @error & " Extended:" & @extended)

	;Check if the paragraph style exists.
	$bExists = _LOWriter_ParStyleExists($oDoc, "NewParStyle")
	If (@error > 0) Then _ERROR("Failed to test for Paragraph Style existing in document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Paragraph style called ""NewParStyle"" exist in the document? True/False: " & $bExists)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

