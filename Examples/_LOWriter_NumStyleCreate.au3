#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $bReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new NumberingStyle named "Test Style"
	_LOWriter_NumStyleCreate($oDoc, "Test Style")
	If (@error > 0) Then _ERROR("Failed to create a Numbering Style. Error:" & @error & " Extended:" & @extended)

	;See if a Numbering Style called "Test Style" exists.
	$bReturn = _LOWriter_NumStyleExists($oDoc, "Test Style")
	If (@error > 0) Then _ERROR("Failed to query for Numbering Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Numbering style called ""Test Style"" exist for this document? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
