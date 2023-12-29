
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oCharStyle
	Local $bExists

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new Character Style for demonstration.
	$oCharStyle = _LOWriter_CharStyleCreate($oDoc, "NewCharStyle")
	If (@error > 0) Then _ERROR("Failed to create Character style. Error:" & @error & " Extended:" & @extended)

	;Check if the Character style exists.
	$bExists = _LOWriter_CharStyleExists($oDoc, "NewCharStyle")
	If (@error > 0) Then _ERROR("Failed to test for Character Style existing in document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Character style called ""NewCharStyle"" exist in the document? True/False: " & $bExists)

	;Delete the Character style, Force delete if it is in use, and replace it with Character style "Example".
	_LOWriter_CharStyleDelete($oDoc, $oCharStyle, True, "Example")

	;Check if the Character style still exists.
	$bExists = _LOWriter_CharStyleExists($oDoc, "NewCharStyle")
	If (@error > 0) Then _ERROR("Failed to test for Character Style existing in document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Character style called ""NewCharStyle"" still exist in the document? True/False: " & $bExists)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



