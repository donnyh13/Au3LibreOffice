#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $bExists

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a new Page Style to use for demonstration.
	_LOWriter_PageStyleCreate($oDoc, "NewPageStyle")
	If @error Then _ERROR("Failed to Create a new Page Style. Error:" & @error & " Extended:" & @extended)

	; Check if the page style exists.
	$bExists = _LOWriter_PageStyleExists($oDoc, "NewPageStyle")
	If @error Then _ERROR("Failed to test for Page Style existing in document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Page style called ""NewPageStyle"" exist in the document? True/False: " & $bExists)

	; Check if a fake page style exists.
	$bExists = _LOWriter_PageStyleExists($oDoc, "FakePageStyle")
	If @error Then _ERROR("Failed to test for Page Style existing in document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Page style called ""FakePageStyle"" exist in the document? True/False: " & $bExists)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
