#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Minimize the document
	_LOWriter_DocMinimize($oDoc, True)
	If @error Then _ERROR("Failed to Minimize Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to bring the document to the front.")

	_LOWriter_DocToFront($oDoc)
	If @error Then _ERROR("Failed to bring document to the front. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
