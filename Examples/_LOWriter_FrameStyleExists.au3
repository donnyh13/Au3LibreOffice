#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; See if a Frame Style called "Labels" exists.
	$bReturn = _LOWriter_FrameStyleExists($oDoc, "Labels")
	If @error Then _ERROR("Failed to query for Frame Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a frame style called ""Labels"" exist for this document? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
