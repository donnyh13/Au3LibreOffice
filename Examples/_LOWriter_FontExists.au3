
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $bResult1, $bResult2

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Test for a font called "Times New Roman"
	$bResult1 = _LOWriter_FontExists($oDoc, "Times New Roman")
	If (@error > 0) Then _ERROR("Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended)

	;Test for a font called "Fake Font"
	$bResult2 = _LOWriter_FontExists($oDoc, "Fake Font")
	If (@error > 0) Then _ERROR("Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does the document have a font called ""Times New Roman"" ? True/False: " & $bResult1 & @CRLF & @CRLF & _
			"Does the document have a font called ""Fake Font"" ? True/False: " & $bResult2)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)

	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

