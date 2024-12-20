#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $bResult1, $bResult2

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test for a font called "Times New Roman"
	$bResult1 = _LOWriter_FontExists($oDoc, "Times New Roman")
	If @error Then _ERROR($oDoc, "Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test for a font called "Fake Font"
	$bResult2 = _LOWriter_FontExists($oDoc, "Fake Font")
	If @error Then _ERROR($oDoc, "Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does the document have a font called ""Times New Roman"" ? True/False: " & $bResult1 & @CRLF & @CRLF & _
			"Does the document have a font called ""Fake Font"" ? True/False: " & $bResult2)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
