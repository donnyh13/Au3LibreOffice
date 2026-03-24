#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $bResult1, $bResult2

	; Test for a font called "Times New Roman"
	$bResult1 = _LOCalc_FontExists("Times New Roman")
	If @error Then _ERROR("Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test for a font called "Fake Font"
	$bResult2 = _LOCalc_FontExists("Fake Font")
	If @error Then _ERROR("Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the font called ""Times New Roman"" available? True/False: " & $bResult1 & @CRLF & @CRLF & _
			"Is the font called ""Fake Font"" available? True/False: " & $bResult2)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
If @error Then Return _ERROR("Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
