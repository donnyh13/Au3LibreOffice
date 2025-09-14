#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $bResult1, $bResult2

	; Test for a font called "Times New Roman"
	$bResult1 = _LOImpress_FontExists("Times New Roman")
	If @error Then _ERROR("Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test for a font called "Fake Font"
	$bResult2 = _LOImpress_FontExists("Fake Font")
	If @error Then _ERROR("Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the font called ""Times New Roman"" available? True/False: " & $bResult1 & @CRLF & @CRLF & _
			"Is the font called ""Fake Font"" available? True/False: " & $bResult2)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
