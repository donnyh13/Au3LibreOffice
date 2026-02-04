#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $sHex
	Local $aiRGB, $aiCMYK, $aiHSB

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I am going to demonstrate how to convert the RGB Color Integer, $LO_COLOR_MAGENTA (12517441), into R(ed), G(reen), " & _
			"B(lue) values, a Hexadecimal value, C(yan), M(agenta), Y(ellow), and K(ey) values, and H(ue), S(aturation) B(rightness) values.")

	; Convert to RGB From a RGB Color Integer, the RGB values are returned as an array in their order.
	$aiRGB = _LO_ConvertColorFromLong(Null, $LO_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to RGB color value from a RGB Color Integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert to Hex From a RGB Color Integer, Hex is returned as a string.
	$sHex = _LO_ConvertColorFromLong($LO_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to HEX color value from a RGB Color Integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert to CMYK From a RGB Color Integer, the CMYK values are returned as an array in their order.
	$aiCMYK = _LO_ConvertColorFromLong(Null, Null, Null, $LO_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to CMYK color value from a RGB Color Integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert to HSB a RGB Color Integer, the HSB values are returned as an array in their order.
	$aiHSB = _LO_ConvertColorFromLong(Null, Null, $LO_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to HSB color value from a RGB Color Integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The conversion results are as follows: " & @CRLF & _
			"RGB Integer->RGB = " & "R, " & $aiRGB[0] & "; G, " & $aiRGB[1] & "; B, " & $aiRGB[2] & @CRLF & " Should be: R, 191; G, 0; B, 65" & @CRLF & @CRLF & _
			"RGB Integer->Hex = " & $sHex & @CRLF & " Should be bf0041" & @CRLF & @CRLF & _
			"RGB Integer->CMYK = " & "C, " & $aiCMYK[0] & "; M " & $aiCMYK[1] & "; Y " & $aiCMYK[2] & "; K " & $aiCMYK[3] & @CRLF & " Should be: C, 0; M, 100; Y, 66; K, 25." & @CRLF & @CRLF & _
			"RGB Integer->HSB = " & "H, " & $aiHSB[0] & "; S " & $aiHSB[1] & "; B " & $aiHSB[2] & @CRLF & " Should be: H, 340; S, 100; B, 75")
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
