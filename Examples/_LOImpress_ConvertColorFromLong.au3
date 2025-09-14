#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $sHex
	Local $aiRGB, $aiCMYK, $aiHSB

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I am going to demonstrate how to convert the Long color format integer value, $LOI_COLOR_MAGENTA (12517441), into R(ed), G(reen), " & _
			"B(lue) values, a Hexadecimal value, C(yan), M(agenta), Y(ellow), and K(ey) values, and H(ue), S(aturation) B(rightness) values.")

	; Convert to RGB From Long Color format, the RGB values are returned as an array in their order.
	$aiRGB = _LOImpress_ConvertColorFromLong(Null, $LOI_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to RGB color value from Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert to Hex From Long color format, Hex is returned as a string.
	$sHex = _LOImpress_ConvertColorFromLong($LOI_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to HEX color value from Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert to CMYK From Long Color format, the CMYK values are returned as an array in their order.
	$aiCMYK = _LOImpress_ConvertColorFromLong(Null, Null, Null, $LOI_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to CMYK color value from Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert to HSB From Long Color format, the HSB values are returned as an array in their order.
	$aiHSB = _LOImpress_ConvertColorFromLong(Null, Null, $LOI_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to HSB color value from Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The conversion results are as follows: " & @CRLF & _
			"Long->RGB = " & "R, " & $aiRGB[0] & "; G, " & $aiRGB[1] & "; B, " & $aiRGB[2] & @CRLF & " Should be: R, 191; G, 0; B, 65" & @CRLF & @CRLF & _
			"Long->Hex = " & $sHex & @CRLF & " Should be bf0041" & @CRLF & @CRLF & _
			"Long->CMYK = " & "C, " & $aiCMYK[0] & "; M " & $aiCMYK[1] & "; Y " & $aiCMYK[2] & "; K " & $aiCMYK[3] & @CRLF & " Should be: C, 0; M, 100; Y, 66; K, 25." & @CRLF & @CRLF & _
			"Long->HSB = " & "H, " & $aiHSB[0] & "; S " & $aiHSB[1] & "; B " & $aiHSB[2] & @CRLF & " Should be: H, 340; S, 100; B, 75")
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
