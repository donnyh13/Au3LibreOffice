#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $sHex
	Local $aiRGB, $aiCMYK, $aiHSB

	MsgBox($MB_OK, "", "I am going to demonstrate how to convert the Long color format integer value, $LOC_COLOR_MAGENTA (12517441), into R(ed), G(reen), " & _
			"B(lue) values, a Hexadecimal value, C(yan), M(agenta), Y(ellow), and K(ey) values, and H(ue), S(aturation) B(rightness) values.")

	; Convert to RGB From Long Color format, the RGB values are returned as an array in their order.
	$aiRGB = _LOCalc_ConvertColorFromLong(Null, $LOC_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to RGB color value from Long color format integer. Error:" & @error & " Extended:" & @extended)

	; Convert to Hex From Long color format, Hex is returned as a string.
	$sHex = _LOCalc_ConvertColorFromLong($LOC_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to HEX color value from Long color format integer. Error:" & @error & " Extended:" & @extended)

	; Convert to CMYK From Long Color format, the CMYK values are returned as an array in their order.
	$aiCMYK = _LOCalc_ConvertColorFromLong(Null, Null, Null, $LOC_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to CMYK color value from Long color format integer. Error:" & @error & " Extended:" & @extended)

	; Convert to HSB From Long Color format, the HSB values are returned as an array in their order.
	$aiHSB = _LOCalc_ConvertColorFromLong(Null, Null, $LOC_COLOR_MAGENTA)
	If @error Then _ERROR("Failed to convert to HSB color value from Long color format integer. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The conversion results are as follows: " & @CRLF & _
			"Long->RGB = " & "R, " & $aiRGB[0] & "; G, " & $aiRGB[1] & "; B, " & $aiRGB[2] & " Should be: R, 191; G, 0; B, 65" & @CRLF & _
			"Long->Hex = " & $sHex & " Should be bf0041" & @CRLF & _
			"Long->CMYK = " & "C, " & $aiCMYK[0] & "; M " & $aiCMYK[1] & "; Y " & $aiCMYK[2] & "; K " & $aiCMYK[3] & " Should be: C, 0; M, 100; Y, 66; K, 25." & @CRLF & _
			"Long->HSB = " & "H, " & $aiHSB[0] & "; S " & $aiHSB[1] & "; B " & $aiHSB[2] & " Should be: H, 340; S, 100; B, 75")

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
