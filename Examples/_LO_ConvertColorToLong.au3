#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $iRGB_TO_LONG, $iHEX_TO_LONG, $iCMYK_TO_LONG, $iHSB_TO_LONG

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I am going to demonstrate how to find the Long color format integer value from R(ed), G(reen), B(lue) values, a Hexadecimal value, " & _
			" C(yan), M(agenta), Y(ellow), and K(ey) values, and H(ue), S(aturation) B(rightness) values, from the color constant $LO_COLOR_MAGENTA. According to " & _
			"Libre Office, Magenta has the following color values: RGB = R, 191; G, 0; B, 65;" & @CRLF & _
			"Hexadecimal = bf0041" & @CRLF & _
			"CMYK = Cyan, 0; Magenta, 100; Yellow, 66; Key, 25." & @CRLF & _
			"HSB = Hue, 340; Saturation, 100; Brightness, 75;" & @CRLF & @CRLF & _
			"The final total should be 12517441 in Long color format.")

	; Convert RGB to Long Color format, the RGB values are input as integers in their order.
	$iRGB_TO_LONG = _LO_ConvertColorToLong(191, 0, 65)
	If @error Then _ERROR("Failed to convert RGB color value to Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert Hex to Long color format, Hex is input as a string.
	$iHEX_TO_LONG = _LO_ConvertColorToLong("bf0041")
	If @error Then _ERROR("Failed to convert HEX color value to Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert CMYK to Long Color format, the CMYK values are input as integers in their order.
	$iCMYK_TO_LONG = _LO_ConvertColorToLong(0, 100, 66, 25)
	If @error Then _ERROR("Failed to convert CMYK color value to Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert HSB to Long Color format, the HSB values are input as strings in their order.
	$iHSB_TO_LONG = _LO_ConvertColorToLong("340", "100", "75")
	If @error Then _ERROR("Failed to convert HSB color value to Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The expected result is 12517441, the conversion results are as follows: " & @CRLF & _
			"RGB->Long = " & $iRGB_TO_LONG & @CRLF & _
			"Hex->Long = " & $iHEX_TO_LONG & @CRLF & _
			"CMYK->Long = " & $iCMYK_TO_LONG & @CRLF & _
			"HSB->Long = " & $iHSB_TO_LONG & @CRLF & @CRLF & _
			"HSB is a little off, however that is as close as I can mathematically get it. It shouldn't cause a noticeable color difference.")
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
