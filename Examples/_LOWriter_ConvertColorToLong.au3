#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $iRGB_TO_LONG, $iHEX_TO_LONG, $iCMYK_TO_LONG, $iHSB_TO_LONG

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a Table, 5 rows, 3 columns.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If @error Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	; Insert the Table into the document at the View Cursor's location.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	; Set the table Background color to and set Transparent to False.
	_LOWriter_TableColor($oTable, $LOW_COLOR_MAGENTA, False)
	If @error Then _ERROR("Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I am going to demonstrate how to find the Long color format integer value from R(ed), G(reen), B(lue) values, a Hexadecimal value, " & _
			" C(yan), M(agenta), Y(ellow), and K(ey) values, and H(ue), S(aturation) B(rightness) values, from the color constant $LOW_COLOR_MAGENTA. According to " & _
			"Libre Office, Magenta has the following color values: RGB = R, 191; G, 0; B, 65;" & @CRLF & _
			"Hexadecimal = bf0041" & @CRLF & _
			"CMYK = Cyan, 0; Magenta, 100; Yellow, 66; Key, 25." & @CRLF & _
			"HSB = Hue, 340; Saturation, 100; Brightness, 75;" & @CRLF & @CRLF & _
			"The final total should be 12517441 in Long color format.")

	; Convert RGB to Long Color format, the RGB values are input as integers in their order.
	$iRGB_TO_LONG = _LOWriter_ConvertColorToLong(191, 0, 65)
	If @error Then _ERROR("Failed to convert RGB color value to Long color format integer. Error:" & @error & " Extended:" & @extended)

	; Convert Hex to Long color format, Hex is input as a string.
	$iHEX_TO_LONG = _LOWriter_ConvertColorToLong("bf0041")
	If @error Then _ERROR("Failed to convert HEX color value to Long color format integer. Error:" & @error & " Extended:" & @extended)

	; Convert CMYK to Long Color format, the CMYK values are input as integers in their order.
	$iCMYK_TO_LONG = _LOWriter_ConvertColorToLong(0, 100, 66, 25)
	If @error Then _ERROR("Failed to convert CMYK color value to Long color format integer. Error:" & @error & " Extended:" & @extended)

	; Convert HSB to Long Color format, the HSB values are input as strings in their order.
	$iHSB_TO_LONG = _LOWriter_ConvertColorToLong("340", "100", "75")
	If @error Then _ERROR("Failed to convert HSB color value to Long color format integer. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The expected result is 12517441, the conversion results are as follows: " & @CRLF & _
			"RGB->Long = " & $iRGB_TO_LONG & @CRLF & _
			"Hex->Long = " & $iHEX_TO_LONG & @CRLF & _
			"CMYK->Long = " & $iCMYK_TO_LONG & @CRLF & _
			"HSB->Long = " & $iHSB_TO_LONG & @CRLF & @CRLF & _
			"HSB is a little off, however that is as close as I can mathematically get it. It shouldn't cause a noticable color difference.")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
