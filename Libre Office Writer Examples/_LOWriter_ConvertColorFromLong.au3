
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $sHex
	Local $aiRGB, $aiCMYK, $aiHSB

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a Table, 5 rows, 3 columns.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document at the View Cursor's location.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Set the table Background color to and set Transparent to False.
	_LOWriter_TableColor($oTable, $LOW_COLOR_MAGENTA, False)
	If (@error > 0) Then _ERROR("Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I am going to demonstrate how to convert the Long color format integer value, $LOW_COLOR_MAGENTA (12517441), into R(ed), G(reen), " & _
			"B(lue) values, a Hexadecimal value, C(yan), M(agenta), Y(ellow), and K(ey) values, and H(ue), S(aturation) B(rightness) values.")

	;Convert to RGB From Long Color format, the RGB values are returned as an array in their order.
	$aiRGB = _LOWriter_ConvertColorFromLong(Null, $LOW_COLOR_MAGENTA)
	If (@error > 0) Then _ERROR("Failed to convert to RGB color value from Long color format integer. Error:" & @error & " Extended:" & @extended)

	;Convert to Hex From Long color format, Hex is returned as a string.
	$sHex = _LOWriter_ConvertColorFromLong($LOW_COLOR_MAGENTA)
	If (@error > 0) Then _ERROR("Failed to convert to HEX color value from Long color format integer. Error:" & @error & " Extended:" & @extended)

	;Convert to CMYK From Long Color format, the CMYK values are returned as an array in their order.
	$aiCMYK = _LOWriter_ConvertColorFromLong(Null, Null, Null, $LOW_COLOR_MAGENTA)
	If (@error > 0) Then _ERROR("Failed to convert to CMYK color value from Long color format integer. Error:" & @error & " Extended:" & @extended)

	;Convert to HSB From Long Color format, the HSB values are returned as an array in their order.
	$aiHSB = _LOWriter_ConvertColorFromLong(Null, Null, $LOW_COLOR_MAGENTA)
	If (@error > 0) Then _ERROR("Failed to convert to HSB color value from Long color format integer. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The conversion results are as follows: " & @CRLF & _
			"Long->RGB = " & "R, " & $aiRGB[0] & "; G, " & $aiRGB[1] & "; B, " & $aiRGB[2] & " Should be: R, 191; G, 0; B, 65" & @CRLF & _
			"Long->Hex = " & $sHex & " Should be bf0041" & @CRLF & _
			"Long->CMYK = " & "C, " & $aiCMYK[0] & "; M " & $aiCMYK[1] & "; Y " & $aiCMYK[2] & "; K " & $aiCMYK[3] & " Should be: C, 0; M, 100; Y, 66; K, 25." & @CRLF & _
			"Long->HSB = " & "H, " & $aiHSB[0] & "; S " & $aiHSB[1] & "; B " & $aiHSB[2] & " Should be: H, 340; S, 100; B, 75")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



