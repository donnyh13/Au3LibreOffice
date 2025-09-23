#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iMicrometers
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Footnote separator line settings to: Position = $LOW_ALIGN_HORI_CENTER, Line style = $LOW_LINE_STYLE_DOTTED,
	; Thickness = 1.25 Printer's Points, Color = $LO_COLOR_BLACK, Length = 75%, Spacing to 1/4".
	_LOWriter_PageStyleFootnoteLine($oPageStyle, $LOW_ALIGN_HORI_CENTER, $LOW_LINE_STYLE_DOTTED, 1.25, $LO_COLOR_BLACK, 75, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleFootnoteLine($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Footnote separator line settings are as follows: " & @CRLF & _
			"The Separator line Position is, (see UDF Constants): " & $avPageStyleSettings[0] & @CRLF & _
			"The Separator line Style is, (see UDF Constants): " & $avPageStyleSettings[1] & @CRLF & _
			"The Separator line's thickness is, in Printer's Points: " & $avPageStyleSettings[2] & @CRLF & _
			"The Separator line's Color is, in Long Color Code format: " & $avPageStyleSettings[3] & @CRLF & _
			"The percentage of the Separator line's length, is: " & $avPageStyleSettings[4] & @CRLF & _
			"The distance between the Footnote body and the separator line, in Micrometers, is: " & $avPageStyleSettings[5])

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
