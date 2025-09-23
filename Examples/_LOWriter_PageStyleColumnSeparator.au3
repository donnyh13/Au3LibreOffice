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

	; Convert 1/16" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(.0625)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Column count to 4.
	_LOWriter_PageStyleColumnSettings($oPageStyle, 4)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Column Separator line settings to: Separator on (True), Line Style = $LOW_LINE_STYLE_SOLID, Line width to 1/16"
	; Line Color to $LO_COLOR_RED, Height to 75%, Line Position to $LOW_ALIGN_VERT_MIDDLE
	_LOWriter_PageStyleColumnSeparator($oPageStyle, True, $LOW_LINE_STYLE_SOLID, $iMicrometers, $LO_COLOR_RED, 75, $LOW_ALIGN_VERT_MIDDLE)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleColumnSeparator($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Column Separator Line settings are as follows: " & @CRLF & _
			"Is Column separated by a line? True/False: " & $avPageStyleSettings[0] & @CRLF & _
			"The Separator Line style is, (see UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
			"The Separator Line width is, in Micrometers: " & $avPageStyleSettings[2] & @CRLF & _
			"The Separator Line color is, in Long color format: " & $avPageStyleSettings[3] & @CRLF & _
			"The Separator Line length percentage is: " & $avPageStyleSettings[4] & @CRLF & _
			"The Separator Line position is, (see UDF constants): " & $avPageStyleSettings[5])

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
