#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oParStyle
	Local $iHMM, $iHMM2
	Local $avParStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the Default Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style." & @CR & "Next Line")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Paragraph Style object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Default Paragraph Style Border Width to $LOW_BORDERWIDTH_THICK.
	_LOWriter_ParStyleBorderWidth($oParStyle, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(0.25, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Default Paragraph Style Border padding to 1/4"
	_LOWriter_ParStyleBorderPadding($oParStyle, $iHMM)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avParStyleSettings = _LOWriter_ParStyleBorderPadding($oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Border padding distance settings are as follows: " & @CRLF & _
			"All Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[1] & @CRLF & _
			"Bottom Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[2] & @CRLF & _
			"Left Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[3] & @CRLF & _
			"Right Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[4] & @CRLF & @CRLF & _
			"Press Ok, and I will demonstrate setting individual border padding settings.")

	; Convert 1/2" to Hundredths of a Millimeter (HMM)
	$iHMM2 = _LO_UnitConvert(0.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Default Paragraph Style Border padding to, Top and Right, 1/4", Bottom and left, 1/2".
	_LOWriter_ParStyleBorderPadding($oParStyle, Null, $iHMM, $iHMM2, $iHMM2, $iHMM)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avParStyleSettings = _LOWriter_ParStyleBorderPadding($oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Border padding distance settings are as follows: " & @CRLF & _
			"All Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[1] & @CRLF & _
			"Bottom Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[2] & @CRLF & _
			"Left Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[3] & @CRLF & _
			"Right Padding distance, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[4])

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
