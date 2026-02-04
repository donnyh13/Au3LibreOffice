#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oParStyle
	Local $iHMM, $iHMM2, $iHMM3
	Local $avParStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the Default Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style." & @LF & "Next Line" & @CR & "Next Line" & @LF)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Paragraph Style object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(0.25, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2" to Hundredths of a Millimeter (HMM)
	$iHMM2 = _LO_UnitConvert(0.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1" to Hundredths of a Millimeter (HMM)
	$iHMM3 = _LO_UnitConvert(1, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Default Paragraph Style Indent settings to, 1/4" Before paragraph, 1/2" after paragraph, First line indent = 1",
	; Auto First line = False
	_LOWriter_ParStyleIndent($oParStyle, $iHMM, $iHMM2, $iHMM3, False)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avParStyleSettings = _LOWriter_ParStyleIndent($oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Indent settings are as follows: " & @CRLF & _
			"Indent spacing before the paragraph, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[0] & @CRLF & _
			"Spacing after the paragraph, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[1] & @CRLF & _
			"First line Indent spacing, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[2] & @CRLF & _
			"Automatically indent First line? True/False: " & $avParStyleSettings[3])

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
