#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oNumStyle, $oViewCursor
	Local $iMicrometers, $iMicrometers2, $iMicrometers3
	Local $aavSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Numbering Style named "Test Style"
	$oNumStyle = _LOWriter_NumStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to create a Numbering Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Numbering Style at the View Cursor to the new style.
	_LOWriter_NumStyleSet($oDoc, $oViewCursor, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to Set the Numbering Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Line 1" & @LF & "Line 1.1" & @LF & "Line 1.2" & @CR & _
			"Line 2" & @LF & "Line 2.1" & @LF & "Line 2.2" & @CR & "Line 3" & @LF & "Line 3.1" & @LF & "Line 3.2")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(.5)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 3/4" to Micrometers
	$iMicrometers2 = _LO_ConvertToMicrometer(.75)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1" to Micrometers
	$iMicrometers3 = _LO_ConvertToMicrometer(1)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Numbering Style position settings: Modify all Levels (0), Aligned at = 1/2", Numbering Align = $LOW_ORIENT_HORI_CENTER,
	; Followed by = $LOW_FOLLOW_BY_TABSTOP, TabStop = 3/4", Indent = 1"
	_LOWriter_NumStylePosition($oDoc, $oNumStyle, 0, $iMicrometers, $LOW_ORIENT_HORI_CENTER, $LOW_FOLLOW_BY_TABSTOP, $iMicrometers2, $iMicrometers3)
	If @error Then _ERROR($oDoc, "Failed to set Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Numbering Style settings for all levels. Return will be an array containing arrays with values in order of function parameters.
	$aavSettings = _LOWriter_NumStylePosition($oDoc, $oNumStyle, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($aavSettings) - 1
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Numbering style's current Position settings for level " & ($i + 1) & " are as follows: " & @CRLF & _
				"The First line indent is, in Micrometers: " & ($aavSettings[$i])[0] & @CRLF & _
				"The Numbering symbols are aligned to, (see UDF constants): " & ($aavSettings[$i])[1] & @CRLF & _
				"The Numbering Symbol is followed by, (see UDF Constants): " & ($aavSettings[$i])[2] & @CRLF & _
				"The Tab Stop size following the Numbering symbols is, in Micrometers: " & ($aavSettings[$i])[3] & @CRLF & _
				"The indent from the left page margin to the Numbering symbols is, in micrometers: " & ($aavSettings[$i])[4])
	Next

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
