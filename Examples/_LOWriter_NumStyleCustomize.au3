#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oNumStyle, $oViewCursor
	Local $avSettings

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
	_LOWriter_NumStyleCurrent($oDoc, $oViewCursor, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to Set the Numbering Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Line 1" & @LF & "Line 1.1" & @LF & "Line 1.2" & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Numbering Style Level for This Paragraph to 2.
	_LOWriter_NumStyleSetLevel($oViewCursor, 2)
	If @error Then _ERROR($oDoc, "Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Level 2" & @LF & "Line 2.1" & @LF & "Line 2.2" & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Numbering Style Level for This Paragraph to 3.
	_LOWriter_NumStyleSetLevel($oViewCursor, 3)
	If @error Then _ERROR($oDoc, "Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Line 3" & @LF & "Line 3.1" & @LF & "Line 3.2")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Numbering Style Customization settings: Modify Level 2, Numbering format = $LOW_NUM_STYLE_ARABIC, Start at 4, Char Style = "Emphasis",
	; Sub levels = 2, Separator before = ~ , Separator after = #, Consecutive Numbering = False.
	_LOWriter_NumStyleCustomize($oDoc, $oNumStyle, 2, $LOW_NUM_STYLE_ARABIC, 4, "Emphasis", 2, "~", "#", False)
	If @error Then _ERROR($oDoc, "Failed to set Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Numbering Style settings for level 2. Return will be an array in order of function parameters, Return will only have
	; seven elements, because Numbering Style is not set to special, so there will not be a Bullet or Char Decimal value.
	$avSettings = _LOWriter_NumStyleCustomize($oDoc, $oNumStyle, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Numbering style's current Customization settings for level 2 are as follows: " & @CRLF & _
			"The Number format used is, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The Numbering starts at: " & $avSettings[1] & @CRLF & _
			"The Character Style to use is: " & $avSettings[2] & @CRLF & _
			"The number of sub levels to include is: " & $avSettings[3] & @CRLF & _
			"The Separator before the Numbering symbol is: " & $avSettings[4] & @CRLF & _
			"The Separator After the Numbering symbol is: " & $avSettings[5] & @CRLF & _
			"Consecutively number levels? True/False: " & $avSettings[6])

	; Modify the Numbering Style Customization settings: Modify all levels, Numbering format = $LOW_NUM_STYLE_ROMAN_UPPER, Consecutive Numbering = True.
	_LOWriter_NumStyleCustomize($oDoc, $oNumStyle, 0, $LOW_NUM_STYLE_ROMAN_UPPER, Null, Null, Null, Null, Null, True)
	If @error Then _ERROR($oDoc, "Failed to set Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Numbering Style settings for all levels. Return will be an array in order of function parameters, Return will only have
	; seven elements, because Numbering Style is not set to special, so there will not be a Bullet or Char Decimal value.
	$avSettings = _LOWriter_NumStyleCustomize($oDoc, $oNumStyle, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($avSettings) - 1
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Numbering style's current Customization settings for level " & ($i + 1) & " are as follows: " & @CRLF & _
				"The Number format used is, (see UDF constants): " & ($avSettings[$i])[0] & @CRLF & _
				"The Numbering starts at: " & ($avSettings[$i])[1] & @CRLF & _
				"The Character Style to use is: " & ($avSettings[$i])[2] & @CRLF & _
				"The number of sub levels to include is: " & ($avSettings[$i])[3] & @CRLF & _
				"The Separator before the Numbering symbol is: " & ($avSettings[$i])[4] & @CRLF & _
				"The Separator After the Numbering symbol is: " & ($avSettings[$i])[5] & @CRLF & _
				"Consecutively number levels? True/False: " & ($avSettings[$i])[6])
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
