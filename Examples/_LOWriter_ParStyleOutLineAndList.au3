#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oParStyle
	Local $avParStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the Default Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style." & @CR & "Next Line" & @CR & "Next Line" & @LF)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Paragraph Style object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Default Paragraph Style Outline and List settings to, Outline level $LOW_OUTLINE_LEVEL_1, Numbering style,"Numbering 123",
	; Include line in paragraph count? True, Line count value = 1.
	_LOWriter_ParStyleOutLineAndList($oDoc, $oParStyle, $LOW_OUTLINE_LEVEL_1, "Numbering 123", True, 1)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avParStyleSettings = _LOWriter_ParStyleOutLineAndList($oDoc, $oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Outline and List settings are as follows: " & @CRLF & _
			"What Outline level is this Paragraph style set to? (see UDF constants): " & $avParStyleSettings[0] & @CRLF & _
			"What Numbering style is used, If any?: " & $avParStyleSettings[1] & @CRLF & _
			"Is this paragraph included in line numbering? True/False: " & $avParStyleSettings[2] & @CRLF & _
			"Should line counting restart at this paragraph? If so, what number should it start at? (If not, value will be 0): " & $avParStyleSettings[3])

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
