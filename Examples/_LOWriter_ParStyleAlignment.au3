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
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style." & @LF & "Next Line" & @CR & "Next Line" & @LF)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Paragraph Style object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Default Paragraph Style Alignment settings to, Horizontal alignment = $LOW_PAR_ALIGN_HOR_JUSTIFIED,
	; Vertical alignment = $LOW_PAR_ALIGN_VERT_CENTER, Last line alignment = $LOW_PAR_LAST_LINE_JUSTIFIED,
	; Expand single word = True, Snap to grid = False, Text direction = $LOW_TXT_DIR_LR_TB
	_LOWriter_ParStyleAlignment($oParStyle, $LOW_PAR_ALIGN_HOR_JUSTIFIED, $LOW_PAR_ALIGN_VERT_CENTER, $LOW_PAR_LAST_LINE_JUSTIFIED, True, False, $LOW_TXT_DIR_LR_TB)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avParStyleSettings = _LOWriter_ParStyleAlignment($oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Alignment settings are as follows: " & @CRLF & _
			"Horizontal alignment, (See UDF constants): " & $avParStyleSettings[0] & @CRLF & _
			"Vertical alignment, (See UDF constants): " & $avParStyleSettings[1] & @CRLF & _
			"Last line alignment, (See UDF constants): " & $avParStyleSettings[2] & @CRLF & _
			"Expand a single word on the last line? True/False: " & $avParStyleSettings[3] & @CRLF & _
			"Snap to grid, if one is active? True/False: " & $avParStyleSettings[4] & @CRLF & _
			"Text Direction, (See UDF constants): " & $avParStyleSettings[5])

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
