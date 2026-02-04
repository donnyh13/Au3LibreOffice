#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oTextCursor
	Local $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve active Sheet's Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A1's Object
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the Cell
	$oTextCursor = _LOCalc_CellCreateTextCursor($oCell)
	If @error Then _ERROR($oDoc, "Failed to create Text Cursor in Cell. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some words
	_LOCalc_TextCursorInsertString($oTextCursor, "Hi! Testing.")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now select all the words, and set Overline only words to True, Overline style, Dash Dot, and Overline in the color Gold.")

	; Go to the Start.
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GOTO_START, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select all the words.
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GOTO_END, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Overline settings to, Underline words only, Underline style $LOC_UNDERLINE_BOLD_DASH_DOT, has color = True, and underline color = Gold.
	_LOCalc_TextCursorOverline($oTextCursor, True, $LOC_UNDERLINE_BOLD_DASH_DOT, True, $LO_COLOR_GOLD)
	If @error Then _ERROR($oDoc, "Failed to set text formatting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_TextCursorOverline($oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve current format settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Overline settings at the Cursor's current position are as follows: " & @CRLF & _
			"Are words only Overlined? True/False: " & $avSettings[0] & @CRLF & _
			"The Overline style is (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"Is the Overline colored? True/False: " & $avSettings[2] & @CRLF & _
			"The Overline color is (as a RGB Color Integer): " & $avSettings[3])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
