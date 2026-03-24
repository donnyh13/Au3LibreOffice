#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oTextCursor
	Local $aoPar[0], $aoPortions[0][2]

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
	_LOCalc_TextCursorInsertString($oTextCursor, "Hi! Just Testing.")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select the word "Testing".
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_LEFT, 8, True)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the font weight to Bold.
	_LOCalc_TextCursorFont($oTextCursor, Null, Null, Null, $LOC_CHAR_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to set text formatting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Go to the Start.
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GOTO_START, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select the word "Hi!".
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_RIGHT, 3, True)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the font Posture to Italic to Bold.
	_LOCalc_TextCursorFont($oTextCursor, Null, Null, $LOC_CHAR_POSTURE_ITALIC)
	If @error Then _ERROR($oDoc, "Failed to set text formatting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Paragraph objects
	$aoPar = _LOCalc_TextCursorParObjCreateList($oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of paragraph Objects. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Text Portions for the first paragraph. There will be two, because there is different formatting than the rest of the cell.
	$aoPortions = _LOCalc_TextCursorParObjSectionsGet($aoPar[0])
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Text Portion Objects. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I retrieved an Array of Text portion Object for the first Paragraph of Cell A1" & @CRLF & _
			"I will now move the Text cursor to, and select, the middle word of the paragraph which hasn't been formatted yet using the Paragraph Text portion Object." & @CRLF & _
			"Then I will set the Font Style to Stencil.")

	; Move the Text Cursor to the middle word, which should be the second element of the Text portion array.
	_LOCalc_TextCursorGoToRange($oTextCursor, $aoPortions[1][0])
	If @error Then _ERROR($oDoc, "Failed to move cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the selected word to use the font Stencil.
	_LOCalc_TextCursorFont($oTextCursor, "Stencil")
	If @error Then _ERROR($oDoc, "Failed to set text formatting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
