#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oTextCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve active Sheet's Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell A1's Object
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Create a Text Cursor in the Cell
	$oTextCursor = _LOCalc_CellCreateTextCursor($oCell)
	If @error Then _ERROR($oDoc, "Failed to create Text Cursor in Cell. Error:" & @error & " Extended:" & @extended)

	; Insert some text
	_LOCalc_TextCursorInsertString($oTextCursor, "Hi! This is some text.")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended)

	; Move the Text Cursor left 5 spaces, selecting as I move.
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_LEFT, 5, True)
	If @error Then _ERROR($oDoc, "Failed to move Text cursor. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have selected the word ""Text"". You won't be able to see anything selected, but to show it is selected, I will use the cursor to set the text to Bold.")

	; Set the selected text to Bold.
	_LOCalc_TextCursorFont($oDoc, $oTextCursor, Null, Null, Null, $LOC_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to set font weight. Error:" & @error & " Extended:" & @extended)

	; Move the Text Cursor to the Start.
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move Text cursor. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have moved the cursor to the beginning of the text. I will insert a # sign at its location.")

	; Insert a # sign.
	_LOCalc_TextCursorInsertString($oTextCursor, "#")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended)

	; Move the Text Cursor to the right 4 spaces, to just behind the word "This".
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_RIGHT, 4, False)
	If @error Then _ERROR($oDoc, "Failed to move Text cursor. Error:" & @error & " Extended:" & @extended)

	; Move the Text Cursor to the right 4 spaces, selecting the word "This".
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_RIGHT, 4, True)
	If @error Then _ERROR($oDoc, "Failed to move Text cursor. Error:" & @error & " Extended:" & @extended)

	; Set the selected text to Bold.
	_LOCalc_TextCursorFont($oDoc, $oTextCursor, Null, Null, Null, $LOC_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to set font weight. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have moved the cursor to, and selected the word ""This"", setting it to Bold font weight. " & _
			"I will collapse the selection to the start, and insert a # sign at the Cursor location.")

	; Move the Text Cursor to the right 4 spaces, selecting the word "This".
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_COLLAPSE_TO_START)
	If @error Then _ERROR($oDoc, "Failed to move Text cursor. Error:" & @error & " Extended:" & @extended)

	; Insert a # sign.
	_LOCalc_TextCursorInsertString($oTextCursor, "#")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
