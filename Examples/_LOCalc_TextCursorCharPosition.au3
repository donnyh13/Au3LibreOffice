#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oTextCursor
	Local $avSettings[0]

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

	; Insert a Word
	_LOCalc_TextCursorInsertString($oTextCursor, "Hi! Testing.")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now set the word ""Hi!"" to Auto Superscript, 65 percent relative size.")

	; Go to the Start.
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GOTO_START, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Select the word Hi!
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_RIGHT, 3, True)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Set the Word Hi to Auto Superscript, and 65% Relative size.
	_LOCalc_TextCursorCharPosition($oTextCursor, True, Null, Null, Null, 65)
	If @error Then _ERROR($oDoc, "Failed to set text Position. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now select the word ""Testing"" and set it to Subscript 75%, and 85% relative size.")

	; Move the cursor to the right once
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_RIGHT, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Move the cursor to the right 7 times to select the word Testing.
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_RIGHT, 7, True)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Set the Word Testing to 75% Subscript, and 85% Relative size.
	_LOCalc_TextCursorCharPosition($oTextCursor, Null, Null, Null, 75, 85)
	If @error Then _ERROR($oDoc, "Failed to set text Position. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Position settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_TextCursorCharPosition($oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve current format settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Character Position settings at the Cursor's current position are as follows: " & @CRLF & _
			"Is auto Superscript currently active? True/False: " & $avSettings[0] & @CRLF & _
			"The Superscript percentage is: " & $avSettings[1] & @CRLF & _
			"Is auto Subscript currently active? True/False: " & $avSettings[2] & @CRLF & _
			"The Subscript percentage is: " & $avSettings[3] & @CRLF & _
			"The relative percentage is: " & $avSettings[4])

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
