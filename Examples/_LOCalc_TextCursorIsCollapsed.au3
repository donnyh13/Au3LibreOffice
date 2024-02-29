#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oTextCursor
	Local $bReturn

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
	_LOCalc_TextCursorInsertString($oTextCursor, "Hi! This is a short string of text.")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended)

	; Check if the Cursor's selection is currently collapsed, (no data is selected).
	$bReturn = _LOCalc_TextCursorIsCollapsed($oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to check cursor status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the Text Cursor's selection currently empty? True/False: " & $bReturn)

	; Move the Text Cursor left 3 spaces, selecting as I move.
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_LEFT, 3, True)
	If @error Then _ERROR($oDoc, "Failed to move Text cursor. Error:" & @error & " Extended:" & @extended)

	; Check if the Cursor's selection is currently collapsed after the move.
	$bReturn = _LOCalc_TextCursorIsCollapsed($oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to check cursor status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Now is the Text Cursor's selection currently empty? True/False: " & $bReturn)

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
