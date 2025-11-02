#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oSheetCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to A5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will create a Sheet Cursor for the Range A1 to A5, then I will set the background Color to Black to show what area that Cursor is covering.")

	; Create a Sheet Cursor for the Range.
	$oSheetCursor = _LOCalc_RangeCreateCursor($oSheet, $oCellRange)
	If @error Then _ERROR($oDoc, "Failed to create a Sheet Cursor for the Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set background color to black.
	_LOCalc_CellBackColor($oSheetCursor, $LO_COLOR_BLACK)
	If @error Then _ERROR($oDoc, "Failed to set background color for the Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now move the cursor right 5 Spaces, then set the background color to Blue.")

	; Move the Cursor 5 Spaces Next(Right)
	_LOCalc_SheetCursorMove($oSheetCursor, $LOC_SHEETCUR_GOTO_NEXT, 0, 0, 5)
	If @error Then _ERROR($oDoc, "Failed to perform a Cursor move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set background color to blue.
	_LOCalc_CellBackColor($oSheetCursor, $LO_COLOR_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set background color for the Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now move the cursor right 1 Space, then expand the selection to cover 3 columns, and 6 Rows, then merge the Range.")

	; Move the Cursor 1 Spaces Next(Right)
	_LOCalc_SheetCursorMove($oSheetCursor, $LOC_SHEETCUR_GOTO_NEXT, 0, 0, 1)
	If @error Then _ERROR($oDoc, "Failed to perform a Cursor move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Expand the Cursor selection to 3 Columns Right and 6 Rows down.
	_LOCalc_SheetCursorMove($oSheetCursor, $LOC_SHEETCUR_COLLAPSE_TO_SIZE, 3, 6)
	If @error Then _ERROR($oDoc, "Failed to perform a Cursor move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now merge the Range covered by the cursor.")

	; Merge the Range covered by the cursor.
	_LOCalc_RangeMerge($oSheetCursor, True)
	If @error Then _ERROR($oDoc, "Failed to merge the Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
