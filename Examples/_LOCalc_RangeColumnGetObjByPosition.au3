#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oColumn, $oCellRange

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the column D, which is in position 3 because L.O. Columns are 0 based.
	$oColumn = _LOCalc_RangeColumnGetObjByPosition($oSheet, 3)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Column D's Background color to Black.
	_LOCalc_CellBackColor($oColumn, $LO_COLOR_BLACK)
	If @error Then _ERROR($oDoc, "Failed to set Column's Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now Retrieve Cell Range D3 to F7, and set the third over column's background color to black also, using the Cell Range." & @CRLF & _
			"Notice it doesn't matter that the Cell Range doesn't cover the entire column, the whole Column is set to black.")

	; Retrieve Cell Range D3 to F7.
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "D3", "F7")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the third over column.
	$oColumn = _LOCalc_RangeColumnGetObjByPosition($oCellRange, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Column F's Background color to Black.
	_LOCalc_CellBackColor($oColumn, $LO_COLOR_BLACK)
	If @error Then _ERROR($oDoc, "Failed to set Column's Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
