#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oRow, $oCellRange

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Row 3, which is in position 2 because L.O. Rows are 0 based.
	$oRow = _LOCalc_RangeRowGetObjByPosition($oSheet, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve Row Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Row 3's Background color to Black.
	_LOCalc_CellBackColor($oRow, $LO_COLOR_BLACK)
	If @error Then _ERROR($oDoc, "Failed to set Row's Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now Retrieve Cell Range D3 to F7, and set the fourth down row's background color to black also, using the Cell Range." & @CRLF & _
			"Notice it doesn't matter that the Cell Range doesn't cover the entire Row, the whole Row is set to black.")

	; Retrieve Cell Range D3 to F7.
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "D3", "F7")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the fourth down Row using the Cell Range.
	$oRow = _LOCalc_RangeRowGetObjByPosition($oCellRange, 3)
	If @error Then _ERROR($oDoc, "Failed to retrieve Row Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Row 6's Background color to Black.
	_LOCalc_CellBackColor($oRow, $LO_COLOR_BLACK)
	If @error Then _ERROR($oDoc, "Failed to set Row's Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
