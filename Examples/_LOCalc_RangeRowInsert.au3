#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oCellRange
	Local $iCount = 0

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the Cell Range of A1 to C5 with numbers, one cell at a time. (Remember Columns and Rows are 0 based.)
	For $i = 0 To 2
		For $j = 0 To 4
			; Retrieve the Cell Object
			$oCell = _LOCalc_RangeGetCellByPosition($oSheet, $i, $j)
			If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			; Set the Cell to a Number
			_LOCalc_CellValue($oCell, $iCount)
			If @error Then _ERROR($oDoc, "Failed to set Cell Value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			$iCount += 1
		Next
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now Insert 2 new Rows starting at Row 2.")

	; Insert new Rows at Row 2, Row 2 is counted as Row 1 because L.O. Rows are 0 based.
	_LOCalc_RangeRowInsert($oSheet, 1, 2)
	If @error Then _ERROR($oDoc, "Failed to delete rows. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now Retrieve Cell Range B4 to D6, and insert a new Row in the second down Row position.")

	; Retrieve Cell Range B4 to D6.
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B4", "D6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert new Rows at the second down Row position.
	_LOCalc_RangeRowInsert($oCellRange, 1)
	If @error Then _ERROR($oDoc, "Failed to delete rows. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
