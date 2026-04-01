#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oFilterRange, $oCell
	Local $aavData[6]
	Local $avRowData[4]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Values I want in Column A, B, C and D.
	$avRowData[0] = "Month" ; A1
	$avRowData[1] = "Income" ; B1
	$avRowData[2] = "Expenses" ; C1
	$avRowData[3] = "Balance" ; D1
	$aavData[0] = $avRowData

	$avRowData[0] = "January" ; A2
	$avRowData[1] = 10213.00 ; B2
	$avRowData[2] = 8711.53 ; C2
	$avRowData[3] = 1501.47 ; D2
	$aavData[1] = $avRowData

	$avRowData[0] = "February" ; A3
	$avRowData[1] = 7544.10 ; B3
	$avRowData[2] = 2358.68 ; C3
	$avRowData[3] = 6686.89 ; D3
	$aavData[2] = $avRowData

	$avRowData[0] = "March" ; A4
	$avRowData[1] = 9432.89 ; B4
	$avRowData[2] = 15294.36 ; C4
	$avRowData[3] = 825.42 ; D4
	$aavData[3] = $avRowData

	$avRowData[0] = "April" ; A5
	$avRowData[1] = 6588.25 ; B5
	$avRowData[2] = 3687.23 ; C5
	$avRowData[3] = 3726.44 ; D5
	$aavData[4] = $avRowData

	$avRowData[0] = "May"     ; A6
	$avRowData[1] = 2198.29 ; B6
	$avRowData[2] = 5679.25 ; C6
	$avRowData[3] = 245.48 ; D6
	$aavData[5] = $avRowData

	; Retrieve Cell range A1 to D6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "D6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	ReDim $aavData[1]

	; Prepare my filter Criteria.
	; Fill my arrays with the same column headers found in Columns A, B, C and D.
	$avRowData[0] = "Month"
	$avRowData[1] = "Income"
	$avRowData[2] = "Expenses"
	$avRowData[3] = "Balance"
	$aavData[0] = $avRowData

	; Retrieve Cell range C8 to F8
	$oFilterRange = _LOCalc_RangeGetCellByName($oSheet, "C8", "F8")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oFilterRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range C8 to F10
	$oFilterRange = _LOCalc_RangeGetCellByName($oSheet, "C8", "F10")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Get the Cell "F9"
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "F9")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the cell's string
	_LOCalc_CellString($oCell, "<1000") ; Look for Values less than 1000, under the "Balance" column.
	If @error Then _ERROR($oDoc, "Failed to set Cell String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Get the Cell D10
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "D10")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the cell's string
	_LOCalc_CellString($oCell, ">9000") ; Also look for values greater than 9,000 under the "Income" column.
	If @error Then _ERROR($oDoc, "Failed to set Cell String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Get the Cell E10
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "E10")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the cell's string
	_LOCalc_CellString($oCell, "<10000") ; But only leave values that have "Expenses" less than 10,000
	If @error Then _ERROR($oDoc, "Failed to set Cell String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to perform the Filtering Operation.")

	; Filter the Range with my Filter Criteria range.
	_LOCalc_RangeFilterAdvanced($oCellRange, $oFilterRange)
	If @error Then _ERROR($oDoc, "Failed to Filter Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
