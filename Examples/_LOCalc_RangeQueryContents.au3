#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $aoRanges[0]
	Local $iResults

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_FillCells($oDoc, $oSheet)

	; Retrieve Cell range A1 to C6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "C6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Query the Cell Range for cells containing Strings
	$aoRanges = _LOCalc_RangeQueryContents($oCellRange, $LOC_CELL_FLAG_STRING)
	If @error Then _ERROR($oDoc, "Failed to Query Cell Range Contents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	$iResults = @extended

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now highlight in yellow the cell ranges returned that contain Strings.")

	; Cycle through the results and set the background color to yellow for each Cell range found
	For $i = 0 To $iResults - 1
		_LOCalc_CellBackColor($aoRanges[$i], $LO_COLOR_YELLOW)
		If @error Then _ERROR($oDoc, "Failed to set Cell Range Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Query the Cell Range for cells containing Numbers and Formulas
	$aoRanges = _LOCalc_RangeQueryContents($oCellRange, BitOR($LOC_CELL_FLAG_VALUE, $LOC_CELL_FLAG_FORMULA))
	If @error Then _ERROR($oDoc, "Failed to Query Cell Range Contents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	$iResults = @extended

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now highlight in Red the cell ranges returned that contain Numbers and Formulas.")

	; Cycle through the results and set the background color to Red for each Cell range found
	For $i = 0 To $iResults - 1
		_LOCalc_CellBackColor($aoRanges[$i], $LO_COLOR_RED)
		If @error Then _ERROR($oDoc, "Failed to set Cell Range Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _FillCells(ByRef $oDoc, ByRef $oSheet)
	Local $oCellRange, $oCell
	Local $aavData[6]
	Local $avRowData[1]

	; Fill my arrays with the desired Number and String Values I want in A Column.
	$avRowData[0] = 1 ; A1
	$aavData[0] = $avRowData

	$avRowData[0] = "String" ; A2
	$aavData[1] = $avRowData

	$avRowData[0] = 5 ; A3
	$aavData[2] = $avRowData

	$avRowData[0] = "Hi" ; A4
	$aavData[3] = $avRowData

	$avRowData[0] = 10 ; A5
	$aavData[4] = $avRowData

	$avRowData[0] = 0 ; A6
	$aavData[5] = $avRowData

	; Retrieve Cell range A1 to A6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with my Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A6
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell A6 to a Formula
	_LOCalc_CellFormula($oCell, "=A1 - A3")
	If @error Then _ERROR($oDoc, "Failed to set Cell Formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Prepare data to fill Row B with Data
	$avRowData[0] = "B1" ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = 457 ; B2
	$aavData[1] = $avRowData

	$avRowData[0] = "B3" ; B3
	$aavData[2] = $avRowData

	$avRowData[0] = 2300 ; B4
	$aavData[3] = $avRowData

	$avRowData[0] = 0 ; B5
	$aavData[4] = $avRowData

	$avRowData[0] = 5 ; B6
	$aavData[5] = $avRowData

	; Retrieve Cell range B1 to B6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B1", "B6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with the Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell B5
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B5 to a Formula
	_LOCalc_CellFormula($oCell, "=B2 - B4")
	If @error Then _ERROR($oDoc, "Failed to set Cell Formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values to fill Column C with.
	$avRowData[0] = "C1" ; C1
	$aavData[0] = $avRowData

	$avRowData[0] = 0 ; C2
	$aavData[1] = $avRowData

	$avRowData[0] = 1 ; C3
	$aavData[2] = $avRowData

	$avRowData[0] = 7 ; C4
	$aavData[3] = $avRowData

	$avRowData[0] = "C5" ; C5
	$aavData[4] = $avRowData

	$avRowData[0] = 6 ; C6
	$aavData[5] = $avRowData

	; Retrieve Cell range C1 to C6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "C1", "C6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
