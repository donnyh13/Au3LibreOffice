#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCell
	Local $tSortField
	Local $aavData[5]
	Local $avRowData[1]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values I want in Column A.
	$avRowData[0] = 1 ; A1
	$aavData[0] = $avRowData

	$avRowData[0] = 3 ; A2
	$aavData[1] = $avRowData

	$avRowData[0] = 2 ; A3
	$aavData[2] = $avRowData

	$avRowData[0] = 5 ; A4
	$aavData[3] = $avRowData

	$avRowData[0] = 4 ; A5
	$aavData[4] = $avRowData

	; Retrieve Cell range A1 to A5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell C5
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create The Sort Field.
	; Make this Sort Field apply to Column A (0, because Columns are 0 based internally in Libre Office Calc.)
	; Set Data type to Numeric, and Ascending sort to False (Descending).
	$tSortField = _LOCalc_SortFieldCreate(0, $LOC_SORT_DATA_TYPE_NUMERIC, False)
	If @error Then _ERROR($oDoc, "Failed to create a Sort Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to perform the Sorting Operation.")

	; Perform a Sort operation on Range A1 to A5.
	; Set SortColumns to False, meaning sort Rows, Has Headers= False, Bind format = False, Copy output = True, Copy output to Cell C5.
	_LOCalc_RangeSort($oDoc, $oCellRange, $tSortField, False, False, False, True, $oCell)
	If @error Then _ERROR($oDoc, "Failed to perform Sort Operation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
