#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $tSortField
	Local $aavData[1]
	Local $avRowData[6]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values I want in Row 1.
	$avRowData[0] = "A11" ; A1
	$avRowData[1] = "A1" ; B1
	$avRowData[2] = "B1" ; C1
	$avRowData[3] = "B2" ; D1
	$avRowData[4] = "B23" ; E1
	$avRowData[5] = "B3" ; F1
	$aavData[0] = $avRowData

	; Retrieve Cell range A1 to F1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "F1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create The Sort Field.
	; Make this Sort Field apply to Row 1 (0, because Rows are 0 based internally in Libre Office Calc.)
	; Set Data type to Auto, and Ascending sort to True.
	$tSortField = _LOCalc_SortFieldCreate(0, $LOC_SORT_DATA_TYPE_AUTO, True)
	If @error Then _ERROR($oDoc, "Failed to create a Sort Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to perform the Sorting Operation.")

	; Perform a Sort operation on Range A1 to F1.
	; Set SortColumns to True, meaning sort Columns Left tot Right, Has Headers= False, Bind format = False, Natural Order = True
	_LOCalc_RangeSortAlt($oDoc, $oCellRange, $tSortField, True, False, False, True)
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
