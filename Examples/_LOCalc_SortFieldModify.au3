#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $tSortField
	Local $aavData[5]
	Local $avRowData[2], $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values I want in Column A.
	$avRowData[0] = 1 ; A1
	$avRowData[1] = "c" ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = 3 ; A2
	$avRowData[1] = "a" ; B2
	$aavData[1] = $avRowData

	$avRowData[0] = 2 ; A3
	$avRowData[1] = "B" ; B3
	$aavData[2] = $avRowData

	$avRowData[0] = 5 ; A4
	$avRowData[1] = "b" ; B4
	$aavData[3] = $avRowData

	$avRowData[0] = 4 ; A5
	$avRowData[1] = 1 ; B5
	$aavData[4] = $avRowData

	; Retrieve Cell range A1 to B5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create the Sort Field.
	; Make this Sort Field apply to Row Column B (1, because Columns are 0 based internally in Libre Office Calc.)
	; Set Data type to ALPHANUMERIC, and Ascending sort to True.
	$tSortField = _LOCalc_SortFieldCreate(1, $LOC_SORT_DATA_TYPE_ALPHANUMERIC, True, True)
	If @error Then _ERROR($oDoc, "Failed to create a Sort Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Perform a Sort operation on Range A1 to B5.
	; Set SortColumns to False, meaning sort Rows Top to Bottom, Has Headers= False, Bind format = False
	_LOCalc_RangeSort($oDoc, $oCellRange, $tSortField, False, False, False)
	If @error Then _ERROR($oDoc, "Failed to perform Sort Operation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have performed a Sorting Operation for Range A1 to B5. I will now modify the sort field, and perform the sort operation again.")

	; Modify The Sort Field.
	; Make this Sort Field apply to Row Column A (0, because Columns are 0 based internally in Libre Office Calc.)
	; Set Data type to NUMERIC, and Ascending sort to False, Case Sensitive = False
	_LOCalc_SortFieldModify($tSortField, 0, $LOC_SORT_DATA_TYPE_NUMERIC, False, False)
	If @error Then _ERROR($oDoc, "Failed to create a Sort Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Perform a Sort operation on Range A1 to B5 again.
	; Set SortColumns to False, meaning sort Rows Top to Bottom, Has Headers= False, Bind format = False
	_LOCalc_RangeSort($oDoc, $oCellRange, $tSortField, False, False, False)
	If @error Then _ERROR($oDoc, "Failed to perform Sort Operation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_SortFieldModify($tSortField)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Sort Field's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Sort Field's current settings are as follows: " & @CRLF & _
			"Which Column or Row is this Sort Field being applied to?: " & $avSettings[0] & @CRLF & _
			"What data type is this Sort Field set to? (See UDF constants): " & $avSettings[1] & @CRLF & _
			"Will data be sorted in ascending order? True/False: " & $avSettings[2] & @CRLF & _
			"Is the sort Case Sensitive? True/False: " & $avSettings[3])

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
