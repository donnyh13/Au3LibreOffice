#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oDestination, $oPivot, $oField
	Local $atFilterFields[1]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_PrepareRange($oDoc, $oSheet)

	; Retrieve the range containing the Data.
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "F13")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Cell where I will output the Pivot Table to.
	$oDestination = _LOCalc_RangeGetCellByName($oSheet, "B15")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert the Pivot Table.
	_LOCalc_RangePivotInsert($oCellRange, $oDestination, "AutoIt_Pivot", "Item", $LOC_PIVOT_TBL_FIELD_TYPE_COLUMN)
	If @error Then _ERROR($oDoc, "Failed to insert Pivot Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Pivot Table Object.
	$oPivot = _LOCalc_RangePivotGetObjByName($oSheet, "AutoIt_Pivot")
	If @error Then _ERROR($oDoc, "Failed to retrieve Pivot Table Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Object for Field "Agent".
	$oField = _LOCalc_RangePivotFieldGetObjByName($oPivot, "Agent")
	If @error Then _ERROR($oDoc, "Failed to retrieve Pivot Table Field object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Field/Column "Agent" as a Row Field.
	_LOCalc_RangePivotFieldSettings($oField, $LOC_PIVOT_TBL_FIELD_TYPE_ROW)
	If @error Then _ERROR($oDoc, "Failed to set Pivot Table Field settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Object for Field "2012".
	$oField = _LOCalc_RangePivotFieldGetObjByName($oPivot, "2012")
	If @error Then _ERROR($oDoc, "Failed to retrieve Pivot Table Field object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Field/Column "2012" as a Data Field. set Subtotal to Max, don't include empty values and display the difference between "Item" and "Books".
	_LOCalc_RangePivotFieldSettings($oField, $LOC_PIVOT_TBL_FIELD_TYPE_DATA, $LOC_COMPUTE_FUNC_SUM, False)
	If @error Then _ERROR($oDoc, "Failed to set Pivot Table Field settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create my first Filter Field, I will insert it directly into my Array.
	; Make this Filter Field apply to the fourth column (3, because Columns are 0 based internally in Libre Office Calc.)
	; Set Numeric to True, and my value to 15,000, Skip String, set Condition to "Greater" than my value (15,000).
	$atFilterFields[0] = _LOCalc_FilterFieldCreate(3, True, 15000, "", $LOC_FILTER_CONDITION_GREATER)
	If @error Then _ERROR($oDoc, "Failed to create a Filter Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "press ok to filter the Pivot Table for only values greater than 15,000 in the 2012 column.")

	; Filter the Pivot Table's column 3, "2012", Skip Case sensitive, Set Skip duplicates to True, and Regular Expressions to False.
	_LOCalc_RangePivotFilter($oPivot, $atFilterFields, Null, True, False)
	If @error Then _ERROR($oDoc, "Failed to apply a Filter to the Pivot Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _PrepareRange($oDoc, $oSheet)
	Local $oCellRange
	Local $iFormatKey

	Local $avRow1[6] = ["Item", "Province", "Agent", 2012, 2013, 2014]
	Local $avRow2[6] = ["Books", "B.C.", "Michael", 17899.00, 21522.00, 10215.00]
	Local $avRow3[6] = ["Pens", "B.C.", "Michael", 13253.00, 18547.00, 16545.00]
	Local $avRow4[6] = ["Tape", "B.C.", "Michael", 23435.00, 17446.00, 25879.00]
	Local $avRow5[6] = ["Books", "Manitoba", "Don", 35669.00, 9855.00, 13874.00]
	Local $avRow6[6] = ["Pens", "Manitoba", "Don", 5488.00, 9487.00, 16598.00]
	Local $avRow7[6] = ["Tape", "Manitoba", "Don", 16899.00, 15874.00, 12845.00]
	Local $avRow8[6] = ["Books", "Alberta", "Nik", 18966.00, 8755.00, 14533.00]
	Local $avRow9[6] = ["Pens", "Alberta", "Nik", 13578.00, 9844.00, 17855.00]
	Local $avRow10[6] = ["Tape", "Alberta", "Nik", 10258.00, 6554.00, 16941.00]
	Local $avRow11[6] = ["Books", "P.E.I.", "Bohdan", 22469.00, 9985.00, 15897.00]
	Local $avRow12[6] = ["Pens", "P.E.I.", "Bohdan", 15885.00, 27488.00, 9885.00]
	Local $avRow13[6] = ["Tape", "P.E.I.", "Bohdan", 16987.00, 32369.00, 10255.00]
	Local $aavData[13] = [$avRow1, $avRow2, $avRow3, $avRow4, $avRow5, $avRow6, $avRow7, $avRow8, $avRow9, $avRow10, $avRow11, $avRow12, $avRow13]

	; Retrieve Cell range A1 to F13
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "F13")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range D2 to F13
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "D2", "F13")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Query the standard Format Key for the Format Key type of $LOC_FORMAT_KEYS_CURRENCY
	$iFormatKey = _LOCalc_FormatKeyGetStandard($oDoc, $LOC_FORMAT_KEYS_CURRENCY)
	If @error Then _ERROR($oDoc, "Failed to retrieve the standard format key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell Range's Numbering Format to Currency.
	_LOCalc_CellNumberFormat($oDoc, $oCellRange, $iFormatKey)
	If @error Then _ERROR($oDoc, "Failed to set Cell Numbering Format. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to F1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "F1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Bold the Titles.
	_LOCalc_CellFont($oCellRange, Null, Null, Null, $LOC_CHAR_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to set Cell Range weight. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
