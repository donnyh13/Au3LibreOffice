#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCell, $oFilterDesc
	Local $aavData[5]
	Local $avRowData[2]
	Local $atFilterFields[2]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired String Values I want in Column A and B.
	$avRowData[0] = "First" ; A1
	$avRowData[1] = "has a string" ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = "A String with a Number 1" ; A2
	$avRowData[1] = "Hello" ; B2
	$aavData[1] = $avRowData

	$avRowData[0] = "Column A" ; A3
	$avRowData[1] = "Testing" ; B3
	$aavData[2] = $avRowData

	$avRowData[0] = 123 ; A4
	$avRowData[1] = "fourth string" ; B4
	$aavData[3] = $avRowData

	$avRowData[0] = "Last"     ; A5
	$avRowData[1] = 75 ; B5
	$aavData[4] = $avRowData

	; Retrieve Cell range A1 to B5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell D5
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "D5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create my first Filter Field, I will insert it directly into my Array.
	; Make this Filter Field apply to Column A (0, because Columns are 0 based internally in Libre Office Calc.)
	; Set Numeric to False, and my value to 0 to skip it, String = \d (any Number Value), Condition to "Contains" my value. (I don't need to worry about Operator, because this is the first Field in my Array.)
	$atFilterFields[0] = _LOCalc_FilterFieldCreate(0, False, 0, "\d", $LOC_FILTER_CONDITION_CONTAINS)
	If @error Then _ERROR($oDoc, "Failed to create a Filter Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create my second Filter Field.
	; Make this Filter Field apply to Column B (1)
	; Set Numeric to False, and my value to 0 to skip it, Set String to "[A-Z]" to obtain any Captitals, Condition to "Contain" my value (0).
	; Set Operator to OR, because I want to find either Cells containing Digits in Column A, or Capitals in Column B.
	$atFilterFields[1] = _LOCalc_FilterFieldCreate(1, True, 0, "", $LOC_FILTER_CONDITION_CONTAINS, $LOC_FILTER_OPERATOR_OR)
	If @error Then _ERROR($oDoc, "Failed to create a Filter Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Filter Descriptor.
	; Use my Filter Fields Array I just created, Set Case Sensitive to True, Skip Duplicates to False, Use Regular Expressions = True, Headers = False,
	; Copy Output = True, and set Output to Cell D5
	$oFilterDesc = _LOCalc_FilterDescriptorCreate($oCellRange, $atFilterFields, True, False, True, False, True, $oCell)
	If @error Then _ERROR($oDoc, "Failed to create a Filter Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to filter the range.")

	; Perform a Filter operation on Range A1 to B5.
	_LOCalc_RangeFilter($oCellRange, $oFilterDesc)
	If @error Then _ERROR($oDoc, "Failed to perform Filter Operation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
