#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $aavData[1]
	Local $avRowData[6]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values I want in Column A.
	$avRowData[0] = 1 ; A1
	$avRowData[1] = 2 ; B1
	$avRowData[2] = 3 ; C1
	$avRowData[3] = 4 ; D1
	$avRowData[4] = 5 ; E1
	$avRowData[5] = 6 ; F1
	$aavData[0] = $avRowData

	; Retrieve Cell range A1 to F1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "F1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeNumbers($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell Range B1 to E1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B1", "E1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Group the Columns of B to E
	_LOCalc_RangeGroup($oCellRange, $LOC_GROUP_ORIENT_COLUMNS, True)
	If @error Then _ERROR($oDoc, "Failed to group Cell range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell Range C1 to D1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "C1", "D1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Group the Columns of C to D
	_LOCalc_RangeGroup($oCellRange, $LOC_GROUP_ORIENT_COLUMNS, True)
	If @error Then _ERROR($oDoc, "Failed to group Cell range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to hide Cell range C1 to D1 group.")

	; Hide Cell Range C1 to D1
	_LOCalc_RangeDetail($oCellRange, False)
	If @error Then _ERROR($oDoc, "Failed to hide Cell range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
