#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to A5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now merge range A1:A5.")

	; Merge the Range A1:A5.
	_LOCalc_RangeMerge($oCellRange, True)
	If @error Then _ERROR($oDoc, "Failed to merge the Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range C2 to C4
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "C2", "C4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now merge range C2:C4.")

	; Merge the Range C2:C4.
	_LOCalc_RangeMerge($oCellRange, True)
	If @error Then _ERROR($oDoc, "Failed to merge the Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now un-merge range A4:C4. Notice nothing happens.")

	; Retrieve Cell range A4 to C4
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A4", "C4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Un-Merge the Range A4:C4.
	_LOCalc_RangeMerge($oCellRange, False)
	If @error Then _ERROR($oDoc, "Failed to merge the Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now un-merge range A1:C1. Notice the merged range of A1:A5 will become un-merged.")

	; Retrieve Cell range A1 to C1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "C1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Un-Merge the Range A1:C1.
	_LOCalc_RangeMerge($oCellRange, False)
	If @error Then _ERROR($oDoc, "Failed to merge the Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
