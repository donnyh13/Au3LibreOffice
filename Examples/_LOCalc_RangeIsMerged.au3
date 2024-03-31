#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell range B1 to B5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B1", "C5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Merge cell range B1:B5
	_LOCalc_RangeMerge($oCellRange, True)
	If @error Then _ERROR($oDoc, "Failed to merge Cell Range. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Cell Range B1:B5 merged? True/False: " & _LOCalc_RangeIsMerged($oCellRange))

	; Retrieve Cell range A3 to C3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A3", "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Cell Range A3:C3 merged? True/False: " & _LOCalc_RangeIsMerged($oCellRange))

	; Retrieve Cell range B2 to B4
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B2", "B4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Cell Range B2:B4 merged? True/False: " & _LOCalc_RangeIsMerged($oCellRange))

	; Retrieve Cell range B1 to B4
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B1", "B4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Cell Range B1:B4 merged? True/False: " & _LOCalc_RangeIsMerged($oCellRange))

	; Retrieve Cell B2
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Cell B2 merged? True/False: " & _LOCalc_RangeIsMerged($oCellRange))

	; Retrieve Cell B1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Cell B1 merged? True/False: " & _LOCalc_RangeIsMerged($oCellRange))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
