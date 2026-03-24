#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2, $oSheet, $oSheet2, $oCellRange
	Local $aavData[3]
	Local $avRowData[2]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Sheet named "New Sheet".
	$oSheet = _LOCalc_SheetAdd($oDoc, "New Sheet")
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to create a new Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Activate the new Sheet
	_LOCalc_SheetActivate($oDoc, $oSheet)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to activate the Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number and String Values I want in Column A to B.
	$avRowData[0] = "Seventy" ; A1
	$avRowData[1] = 5 ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = 10 ; A2
	$avRowData[1] = 20 ; B2
	$aavData[1] = $avRowData

	$avRowData[0] = "Testing" ; A3
	$avRowData[1] = -1700 ; B3

	$aavData[2] = $avRowData

	; Retrieve Cell range A1 to B3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B3")
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now open a new document and import this Sheet into it.")

	; Create a New, visible, Blank Libre Office Document.
	$oDoc2 = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Import the Sheet "New Sheet" from Document 1 into this Document, I'll use a function to retrieve the Sheet's name, but I could also just call the Sheet name as "New Sheet"
	$oSheet2 = _LOCalc_SheetImport($oDoc, $oDoc2, _LOCalc_SheetName($oDoc, $oSheet))
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to import Calc Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Activate the new Sheet
	_LOCalc_SheetActivate($oDoc2, $oSheet2)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to activate Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have imported the Sheet from Document 1 into Document 2, the new sheet is called: " & _LOCalc_SheetName($oDoc2, $oSheet2))

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOCalc_DocClose($oDoc2, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, $oDoc2, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $oDoc2, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	If IsObj($oDoc2) Then _LOCalc_DocClose($oDoc2, False)
	Exit
EndFunc
