#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oRange, $oSelection, $oCell
	Local $aavData[3]
	Local $avRowData[1]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Values.
	$avRowData[0] = "Heading 1" ; A1
	$aavData[0] = $avRowData

	$avRowData[0] = 1 ; A2
	$aavData[1] = $avRowData

	$avRowData[0] = 2 ; A3
	$aavData[2] = $avRowData

	; Retrieve Cell range A1 to A3
	$oRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now select the Range A1:A3 and copy it, then paste it in to the document beginning at Cell C4.")

	; Select the Range.
	_LOCalc_DocSelectionSet($oDoc, $oRange)
	If @error Then _ERROR($oDoc, "Failed to set Selection. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Copy the data selected.
	$oSelection = _LOCalc_DocSelectionCopy($oDoc)
	If @error Then _ERROR($oDoc, "Failed to copy Selection. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell C4 Object.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select the Cell.
	_LOCalc_DocSelectionSet($oDoc, $oCell)
	If @error Then _ERROR($oDoc, "Failed to set Selection. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Paste the data in Range C4.
	_LOCalc_DocSelectionPaste($oDoc, $oSelection)
	If @error Then _ERROR($oDoc, "Failed to paste Selection. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
