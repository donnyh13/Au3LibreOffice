#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oNewDoc, $oSheet, $oNewSheet, $oRange, $oSelection, $oCell
	Local $aavData[3]
	Local $avRowData[1]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Fill my arrays with the desired Values.
	$avRowData[0] = "Heading 1" ; A1
	$aavData[0] = $avRowData

	$avRowData[0] = 1 ; A2
	$aavData[1] = $avRowData

	$avRowData[0] = 2 ; A3
	$aavData[2] = $avRowData

	; Retrieve Cell range A1 to A3
	$oRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Fill the range with Data
	_LOCalc_RangeData($oRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now select the Range A1:A3 and copy it, then paste it in to the document beginning at Cell C4.")

	; Select the Range.
	_LOCalc_DocSelectionSet($oDoc, $oRange)
	If @error Then _ERROR($oDoc, "Failed to set Selection. Error:" & @error & " Extended:" & @extended)

	; Copy the data selected.
	$oSelection = _LOCalc_DocSelectionCopy($oDoc)
	If @error Then _ERROR($oDoc, "Failed to copy Selection. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell C4 Object.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Select the Cell.
	_LOCalc_DocSelectionSet($oDoc, $oCell)
	If @error Then _ERROR($oDoc, "Failed to set Selection. Error:" & @error & " Extended:" & @extended)

	; Paste the data in Range C4.
	_LOCalc_DocSelectionPaste($oDoc, $oSelection)
	If @error Then _ERROR($oDoc, "Failed to paste Selection. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I can even paste it into another sheet.")

	; Create a new Sheet
	$oNewSheet = _LOCalc_SheetAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to create a new Sheet. Error:" & @error & " Extended:" & @extended)

	; Set the new sheet to active.
	_LOCalc_SheetActivate($oDoc, $oNewSheet)
	If @error Then _ERROR($oDoc, "Failed to activate a new Sheet. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Cell B1 in the new sheet.
	$oCell = _LOCalc_RangeGetCellByName($oNewSheet, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Select the Cell.
	_LOCalc_DocSelectionSet($oDoc, $oCell)
	If @error Then _ERROR($oDoc, "Failed to set Selection. Error:" & @error & " Extended:" & @extended)

	; Paste the data in Range B1.
	_LOCalc_DocSelectionPaste($oDoc, $oSelection)
	If @error Then _ERROR($oDoc, "Failed to paste Selection. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "And I can also paste the data into a new document.")

	; Create a New, visible, Blank Libre Office Document.
	$oNewDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended, $oNewDoc)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oNewDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended, $oNewDoc)

	; Retrieve the Cell D5 in the new Document.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "D5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended, $oNewDoc)

	; Select the Cell.
	_LOCalc_DocSelectionSet($oNewDoc, $oCell)
	If @error Then _ERROR($oDoc, "Failed to set Selection. Error:" & @error & " Extended:" & @extended, $oNewDoc)

	; Paste the data in Range D5.
	_LOCalc_DocSelectionPaste($oNewDoc, $oSelection)
	If @error Then _ERROR($oDoc, "Failed to paste Selection. Error:" & @error & " Extended:" & @extended, $oNewDoc)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended, $oNewDoc)

	; Close the second document.
	_LOCalc_DocClose($oNewDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended, $oNewDoc)
EndFunc

Func _ERROR($oDoc, $sErrorText, $oNewDoc = Null)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	If IsObj($oNewDoc) Then _LOCalc_DocClose($oNewDoc, False)
	Exit
EndFunc
