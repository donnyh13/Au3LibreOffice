#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $sRanges
	Local $aEmpty[0]
	Local $aavData[5]
	Local $avRowData[3]
	Local $aoPrintRanges[2]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values I want in Column A and B.
	$avRowData[0] = 1 ; A1
	$avRowData[1] = 8 ; B1
	$avRowData[2] = 55 ; C1
	$aavData[0] = $avRowData

	$avRowData[0] = 457 ; A2
	$avRowData[1] = 2300 ; B2
	$avRowData[2] = 23 ; C2
	$aavData[1] = $avRowData

	$avRowData[0] = 537 ; A3
	$avRowData[1] = 31 ; B3
	$avRowData[2] = 76 ; C3
	$aavData[2] = $avRowData

	$avRowData[0] = 18 ; A4
	$avRowData[1] = 55 ; B4
	$avRowData[2] = 1200 ; C4
	$aavData[3] = $avRowData

	$avRowData[0] = 537     ; A5
	$avRowData[1] = 31 ; B5
	$avRowData[2] = 81 ; C5
	$aavData[4] = $avRowData

	; Retrieve Cell range A1 to C5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "C5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeNumbers($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to A5
	$aoPrintRanges[0] = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range C1 to C3
	$aoPrintRanges[1] = _LOCalc_RangeGetCellByName($oSheet, "C1", "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Print Ranges to A1:A5 and C1:C3.
	_LOCalc_SheetPrintRangeModify($oSheet, $aoPrintRanges)
	If @error Then _ERROR($oDoc, "Failed to set Print Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have set Range A1:A5 and C1:C3 to be printed only, skipping Ranges B1:B5 and C4:C5, you can test it if you like by printing the sheet to pdf or xps.")

	; Retrieve the current print ranges set for this sheet.
	$aoPrintRanges = _LOCalc_SheetPrintRangeModify($oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve Print Range array. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell Range Addresses.
	For $i = 0 To UBound($aoPrintRanges) - 1
		$sRanges &= _LOCalc_RangeGetAddressAsName($aoPrintRanges[$i]) & @CRLF
		If @error Then _ERROR($oDoc, "Failed to retrieve Range address. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Ranges currently set to print are: " & @CRLF & $sRanges & @CRLF & @CRLF & _
			"I will now reset the Print Range.")

	; Reset the print Range to the whole sheet.
	_LOCalc_SheetPrintRangeModify($oSheet, $aEmpty)
	If @error Then _ERROR($oDoc, "Failed to set Print Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current print ranges set for this sheet.
	$aoPrintRanges = _LOCalc_SheetPrintRangeModify($oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve Print Range array. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sRanges = ""

	; Retrieve Cell Range Addresses. (This wont be called because the array will be empty.)
	For $i = 0 To UBound($aoPrintRanges) - 1
		$sRanges &= _LOCalc_RangeGetAddressAsName($aoPrintRanges[$i]) & @CRLF
		If @error Then _ERROR($oDoc, "Failed to retrieve Range address. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Range(s) currently set to print are: " & @CRLF & $sRanges)

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
