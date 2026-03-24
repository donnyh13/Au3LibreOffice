#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $sRanges
	Local $avSettings
	Local $aavData[5]
	Local $avRowData[2]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired String Values I want in Column A and B.
	$avRowData[0] = "These" ; A1
	$avRowData[1] = "Entire" ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = "Columns" ; A2
	$avRowData[1] = "Will" ; B2
	$aavData[1] = $avRowData

	$avRowData[0] = "Be" ; A3
	$avRowData[1] = "Repeated" ; B3
	$aavData[2] = $avRowData

	$avRowData[0] = "On" ; A4
	$avRowData[1] = "Printed" ; B4
	$aavData[3] = $avRowData

	$avRowData[0] = "Pages to" ; A5
	$avRowData[1] = "The Right." ; B5
	$aavData[4] = $avRowData

	; Retrieve Cell range A1 to B5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to B1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Column Heading Range to A & B, and set repeat column Headings to True.
	_LOCalc_SheetPrintColumnsRepeat($oSheet, $oCellRange, True)
	If @error Then _ERROR($oDoc, "Failed to set Column Heading Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have set the Columns A1 to B1 to repeat automatically on every printed page to the right.")

	; Retrieve the current print ranges set for this sheet.
	$avSettings = _LOCalc_SheetPrintColumnsRepeat($oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column repeat Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell Range Address.
	$sRanges &= _LOCalc_RangeGetAddressAsName($avSettings[0])
	If @error Then _ERROR($oDoc, "Failed to retrieve Range address. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Range(s) currently set to be repeated as Column Headers are: " & @CRLF & $sRanges & @CRLF & @CRLF & _
			"I will now reset the Column Repeat Range.")

	; Reset the Row repeat Range to the whole sheet.
	_LOCalc_SheetPrintColumnsRepeat($oSheet, Default)
	If @error Then _ERROR($oDoc, "Failed to set Repeat Column Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current print ranges set for this sheet.
	$avSettings = _LOCalc_SheetPrintColumnsRepeat($oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Repeat Range array. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sRanges = ""

	; Retrieve Cell Range Addresses.
	$sRanges &= _LOCalc_RangeGetAddressAsName($avSettings[0])
	If @error Then _ERROR($oDoc, "Failed to retrieve Range address. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Range(s) currently set to be repeated as Column Headers are: " & @CRLF & $sRanges)

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
