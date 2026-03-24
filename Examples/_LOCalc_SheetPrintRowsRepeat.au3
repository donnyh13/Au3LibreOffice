#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $sRanges
	Local $avSettings
	Local $aavData[1]
	Local $avRowData[5]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired String Values I want in Column A to E.
	$avRowData[0] = "Heading 1" ; A1
	$avRowData[1] = "Heading 2" ; B1
	$avRowData[2] = "Heading 3" ; C1
	$avRowData[3] = "Heading 4" ; D1
	$avRowData[4] = "Heading 5" ; E1
	$aavData[0] = $avRowData

	; Retrieve Cell range A1 to E1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "E1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Row Heading Range to Row 1, and set repeat Row Headings to True.
	_LOCalc_SheetPrintRowsRepeat($oSheet, $oCellRange, True)
	If @error Then _ERROR($oDoc, "Failed to set Row Heading Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have set Row 1 to repeat automatically on every printed page to the bottom.")

	; Retrieve the current print ranges set for this sheet.
	$avSettings = _LOCalc_SheetPrintRowsRepeat($oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve Row repeat Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell Range Address.
	$sRanges &= _LOCalc_RangeGetAddressAsName($avSettings[0])
	If @error Then _ERROR($oDoc, "Failed to retrieve Range address. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Range(s) currently set to be repeated as Row Headers are: " & @CRLF & $sRanges & @CRLF & @CRLF & _
			"I will now reset the Row Repeat Range.")

	; Reset the Row repeat Range to the whole sheet.
	_LOCalc_SheetPrintRowsRepeat($oSheet, Default)
	If @error Then _ERROR($oDoc, "Failed to set Repeat Row Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current print ranges set for this sheet.
	$avSettings = _LOCalc_SheetPrintRowsRepeat($oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve Row Repeat Range array. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sRanges = ""

	; Retrieve Cell Range Addresses.
	$sRanges &= _LOCalc_RangeGetAddressAsName($avSettings[0])
	If @error Then _ERROR($oDoc, "Failed to retrieve Range address. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Range(s) currently set to be repeated as Row Headers are: " & @CRLF & $sRanges)

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
