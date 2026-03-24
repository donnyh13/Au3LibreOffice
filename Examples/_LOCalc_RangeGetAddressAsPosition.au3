#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $asAddress[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to C5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "C5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Cell Range Address as an array of positions.
	$asAddress = _LOCalc_RangeGetAddressAsPosition($oCellRange)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Address. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Address for Cell Range A1 to C5 in Sheet 1 is: " & @CRLF & _
			"The Sheet index number is: " & $asAddress[0] & @CRLF & _
			"I can use this to retrieve the Sheet by Position." & @CRLF & _
			"The top-left cell's Column index of the Cell Range is: " & $asAddress[1] & @CRLF & _
			"The top-left cell's Row index of the Cell Range is: " & $asAddress[2] & @CRLF & _
			"The bottom-right cell's Column index of the Cell Range is: " & $asAddress[3] & @CRLF & _
			"The bottom-right cell's Row index of the Cell Range is: " & $asAddress[4] & @CRLF & _
			"Remember Columns and Rows are 0 based.")

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
