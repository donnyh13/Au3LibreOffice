#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCell
	Local $sData
	Local $aavData[4]
	Local $avRowData[2]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number/String Values I want in Column A.
	$avRowData[0] = "=1+ 1" ; A1
	$avRowData[1] = "=SUM(A1:A3)" ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = "=70 * 7" ; A2
	$avRowData[1] = "=ROUND(18 / 5.5)" ; B2 ;
	$aavData[1] = $avRowData

	$avRowData[0] = "=50 / 10" ; A3
	$avRowData[1] = "=2 + 2"  ; B3
	$aavData[2] = $avRowData

	$avRowData[0] = "=5 * 5" ; A4
	$avRowData[1] = "=A3 + 5" ; B4
	$aavData[3] = $avRowData

	; Retrieve Cell range A1 to B4
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now fill Cell Range A1 to B4 with Formulas." & @CRLF & _
			"I will then replace Cell A4 with a String, and B3 with a Number, to demonstrate what is returned by _LOCalc_RangeFormula when it encounters these data types.")

	; Fill the range with Data
	_LOCalc_RangeFormula($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A4
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell A4 to a String
	_LOCalc_CellString($oCell, "A String")
	If @error Then _ERROR($oDoc, "Failed to fill Cell with text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell B3
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B3 to a Number
	_LOCalc_CellValue($oCell, 87)
	If @error Then _ERROR($oDoc, "Failed to fill Cell with text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the formulas from the Cell Range A1-B4. Return will be an array of Arrays
	$aavData = _LOCalc_RangeFormula($oCellRange)
	If @error Then _ERROR($oDoc, "Failed to numbers in Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($aavData) - 1
		For $j = 0 To UBound($aavData[$i]) - 1
			$sData &= "Column " & $j & ": " & ($aavData[$i])[$j] & @CRLF
		Next

		MsgBox($MB_OK + $MB_TOPMOST, Default, "Array $aavData[" & $i & "] contains the following Formulas:" & @CRLF & $sData)
		$sData = ""
	Next

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
