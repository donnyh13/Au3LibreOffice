#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCell
	Local $sData
	Local $aavData[4]
	Local $avRowData[2]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values I want in Column A.
	$avRowData[0] = 1 ; A1
	$avRowData[1] = 0 ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = 457 ; A2
	$avRowData[1] = 2300 ; B2
	$aavData[1] = $avRowData

	$avRowData[0] = 0 ; A3
	$avRowData[1] = 31 ; B3
	$aavData[2] = $avRowData

	$avRowData[0] = 18 ; A4
	$avRowData[1] = 55 ; B4
	$aavData[3] = $avRowData

	; Retrieve Cell range A1 to B4
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now fill Cell Range A1 to B4 with Numbers." & @CRLF & _
			"I will then replace Cell A3 with a String, and B1 with a formula, to demonstrate what is returned by _LOCalc_RangeNumbers when it encounters these data types.")

	; Fill the range with Data
	_LOCalc_RangeNumbers($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A3
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell A3 to a String
	_LOCalc_CellString($oCell, "A String")
	If @error Then _ERROR($oDoc, "Failed to fill Cell with text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell B1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B1 to a Formula
	_LOCalc_CellFormula($oCell, "=A2 + B2")
	If @error Then _ERROR($oDoc, "Failed to fill Cell with text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Numbers from the Cell Range A1-B4. Return will be an array of Arrays
	$aavData = _LOCalc_RangeNumbers($oCellRange)
	If @error Then _ERROR($oDoc, "Failed to numbers in Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($aavData) - 1
		For $j = 0 To UBound($aavData[$i]) - 1
			$sData &= "Column " & $j & ": " & ($aavData[$i])[$j] & @CRLF
		Next

		MsgBox($MB_OK + $MB_TOPMOST, Default, "Array $aavData[" & $i & "] contains the following Numbers:" & @CRLF & $sData)
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
