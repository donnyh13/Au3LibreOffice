#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oSheet2, $oCell, $oCellRange
	Local $iCount = 0

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the Cell Range of B2 to D4 with numbers, one cell at a time. (Remember Columns and Rows are 0 based.)
	For $i = 1 To 3
		For $j = 1 To 3
			; Retrieve the Cell Object
			$oCell = _LOCalc_RangeGetCellByPosition($oSheet, $i, $j)
			If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			; Set the Cell to a Number
			_LOCalc_CellValue($oCell, $iCount)
			If @error Then _ERROR($oDoc, "Failed to set Cell Value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			$iCount += 1
		Next
	Next

	; Retrieve Cell Range B2-D4
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B2", "D4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell D6
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "D6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will move the cell Range B2-D4 to begin at cell D6")

	; Copy the Cell Range to Cell D6
	_LOCalc_RangeCopyMove($oSheet, $oCellRange, $oCell, False)
	If @error Then _ERROR($oDoc, "Failed to move Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Second Sheet
	$oSheet2 = _LOCalc_SheetAdd($oDoc, "Sheet 2")
	If @error Then _ERROR($oDoc, "Failed to create a new sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell C7 on Sheet 2
	$oCell = _LOCalc_RangeGetCellByName($oSheet2, "C7")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I am going to move Cell Range B2 to D4 from Sheet 1 to Cell C7 on Sheet 2, which I just created.")

	; Move the Cell Range B2-D4 Sheet 1 to Cell C7 Sheet 2.
	_LOCalc_RangeCopyMove($oSheet, $oCellRange, $oCell, True)
	If @error Then _ERROR($oDoc, "Failed to move Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Notice the Cell Range B2 to D4 is now Blank. Press ok to switch to Sheet 2, and you will see the Cell Range Data there.")

	; Switch to Sheet 2
	_LOCalc_SheetActivate($oDoc, $oSheet2)
	If @error Then _ERROR($oDoc, "Failed to switch active sheets. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
