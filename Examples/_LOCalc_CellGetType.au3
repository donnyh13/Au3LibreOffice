#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $iCellType

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the presently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the top left most cell, 0, 0, or A1.
	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Cell Value to 20
	_LOCalc_CellValue($oCell, 20)
	If @error Then _ERROR($oDoc, "Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A1's Type
	$iCellType = _LOCalc_CellGetType($oCell)
	If @error Then _ERROR($oDoc, "Failed to Retrieve Cell Type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Cell A1's Data Type is: " & $iCellType & @CRLF & @CRLF & _
			"$LOC_CELL_TYPE_EMPTY = 0," & @CRLF & _
			"$LOC_CELL_TYPE_VALUE = 1," & @CRLF & _
			"$LOC_CELL_TYPE_TEXT = 2," & @CRLF & _
			"$LOC_CELL_TYPE_FORMULA = 3")

	; Retrieve the A2 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A2")
	If @error Then _ERROR($oDoc, "Failed to retrieve A2 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A2's Type
	$iCellType = _LOCalc_CellGetType($oCell)
	If @error Then _ERROR($oDoc, "Failed to Retrieve Cell Type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Cell A2's Data Type is: " & $iCellType & @CRLF & @CRLF & _
			"$LOC_CELL_TYPE_EMPTY = 0," & @CRLF & _
			"$LOC_CELL_TYPE_VALUE = 1," & @CRLF & _
			"$LOC_CELL_TYPE_TEXT = 2," & @CRLF & _
			"$LOC_CELL_TYPE_FORMULA = 3")

	; Retrieve the A3 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A3")
	If @error Then _ERROR($oDoc, "Failed to retrieve A3 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A3 Cell text to "Equals"
	_LOCalc_CellString($oCell, "Equals")
	If @error Then _ERROR($oDoc, "Failed to Set A3 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A3's Type
	$iCellType = _LOCalc_CellGetType($oCell)
	If @error Then _ERROR($oDoc, "Failed to Retrieve Cell Type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Cell A3's Data Type is: " & $iCellType & @CRLF & @CRLF & _
			"$LOC_CELL_TYPE_EMPTY = 0," & @CRLF & _
			"$LOC_CELL_TYPE_VALUE = 1," & @CRLF & _
			"$LOC_CELL_TYPE_TEXT = 2," & @CRLF & _
			"$LOC_CELL_TYPE_FORMULA = 3")

	; Retrieve the A4 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A4")
	If @error Then _ERROR($oDoc, "Failed to retrieve A4 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A4 Cell formula to "=A1 * A2 + 2
	_LOCalc_CellFormula($oCell, "=A1 * A2 + 2")
	If @error Then _ERROR($oDoc, "Failed to Set A4 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A4's Type
	$iCellType = _LOCalc_CellGetType($oCell)
	If @error Then _ERROR($oDoc, "Failed to Retrieve Cell Type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Cell A4's Data Type is: " & $iCellType & @CRLF & @CRLF & _
			"$LOC_CELL_TYPE_EMPTY = 0," & @CRLF & _
			"$LOC_CELL_TYPE_VALUE = 1," & @CRLF & _
			"$LOC_CELL_TYPE_TEXT = 2," & @CRLF & _
			"$LOC_CELL_TYPE_FORMULA = 3")

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
