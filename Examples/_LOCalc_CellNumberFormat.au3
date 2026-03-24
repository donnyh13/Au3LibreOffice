#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oCellRange
	Local $iFormatKey, $iReturn1, $iReturn2

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill Cells A1 to A6 with values.
	For $i = 0 To 5
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, $i)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		_LOCalc_CellValue($oCell, Int($i + 1 & "0000"))
		If @error Then _ERROR($oDoc, "Failed to set Cell value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Retrieve Cell B5
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B5's Value to 1000
	_LOCalc_CellValue($oCell, 1000)
	If @error Then _ERROR($oDoc, "Failed to set Cell value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create or Retrieve the Numbering Format "#,##0"
	$iFormatKey = _LOCalc_FormatKeyCreate($oDoc, "#,##0")
	If @error Then _ERROR($oDoc, "Failed to Create or Retrieve Number Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B5's Numbering Format to the format key I just retrieved.
	_LOCalc_CellNumberFormat($oDoc, $oCell, $iFormatKey)
	If @error Then _ERROR($oDoc, "Failed to set Cell Numbering Format. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to A6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell A1-A6's Numbering Format to the format key I just retrieved.
	_LOCalc_CellNumberFormat($oDoc, $oCellRange, $iFormatKey)
	If @error Then _ERROR($oDoc, "Failed to set Cell Range Numbering Format. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell B5's current Numbering Format setting, Return will be an Integer.
	$iReturn1 = _LOCalc_CellNumberFormat($oDoc, $oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Numbering Format setting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell Range A1-A6's current Numbering Format setting, Return will be an Integer.
	$iReturn2 = _LOCalc_CellNumberFormat($oDoc, $oCellRange)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Numbering Format setting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Cell B5's Numbering Format Key is: " & $iReturn1 & @CRLF & _
			"Cell Range A1-A6's Numbering Format Key is: " & $iReturn2)

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
