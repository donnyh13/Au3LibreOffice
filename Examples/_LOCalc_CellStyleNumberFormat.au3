#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellStyle, $oCell
	Local $iFormatKey, $iReturn

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

	; Create or Retrieve the Numbering Format "#,##0"
	$iFormatKey = _LOCalc_FormatKeyCreate($oDoc, "#,##0")
	If @error Then _ERROR($oDoc, "Failed to Create or Retrieve Number Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Object for Default Cell Style.
	$oCellStyle = _LOCalc_CellStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve the Object for Cell Style named ""Default"". Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell Style's Number Format to the Format key I just retrieved.
	_LOCalc_CellStyleNumberFormat($oDoc, $oCellStyle, $iFormatKey)
	If @error Then _ERROR($oDoc, "Failed to set Cell Style Number Format. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell Style's current Number Format setting, Return will be an Integer.
	$iReturn = _LOCalc_CellStyleNumberFormat($oDoc, $oCellStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Style's Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The ""Default"" Cell Style Numbering Format Key is: " & $iReturn)

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
