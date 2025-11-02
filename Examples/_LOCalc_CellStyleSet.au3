#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oCellStyle, $oSheet, $oCellRange

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the currently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve active sheet object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Cell Range A2 to B3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A2", "B3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Object for Cell Style "Status".
	$oCellStyle = _LOCalc_CellStyleGetObj($oDoc, "Status")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell Style background color for "Status" to Teal.
	_LOCalc_CellStyleBackColor($oCellStyle, $LO_COLOR_TEAL)
	If @error Then _ERROR($oDoc, "Failed to set Cell Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now set the current Cell style to ""Status."" for Cell Range A2-B3.")

	; Set the Cell range to Cell style "Status"
	_LOCalc_CellStyleSet($oDoc, $oCellRange, "Status")
	If @error Then _ERROR($oDoc, "Failed to set the Cell style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
