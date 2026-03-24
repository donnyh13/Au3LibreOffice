#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $sRedo

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the presently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the top left most cell, 0, 0, or A1.
	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Cell text to "A1"
	_LOCalc_CellString($oCell, "A1")
	If @error Then _ERROR($oDoc, "Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the C2 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C2")
	If @error Then _ERROR($oDoc, "Failed to retrieve C2 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set C2 Cell text to "C2"
	_LOCalc_CellString($oCell, "C2")
	If @error Then _ERROR($oDoc, "Failed to Set C2 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Perform one undo action.
	_LOCalc_DocUndo($oDoc)
	If @error Then _ERROR($oDoc, "Failed to perform an undo action. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the next available Redo action title.
	$sRedo = _LOCalc_DocRedoCurActionTitle($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve next Redo action title. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The next Redo action title is: " & $sRedo & @CRLF & " Press ok to perform it.")

	; Perform one Redo action.
	_LOCalc_DocRedo($oDoc)
	If @error Then _ERROR($oDoc, "Failed to perform an Redo action. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
