#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $iColumn, $iNewColumn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the first visible column.
	$iColumn = _LOCalc_DocWindowFirstColumn($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve first visible Column. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iNewColumn = Int(InputBox("", "The Currently first visible Column is " & $iColumn & ". Please enter a new column number to set as the first visible column.", "10", " M"))

	; Set the first visible column to the entered value.
	_LOCalc_DocWindowFirstColumn($oDoc, $iNewColumn)
	If @error Then _ERROR($oDoc, "Failed to set first visible Column. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
