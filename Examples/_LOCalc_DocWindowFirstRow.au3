#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $iRow, $iNewRow

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the first visible Row.
	$iRow = _LOCalc_DocWindowFirstRow($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve first visible Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iNewRow = Int(InputBox("", "The Currently first visible Row is " & $iRow & ". Please enter a new Row number to set as the first visible Row.", "10", " M"))

	; Set the first visible Row to the entered value.
	_LOCalc_DocWindowFirstRow($oDoc, $iNewRow)
	If @error Then _ERROR($oDoc, "Failed to set first visible Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
