#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Sheet named "New Sheet".
	_LOCalc_SheetAdd($oDoc, "New Sheet")
	If @error Then _ERROR($oDoc, "Failed to create a new Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Sheet named "First Sheet". This will be placed at the end of the list.
	$oSheet = _LOCalc_SheetAdd($oDoc, "First Sheet")
	If @error Then _ERROR($oDoc, "Failed to create a new Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Sheet Object for the sheet at position 0, called "Sheet1"
	$oSheet = _LOCalc_SheetGetObjByPosition($oDoc, 0)
	If @error Then _ERROR($oDoc, "Failed to Retrieve the Sheet's Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Sheet currently at position 0 has the name of: " & _LOCalc_SheetName($oDoc, $oSheet))

	; Retrieve the Sheet Object for the sheet at position 2
	$oSheet = _LOCalc_SheetGetObjByPosition($oDoc, 2)
	If @error Then _ERROR($oDoc, "Failed to Retrieve the Sheet's Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Sheet currently at position 2 has the name of: " & _LOCalc_SheetName($oDoc, $oSheet))

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
