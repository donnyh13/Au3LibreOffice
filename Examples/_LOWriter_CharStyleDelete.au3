#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oCharStyle
	Local $bExists

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Character Style for demonstration.
	$oCharStyle = _LOWriter_CharStyleCreate($oDoc, "NewCharStyle")
	If @error Then _ERROR($oDoc, "Failed to create Character style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the Character style exists.
	$bExists = _LOWriter_CharStyleExists($oDoc, "NewCharStyle")
	If @error Then _ERROR($oDoc, "Failed to test for Character Style existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a Character style called ""NewCharStyle"" exist in the document? True/False: " & $bExists)

	; Delete the Character style, Force delete if it is in use, and replace it with Character style "Example".
	_LOWriter_CharStyleDelete($oDoc, $oCharStyle, True, "Example")

	; Check if the Character style still exists.
	$bExists = _LOWriter_CharStyleExists($oDoc, "NewCharStyle")
	If @error Then _ERROR($oDoc, "Failed to test for Character Style existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a Character style called ""NewCharStyle"" still exist in the document? True/False: " & $bExists)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
