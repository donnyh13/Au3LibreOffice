#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oParStyle
	Local $bExists

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Paragraph Style to use for demonstration.
	$oParStyle = _LOWriter_ParStyleCreate($oDoc, "NewParStyle")
	If @error Then _ERROR($oDoc, "Failed to Create a new Paragraph Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the paragraph style exists.
	$bExists = _LOWriter_ParStyleExists($oDoc, "NewParStyle")
	If @error Then _ERROR($oDoc, "Failed to test for Paragraph Style existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a Paragraph style called ""NewParStyle"" exist in the document? True/False: " & $bExists)

	; Delete the paragraph style, Force delete it, if it is in use, and Replacement it with paragraph style, Default Paragraph Style
	_LOWriter_ParStyleDelete($oDoc, $oParStyle, True, "Standard")
	If @error Then _ERROR($oDoc, "Failed to delete the Paragraph Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the paragraph style still exists.
	$bExists = _LOWriter_ParStyleExists($oDoc, "NewParStyle")
	If @error Then _ERROR($oDoc, "Failed to test for Paragraph Style existing in document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a Paragraph style called ""NewParStyle"" still exist in the document? True/False: " & $bExists)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
