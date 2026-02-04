#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iColor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Turn Header on.
	_LOWriter_PageStyleHeader($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to turn Page Style headers on. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Header Background color to $LO_COLOR_LIME
	_LOWriter_PageStyleHeaderAreaColor($oPageStyle, $LO_COLOR_LIME)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an integer.
	$iColor = _LOWriter_PageStyleHeaderAreaColor($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Header Background color is (as a RGB Color Integer): " & $iColor)

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
