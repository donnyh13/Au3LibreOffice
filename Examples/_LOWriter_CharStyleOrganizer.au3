#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oCharStyle
	Local $avCharStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Character Style for demonstration.
	$oCharStyle = _LOWriter_CharStyleCreate($oDoc, "NewCharStyle")
	If @error Then _ERROR($oDoc, "Failed to create Character style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the new Character style's name to "New-Char-Name", set the parent style to "Example" Character style, and hidden to false.
	_LOWriter_CharStyleOrganizer($oDoc, $oCharStyle, "New-Char-Name", "Example", False)
	If @error Then _ERROR($oDoc, "Failed to modify Character style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avCharStyleSettings = _LOWriter_CharStyleOrganizer($oDoc, $oCharStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Character style's current Organizer settings are as follows: " & @CRLF & _
			"The Character Style's name is: " & $avCharStyleSettings[0] & @CRLF & _
			"The Parent Character Style of this style is: " & $avCharStyleSettings[1] & @CRLF & _
			"Is this style hidden in the User Interface? True/False: " & $avCharStyleSettings[2])

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
