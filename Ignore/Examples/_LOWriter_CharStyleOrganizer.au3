#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oCharStyle
	Local $avCharStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a new Character Style for demonstration.
	$oCharStyle = _LOWriter_CharStyleCreate($oDoc, "NewCharStyle")
	If @error Then _ERROR("Failed to create Character style. Error:" & @error & " Extended:" & @extended)

	; Modify the new Character style's name to "New-Char-Name", set the parent style to "Example" Character style, and hidden to false.
	_LOWriter_CharStyleOrganizer($oDoc, $oCharStyle, "New-Char-Name", "Example", False)
	If @error Then _ERROR("Failed to modify Character style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avCharStyleSettings = _LOWriter_CharStyleOrganizer($oDoc, $oCharStyle)
	If @error Then _ERROR("Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Character style's current Organizer settings are as follows: " & @CRLF & _
			"The Character Style's name is: " & $avCharStyleSettings[0] & @CRLF & _
			"The Parent Character Style of this style is: " & $avCharStyleSettings[1] & @CRLF & _
			"Is this style hidden in the User Interface? True/False: " & $avCharStyleSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
