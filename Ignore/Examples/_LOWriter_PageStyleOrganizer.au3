#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a New Page Style.
	$oPageStyle = _LOWriter_PageStyleCreate($oDoc, "NewPageStyle")
	If @error Then _ERROR("Failed to create a new Page Style. Error:" & @error & " Extended:" & @extended)

	; Change the Page Style's name to "New-PageStyle-Name", hidden to False, Set the follow style to "HTML", the parent style to
	; "Default Page Style".
	_LOWriter_PageStyleOrganizer($oDoc, $oPageStyle, "New-PageStyle-Name", False, "HTML")
	If @error Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avPageStyleSettings = _LOWriter_PageStyleOrganizer($oDoc, $oPageStyle)
	If @error Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Organizer settings are as follows: " & @CRLF & _
			"The Page Style's name is: " & $avPageStyleSettings[0] & @CRLF & _
			"Is this style hidden in the User Interface? True/False: " & $avPageStyleSettings[1] & @CRLF & _
			"The Page Style's name that comes after this style is: " & $avPageStyleSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
