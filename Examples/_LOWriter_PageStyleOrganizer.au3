#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New Page Style.
	$oPageStyle = _LOWriter_PageStyleCreate($oDoc, "NewPageStyle")
	If @error Then _ERROR($oDoc, "Failed to create a new Page Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Page Style's name to "New-PageStyle-Name", hidden to False, Set the follow style to "HTML"
	_LOWriter_PageStyleOrganizer($oDoc, $oPageStyle, "New-PageStyle-Name", False, "HTML")
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleOrganizer($oDoc, $oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Organizer settings are as follows: " & @CRLF & _
			"The Page Style's name is: " & $avPageStyleSettings[0] & @CRLF & _
			"Is this style hidden in the User Interface? True/False: " & $avPageStyleSettings[1] & @CRLF & _
			"The Page Style's name that comes after this style is: " & $avPageStyleSettings[2])

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
