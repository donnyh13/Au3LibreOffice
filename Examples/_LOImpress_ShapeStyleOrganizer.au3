#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oStyle
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Shape Style called AutoIt-Test
	$oStyle = _LOImpress_ShapeStyleCreate($oDoc, "AutoIt-Test")
	If @error Then _ERROR($oDoc, "Failed to create a Shape Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Shape Style's name to "New-AutoIt-Test", parent style = "Graphic", hidden to False
	_LOImpress_ShapeStyleOrganizer($oDoc, $oStyle, "New-AutoIt-Test", "Graphic", False)
	If @error Then _ERROR($oDoc, "Failed to modify Shape Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOImpress_ShapeStyleOrganizer($oDoc, $oStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Shape style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape Style's current Organizer settings are as follows: " & @CRLF & _
			"The Shape Style's name is: " & $avSettings[0] & @CRLF & _
			"The Parent Style's name is: " & $avSettings[1] & @CRLF & _
			"Is this style hidden in the User Interface? True/False: " & $avSettings[2])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
