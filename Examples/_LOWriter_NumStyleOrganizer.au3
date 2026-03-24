#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oNumbStyle
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Numbering Style named "Test Style"
	$oNumbStyle = _LOWriter_NumStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to create a Numbering Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Numbering Style Organizer settings. Change the name to "New Numbering Name", and hidden to False
	_LOWriter_NumStyleOrganizer($oDoc, $oNumbStyle, "New Numbering Name", False)
	If @error Then _ERROR($oDoc, "Failed to set Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Numbering Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_NumStyleOrganizer($oDoc, $oNumbStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Numbering style's organizer settings are as follows: " & @CRLF & _
			"The Numbering Style's name is: " & $avSettings[0] & @CRLF & _
			"Is this frame style hidden in the User Interface? True/False: " & $avSettings[1])

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
