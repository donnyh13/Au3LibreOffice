#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oParStyle
	Local $avParStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Paragraph Style to use for demonstration.
	$oParStyle = _LOWriter_ParStyleCreate($oDoc, "NewParStyle")
	If @error Then _ERROR($oDoc, "Failed to Create a new Paragraph Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the paragraph Style's name to "New-Par-Name", Set the follow style to "List", the parent style to Default Paragraph Style,
	; And Auto update To True, and hidden to False
	_LOWriter_ParStyleOrganizer($oDoc, $oParStyle, "New-Par-Name", "List", "Standard", True, False)
	If @error Then _ERROR($oDoc, "Failed to modify Paragraph Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avParStyleSettings = _LOWriter_ParStyleOrganizer($oDoc, $oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Organizer settings are as follows: " & @CRLF & _
			"The Paragraph Style's name is: " & $avParStyleSettings[0] & @CRLF & _
			"The Paragraph Style's name that comes after this style is: " & $avParStyleSettings[1] & @CRLF & _
			"The Parent paragraph Style of this style is: " & $avParStyleSettings[2] & @CRLF & _
			"Does this style auto update its settings? True/False: " & $avParStyleSettings[3] & @CRLF & _
			"Is this style hidden in the User Interface? True/False: " & $avParStyleSettings[4])

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
