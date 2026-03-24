#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oCellStyle
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Cell Style to use for demonstration.
	$oCellStyle = _LOCalc_CellStyleCreate($oDoc, "NewCellStyle")
	If @error Then _ERROR($oDoc, "Failed to Create a new Cell Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Cell Style's name to "Test-Style", Set the parent style to "Status" Cell Style, and hidden to False
	_LOCalc_CellStyleOrganizer($oDoc, $oCellStyle, "Test-Style", "Status", False)
	If @error Then _ERROR($oDoc, "Failed to modify Cell Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_CellStyleOrganizer($oDoc, $oCellStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cell Style's current Organizer settings are as follows: " & @CRLF & _
			"The Cell Style's name is: " & $avSettings[0] & @CRLF & _
			"The Parent Cell Style of this style is: " & $avSettings[1] & @CRLF & _
			"Is this Cell style hidden in the User Interface? True/False: " & $avSettings[2])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
