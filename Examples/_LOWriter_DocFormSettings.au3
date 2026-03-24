#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $avReturn, $abBackup

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Backup the Document's Form settings.
	$abBackup = _LOWriter_DocFormSettings($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document Form settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Document's form settings to: Form Design Mode = True, Automatically open in Form Design Mode = True, Control Focus = False, Control Wizards = False.
	_LOWriter_DocFormSettings($oDoc, True, True, False, False)
	If @error Then _ERROR($oDoc, "Failed to modify Document Form settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Document's Form settings. Return will be an Array in order of function parameters.
	$avReturn = _LOWriter_DocFormSettings($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document Form settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The document's Form settings are: " & @CRLF & _
			"Is the Document currently in Form Design Mode? True/False: " & $avReturn[0] & @CRLF & _
			"When the Document is opened, is Form Design Mode active? True/False: " & $avReturn[1] & @CRLF & _
			"When the Document opens, are Form Controls to receive focus automatically? True/False: " & $avReturn[2] & @CRLF & _
			"Are Form Control Wizards used? True/False: " & $avReturn[3])

	; Reset the Document's form settings to their former values.
	_LOWriter_DocFormSettings($oDoc, $abBackup[0], $abBackup[1], $abBackup[2], $abBackup[3])
	If @error Then _ERROR($oDoc, "Failed to modify Document Form settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
