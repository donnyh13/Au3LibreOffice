#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBasic.au3"
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oLibContainer
	Local $sExecuteLink, $sMacro = "Sub AutoIt_Test" & @LF & 'ThisComponent.Text.String = "Hello World!"' & @LF & "End Sub" & @LF
	Local $avValues

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oLibContainer = _LOBasic_MacroLibContainerGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleCreate($oLibContainer, "Standard", "Au3_Test", $sMacro)
	If @error Then _ERROR($oDoc, "Failed to create a Module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sExecuteLink = _LOBasic_MacroExecuteLinkCreate($oLibContainer, "Standard", "Au3_Test", "AutoIt_Test")
	If @error Then _ERROR($oDoc, "Failed to create a Macro Execution link. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current macro execution link is:" & @CRLF & @CRLF & $sExecuteLink & @CRLF & @CRLF & _
			"Press ok to modify it.")

	$sExecuteLink = _LOBasic_MacroExecuteLinkModify($sExecuteLink, "MyLibrary", "This_Module", "Great_Macro", "Basic", False)
	If @error Then _ERROR($oDoc, "Failed to modify Macro link's values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current values. Return will be an array with element values in order of function parameters.
	$avValues = _LOBasic_MacroExecuteLinkModify($sExecuteLink)
	If @error Then _ERROR($oDoc, "Failed to retrieve Macro link's current values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Macro Execution link's values are as follows: " & @CRLF & _
			"The Library name is: " & $avValues[0] & @CRLF & _
			"The Module name is: " & $avValues[1] & @CRLF & _
			"The Macro name is: " & $avValues[2] & @CRLF & _
			"The Language Type is: " & $avValues[3] & @CRLF & _
			"Is the macro stored in the document? True/False: " & $avValues[4])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document, don't save changes.
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
