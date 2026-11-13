#include <Array.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBasic.au3"
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oLibContainer
	Local $sExecuteLink, $sMacro = "Sub AutoIt_Test" & @LF & "MsgBox ""Hello World!""" & @LF & "End Sub" & @LF
	Local $asEvents

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oLibContainer = _LOBasic_MacroLibContainerGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleCreate($oLibContainer, "Standard", "AutoIt_Macro1")
	If @error Then _ERROR($oDoc, "Failed to create a module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleMacros($oLibContainer, "Standard", "AutoIt_Macro1", $sMacro)
	If @error Then _ERROR($oDoc, "Failed to set the Macros contained in a Module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sExecuteLink = _LOBasic_MacroExecuteLinkCreate($oLibContainer, "Standard", "AutoIt_Macro1", "AutoIt_Test")
	If @error Then _ERROR($oDoc, "Failed to create a Macro Execution link. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of All events
	$asEvents = _LOWriter_DocEvents($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of events. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & @extended & " events found. I will now display the results using _ArrayDisplay. There will be two columns, " & @CRLF & _
			"-the first column contains the event name, " & @CRLF & _
			"-the second column contains the macro assigned to it, if any.")

	_ArrayDisplay($asEvents, Default, Default, 0, Default, "Event Name|Macro")

	$asEvents[0][1] = $sExecuteLink

	; Assign an event
	_LOWriter_DocEvents($oDoc, $asEvents)
	If @error Then _ERROR($oDoc, "Failed to assign an event. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of All events
	$asEvents = _LOWriter_DocEvents($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of events. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_ArrayDisplay($asEvents, Default, Default, 0, Default, "Event Name|Macro")

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
