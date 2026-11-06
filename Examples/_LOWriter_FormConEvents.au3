#include <Array.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBasic.au3"
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oLibContainer, $oControl
	Local $sExecuteLink, $sMacro = "Sub AutoIt_Test" & @LF & "MsgBox ""Hello World!""" & @LF & "End Sub" & @LF
	Local $asFormConEvents[1][3]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oLibContainer = _LOBasic_MacroLibContainerGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleCreate($oLibContainer, "Standard", "AutoIt-Macro1")
	If @error Then _ERROR($oDoc, "Failed to create a module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleMacros($oLibContainer, "Standard", "AutoIt-Macro1", $sMacro)
	If @error Then _ERROR($oDoc, "Failed to set the Macros contained in a Module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sExecuteLink = _LOBasic_MacroExecuteLinkCreate($oLibContainer, "Standard", "AutoIt-Macro1", "AutoIt_Test")
	If @error Then _ERROR($oDoc, "Failed to create a Macro Execution link. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_LABEL, 500, 300, 6000, 2000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; I will assign the macro to activate when an error occurs
	$asFormConEvents[0][0] = $LOW_FORM_CON_EVENT_ERROR_OCCURRED
	$asFormConEvents[0][1] = $sExecuteLink
	$asFormConEvents[0][2] = "XLoadListener"

	_LOWriter_FormConEvents($oControl, $asFormConEvents)
	If @error Then _ERROR($oDoc, "Failed to assign form events. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of All events
	$asFormConEvents = _LOWriter_FormConEvents($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of events. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & @extended & " events found. I will now display the results using _ArrayDisplay. There will be three columns, " & @CRLF & _
			"-the first column contains the event name, " & @CRLF & _
			"-the second column contains the macro assigned to it, if any." & @CRLF & _
			"-the third column contains the Event Listener type,.")

	_ArrayDisplay($asFormConEvents, Default, Default, 0, Default, "Event Name|Macro|Event Listener Type")

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
