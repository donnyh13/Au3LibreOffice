#include <Array.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"
#include "..\LibreOfficeBasic.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oLibContainer
	Local $sSavePath, $sExecuteLink, $sMacro = "Sub AutoIt_Test" & @LF & "MsgBox ""Hello World!""" & @LF & "End Sub" & @LF
	Local $asEvents

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oLibContainer = _LOBasic_MacroLibContainerGetObj($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleCreate($oLibContainer, "Standard", "AutoIt-Macro1")
	If @error Then Return _ERROR($oDoc, "Failed to create a module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleMacros($oLibContainer, "Standard", "AutoIt-Macro1", $sMacro)
	If @error Then Return _ERROR($oDoc, "Failed to set the Macros contained in a Module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sExecuteLink = _LOBasic_MacroExecuteLinkCreate($oLibContainer, "Standard", "AutoIt-Macro1", "AutoIt_Test")
	If @error Then Return _ERROR($oDoc, "Failed to create a Macro Execution link. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of All events
	$asEvents = _LOBase_DocEvents($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve Array of events. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & @extended & " events found. I will now display the results using _ArrayDisplay. There will be two columns, " & @CRLF & _
			"-the first column contains the event name, " & @CRLF & _
			"-the second column contains the macro assigned to it, if any.")

	_ArrayDisplay($asEvents, Default, Default, 0, Default, "Event Name|Macro")

	$asEvents[0][1] = $sExecuteLink

	; Assign an event
	_LOBase_DocEvents($oDoc, $asEvents)
	If @error Then Return _ERROR($oDoc, "Failed to assign an event. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_ArrayDisplay($asEvents, Default, Default, 0, Default, "Event Name|Macro")

	; Retrieve an Array of All events
	$asEvents = _LOBase_DocEvents($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve Array of events. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	Exit
EndFunc
