#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"
#include "..\LibreOfficeBasic.au3"

Example()

Func Example()
	Local $oDoc, $oLibContainer
	Local $sMacros, $sNewMacros = "Sub Test123" & @LF & "MsgBox ""Hello World!""" & @LF & "End Sub" & @LF & @LF & "Function [Au3 Test](sInput As String)" & @LF & "MsgBox sInput" & @LF & "End Function"

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oLibContainer = _LOBasic_MacroLibContainerGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleCreate($oLibContainer, "Standard", "AutoIt-Macro1")
	If @error Then _ERROR($oDoc, "Failed to create a module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleMacros($oLibContainer, "Standard", "AutoIt-Macro1", $sNewMacros)
	If @error Then _ERROR($oDoc, "Failed to set the Macros contained in a Module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sMacros = _LOBasic_MacroModuleMacros($oLibContainer, "Standard", "AutoIt-Macro1")
	If @error Then _ERROR($oDoc, "Failed to retrieve the Macros contained in a Module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The ""AutoIt-Macro1"" Module contained in the ""Standard"" Document Library contains the following Macros: " & @CRLF & $sMacros)

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
