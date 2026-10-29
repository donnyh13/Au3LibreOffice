#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"
#include "..\LibreOfficeBasic.au3"

Example()

Func Example()
	Local $oDoc, $oLibContainer
	Local $asModules
	Local $sModules = ""

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oLibContainer = _LOBasic_MacroLibContainerGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroModuleCreate($oLibContainer, "Standard", "AutoIt-Macro1")
	If @error Then _ERROR($oDoc, "Failed to create a module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$asModules = _LOBasic_MacroModulesGetNames($oLibContainer, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Module names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		$sModules &= $asModules[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The ""Standard"" Document Library contains the following Modules:" & @CRLF & $sModules & @CRLF & @CRLF & _
			"Press ok to rename ""AutoIt-Macro1"".")

	_LOBasic_MacroModuleRename($oLibContainer, "Standard", "AutoIt-Macro1", "AutoIt-TestModule")
	If @error Then _ERROR($oDoc, "Failed to rename Module. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$asModules = _LOBasic_MacroModulesGetNames($oLibContainer, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Module names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sModules = ""

	For $i = 0 To @extended - 1
		$sModules &= $asModules[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The ""Standard"" Document Library now contains the following Modules:" & @CRLF & $sModules)

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
