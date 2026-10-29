#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"
#include "..\LibreOfficeBasic.au3"

Example()

Func Example()
	Local $oDoc, $oLibContainer
	Local $bVerified
	Local $sFilePathName, $sPath

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sFilePathName = _TempFile(@DesktopDir & "\", "TestExportDoc_", ".odt")

	$oLibContainer = _LOBasic_MacroLibContainerGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroLibraryCreate($oLibContainer, "Au3-Test")
	If @error Then _ERROR($oDoc, "Failed to create a Library. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroLibraryPasswordChange($oLibContainer, "Au3-Test", "", "AutoIt")
	If @error Then _ERROR($oDoc, "Failed to apply a password to library. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; save The New Blank Doc To Desktop Directory using a unique temporary name.
	$sPath = _LOWriter_DocSaveAs($oDoc, $sFilePathName)
	If @error Then _ERROR($oDoc, "Failed to Save the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOBasic_MacroLibrariesSave($oLibContainer)
	If @error Then _ERROR($oDoc, "Failed to save modified libraries. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the document.
	$oDoc = _LOWriter_DocOpen($sPath, True, True)
	If @error Then _ERROR($oDoc, "Failed to open Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Make the document visible by setting visible to True
	_LOWriter_DocVisible($oDoc, True)
	If @error Then _ERROR($oDoc, "Failed to change Document visibility settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oLibContainer = _LOBasic_MacroLibContainerGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$bVerified = _LOBasic_MacroLibraryPasswordIsVerified($oLibContainer, "Au3-Test")
	If @error Then _ERROR($oDoc, "Failed to check if a Library has been verified by its password. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Has the ""Au3-Test"" Document Library been verified by its password? True/False: " & $bVerified)

	_LOBasic_MacroLibraryPasswordVerify($oLibContainer, "Au3-Test", "AutoIt")
	If @error Then _ERROR($oDoc, "Failed to verify Library password. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$bVerified = _LOBasic_MacroLibraryPasswordIsVerified($oLibContainer, "Au3-Test")
	If @error Then _ERROR($oDoc, "Failed to check if a Library has been verified by its password. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now has the ""Au3-Test"" Document Library been verified by its password? True/False: " & $bVerified)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document, don't save changes.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the file.
	FileDelete($sPath)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
