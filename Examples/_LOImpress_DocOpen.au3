#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $sSavePath, $sPath

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will Create and Save a new Impress Doc to begin this example, a screen will flash up and disappear after pressing OK.")

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique temporary name.
	$sSavePath = _TempFile(@TempDir & "\", "DocOpenTestFile_", ".odp")

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOImpress_DocSaveAs($oDoc, $sSavePath, "", True)
	If @error Then _ERROR($oDoc, "Failed to save the Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created and saved a blank L.O. Impress Doc to your Temporary Directory, found at the following Path: " _
			& $sPath & @CRLF & "I will now open it.")

	; Open the document.
	$oDoc = _LOImpress_DocOpen($sPath)
	If @error Then _ERROR($oDoc, "Failed to open Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Document was successfully opened. Press OK to close and delete it.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Following Error codes returned: Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the file.
	FileDelete($sPath)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
