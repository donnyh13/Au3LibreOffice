#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $sSavePath, $sPath

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will Create and Save a new Impress Doc to begin this example, a screen will flash up and disappear after " _
			& "pressing OK.")

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique file name
	$sSavePath = _TempFile(@DesktopDir & "\", "DocOpenTestFile_", ".odp")

	; Save The New Blank Doc To Desktop Directory.
	$sPath = _LOImpress_DocSaveAs($oDoc, $sSavePath, "", True)
	If @error Then _ERROR($oDoc, "Failed to save the Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created and saved a blank L.O. Impress Doc to your Desktop, found at the following Path: " _
			& $sPath & @CRLF & "Press Ok to delete it.")

	; Delete the file.
	FileDelete($sPath)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
