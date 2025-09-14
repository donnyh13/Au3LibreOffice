#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $sSavePath, $sPath
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test if the document has a save location/Path.
	$bReturn = _LOImpress_DocHasPath($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document information. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does this document have a Save location/Path? True/False: " & $bReturn)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odp")

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOImpress_DocSaveAs($oDoc, $sSavePath, "", True)
	If @error Then _ERROR($oDoc, "Failed to save the Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test again if the document has a save location/Path.
	$bReturn = _LOImpress_DocHasPath($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document information. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does this document have a Save location/Path? True/False: " & $bReturn)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the file.
	FileDelete($sPath)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
