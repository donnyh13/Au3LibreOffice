#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $sSavePath, $sPath

	MsgBox($MB_OK, "", "I will Create and Save a new Writer Doc to begin this example, a screen will flash up and disappear after pressing OK.")

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a unique temporary name.
	$sSavePath = _TempFile(@TempDir & "\", "DocOpenTestFile_", ".odt")

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOWriter_DocSaveAs($oDoc, $sSavePath, "", True)
	If @error Then _ERROR($oDoc, "Failed to save the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have created and saved a blank L.O. Writer Doc to your Temporary Directory, found at the following Path: " _
			& $sPath & @CRLF & "I will now open it.")

	; Open the document.
	$oDoc = _LOWriter_DocOpen($sPath)
	If @error Then _ERROR($oDoc, "Failed to open Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Document was successfully opened. Press OK to close and delete it.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Following Error codes returned: Error:" & _
			@error & " Extended:" & @extended)

	; Delete the file.
	FileDelete($sPath)
EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
