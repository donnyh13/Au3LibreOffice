#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $sSavePath, $sPath

	MsgBox($MB_OK, "", "I will Create and Save a new Calc Doc to begin this example, a screen will flash up and disappear after " _
			 & "pressing OK.")

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Create a unique file name
	$sSavePath = _TempFile(@DesktopDir & "\", "DocOpenTestFile_", ".ods")

	; Save The New Blank Doc To Desktop Directory.
	$sPath = _LOCalc_DocSaveAs($oDoc, $sSavePath, "", True)
	If @error Then _ERROR("Failed to save the Calc Document. Error:" & @error & " Extended:" & @extended)

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have created and saved a blank L.O. Calc Doc to your Desktop, found at the following Path: " _
			 & $sPath & @CRLF & "Press Ok to delete it.")

	; Delete the file.
	FileDelete($sPath)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
