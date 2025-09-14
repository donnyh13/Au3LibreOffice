#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $sFilePathName, $sPath

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now export the new Impress Document as a pdf to the desktop folder.")

	$sFilePathName = _TempFile(@DesktopDir & "\", "TestExportDoc_", ".pdf")

	; Export The New Blank Doc To Desktop Directory as a PDF using a unique temporary name.
	$sPath = _LOImpress_DocExport($oDoc, $sFilePathName, False)
	If @error Then _ERROR($oDoc, "Failed to Export the Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created and exported the document as a PDF to your Desktop, found at the following Path: " _
			& $sPath & @CRLF & "Press Ok to delete it.")

	; Delete the file.
	FileDelete($sPath)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
