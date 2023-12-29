#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $sFilePathName, $sPath

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now export the new Calc Document as a pdf to the desktop folder.")

	$sFilePathName = _TempFile(@DesktopDir & "\", "TestExportDoc_", ".pdf")

	; Export The New Blank Doc To Desktop Directory as a PDF using a unique temporary name.
	$sPath = _LOCalc_DocExport($oDoc, $sFilePathName, False)
	If @error Then _ERROR("Failed to Export the Calc Document. Error:" & @error & " Extended:" & @extended)

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have created and exported the document as a PDF to your Desktop, found at the following Path: " _
			 & $sPath & @CRLF & "Press Ok to delete it.")

	; Delete the file.
	FileDelete($sPath)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
