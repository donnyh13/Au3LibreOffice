#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $sSavePath, $sPath, $sReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odt")

	; Save The New Blank Doc To Desktop Directory.
	$sPath = _LOWriter_DocSaveAs($oDoc, $sSavePath, "", True)
	If @error Then _ERROR("Failed to save the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Document's Save Path if it has one, the return will be a string, and the path will be like a computer path.
	If _LOWriter_DocHasPath($oDoc) Then $sReturn = _LOWriter_DocGetPath($oDoc, False)
	If (@error > 0) Or ($sReturn = "") Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The document's save location/Path is: " & $sReturn)

	; Retrieve the Document's Save Path again, the return will be a string, and the path will be a Libre Office URL.
	$sReturn = _LOWriter_DocGetPath($oDoc, True)
	If (@error > 0) Or ($sReturn = "") Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The document's save location/Path, in Libre Office URL format, is: " & $sReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	; Delete the file.
	FileDelete($sPath)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
