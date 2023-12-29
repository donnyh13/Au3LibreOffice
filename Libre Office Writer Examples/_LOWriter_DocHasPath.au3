
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>
#include <File.au3>

Example()

Func Example()
	Local $oDoc
	Local $sSavePath, $sPath
	Local $bReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Test if the document has a save location/Path.
	$bReturn = _LOWriter_DocHasPath($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does this document have a Save location/Path? True/False: " & $bReturn)

	;Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odt")

	;Save The New Blank Doc To Desktop Directory.
	$sPath = _LOWriter_DocSaveAs($oDoc, $sSavePath, "", True)
	If (@error > 0) Then _ERROR("Failed to save the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Test again if the document has a save location/Path.
	$bReturn = _LOWriter_DocHasPath($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does this document have a Save location/Path? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	;Delete the file.
	FileDelete($sPath)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



