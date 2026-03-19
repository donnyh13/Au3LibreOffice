
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>
#include <File.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $sFilePathName, $sPath

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now save the new Writer Document to the desktop folder.")

	$sFilePathName = _TempFile(@DesktopDir & "\", "TestExportDoc_", ".odt")

	;save The New Blank Doc To Desktop Directory using a unique temporary name.
	$sPath = _LOWriter_DocSaveAs($oDoc, $sFilePathName)
	If (@error > 0) Then _ERROR("Failed to Save the Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have created and saved the document to your Desktop, found at the following Path: " _
			 & $sPath & @CRLF & "Press Ok to write some data to it and then save the changes and close the document.")

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve View Cursor for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "This is some text to test saving a document.")
	If (@error > 0) Then _ERROR("Failed to insert a String into the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Save the changes to the document.
	_LOWriter_DocSave($oDoc)
	If (@error > 0) Then _ERROR("Failed to Save the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have written and saved the data, and closed the document. I will now open it up again to show it worked.")

	;Open the document.
	$oDoc = _LOWriter_DocOpen($sPath)
	If (@error > 0) Then _ERROR("Failed to open Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Document was succesfully opened. Press OK to close and delete it.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Following Error codes returned: Error:" & _
			@error & " Extended:" & @extended)

	;Delete the file.
	FileDelete($sPath)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



