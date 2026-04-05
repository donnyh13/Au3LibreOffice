#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $sFilePathName, $sPath

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now save the new Writer Document to the desktop folder.")

	$sFilePathName = _TempFile(@DesktopDir & "\", "TestExportDoc_", ".odt")

	; save The New Blank Doc To Desktop Directory using a unique temporary name.
	$sPath = _LOWriter_DocSaveAs($oDoc, $sFilePathName)
	If @error Then _ERROR($oDoc, "Failed to Save the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created and saved the document to your Desktop, found at the following Path: " _
			& $sPath & @CRLF & "Press Ok to write some data to it and then save the changes and close the document.")

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_ViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve View Cursor for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "This is some text to test saving a document.")
	If @error Then _ERROR($oDoc, "Failed to insert a String into the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save the changes to the document.
	_LOWriter_DocSave($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Save the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have written and saved the data, and closed the document. I will now open it up again to show it worked.")

	; Open the document.
	$oDoc = _LOWriter_DocOpen($sPath)
	If @error Then _ERROR($oDoc, "Failed to open Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Document was successfully opened. Press OK to close and delete it.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the file.
	FileDelete($sPath)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
