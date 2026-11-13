#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oTextBox, $oPage, $oTextCursor
	Local $sFilePathName, $sPath

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now save the new Draw Document to the desktop folder.")

	$sFilePathName = _TempFile(@DesktopDir & "\", "TestExportDoc_", ".odg")

	; Save The New Blank Doc To Desktop Directory using a unique temporary name.
	$sPath = _LODraw_DocSaveAs($oDoc, $sFilePathName)
	If @error Then _ERROR($oDoc, "Failed to Save the Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created and saved the document to your Desktop, found at the following Path: " _
			& $sPath & @CRLF & "Press Ok to write some data to it and then save the changes and close the document.")

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Textbox.
	$oTextBox = _LODraw_ShapeTextBoxInsert($oPage, 10500, 5000, -1, 1500)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the Textbox.
	$oTextCursor = _LODraw_ShapeCreateTextCursor($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LODraw_CursorInsertString($oTextCursor, "Hi!")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set page background.
	_LODraw_PageBackColor($oPage, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
	If @error Then _ERROR($oDoc, "Failed to set Page background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save the changes to the document.
	_LODraw_DocSave($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Save the Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have written and saved the data, and closed the document. I will now open it up again to show it worked.")

	; Open the document.
	$oDoc = _LODraw_DocOpen($sPath)
	If @error Then _ERROR($oDoc, "Failed to open Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the file.
	FileDelete($sPath)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
