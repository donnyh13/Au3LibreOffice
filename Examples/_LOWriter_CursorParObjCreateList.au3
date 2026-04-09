#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $aoParagraphs

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "First Line of text" & @CR & _
			"Second line of text." & @CR & _
			"Third line of text." & @CR & _
			"Fourth Line of Text.")
	If @error Then _ERROR($oDoc, "Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Paragraphs for the document
	$aoParagraphs = _LOWriter_CursorParObjCreateList($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Paragraph Objects for Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & @extended & " paragraph objects returned." & _
			" As an example of what a paragraph Object can be used for, I will change the font size of the first paragraph to 22 point.")

	; An example of what I can do with a paragraph Object. Set the first paragraph's font size to 22 point.
	_LOWriter_DirFrmtFont($aoParagraphs[0], Null, 22)
	If @error Then _ERROR($oDoc, "Failed to direct format Paragraph Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
