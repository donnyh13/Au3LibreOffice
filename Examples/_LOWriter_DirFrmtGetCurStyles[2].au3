#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $aoParagraphs
	Local $asStyles

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "A Line of text to Test Direct Formatting with.")
	If @error Then _ERROR($oDoc, "Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Paragraphs for the document
	$aoParagraphs = _LOWriter_CursorParObjCreateList($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Paragraph Objects for Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current styles contained in the first Paragraph. Will return a 4 element Array in the following order: current paragraph style,
	; current Character style, current Page style, current Numbering Style (if one is active).
	$asStyles = _LOWriter_DirFrmtGetCurStyles($aoParagraphs[0])
	If @error Then _ERROR($oDoc, "Failed to retrieve current styles in the text selection. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "These styles are currently set for the Paragraph. Some may be blank, indicating no style is set for that style " & _
			"type: " & @CRLF & _
			"Paragraph Style: " & $asStyles[0] & @CRLF & _
			"Character Style: " & $asStyles[1] & @CRLF & _
			"Page Style: " & $asStyles[2] & @CRLF & _
			"Numbering Style: " & $asStyles[3])

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
