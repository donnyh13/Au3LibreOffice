#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $aiCounts

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Some text." & @CR & @CR & "Some different text" & @CR & "Another Line.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of statistical counts of the following, in this order: Page count; Line Count; Paragraph Count;
	; Word Count; Character Count; Non-WhiteSpace Character Count; Table Count; Image Count; Object Count.
	$aiCounts = _LOWriter_DocGetCounts($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document counts. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The document counts are as follows: " & @CRLF & _
			"Number of Pages: " & $aiCounts[0] & @CRLF & _
			"Number of Lines: " & $aiCounts[1] & @CRLF & _
			"Number of Paragraphs: " & $aiCounts[2] & @CRLF & _
			"Number of Words: " & $aiCounts[3] & @CRLF & _
			"Number of Characters: " & $aiCounts[4] & @CRLF & _
			"Number of Characters, not counting white-spaces: " & $aiCounts[5] & @CRLF & _
			"Number of Tables: " & $aiCounts[6] & @CRLF & _
			"Number of Images/Graphics: " & $aiCounts[7] & @CRLF & _
			"Number of Objects: " & $aiCounts[8])

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
