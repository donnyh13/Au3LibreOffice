#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc
	Local $sResultString
	Local $atFindFormat[0] ; Create an Empty Array to fill.
	Local $aoResults

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_ViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text, to use for searching later.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Some text to Search." & @CR & "A New Line to SEARCH.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor Down once
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_Down, 1)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the paragraph to Automatic hyphenation.
	_LOWriter_DirFrmtParHyphenation($oViewCursor, True)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a search descriptor for searching with. Set Backward, Match Case, Whole word, Regular Expression, and Search Styles to false, and
	; Search property values to True.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc, False, False, False, False, False, True)
	If @error Then _ERROR($oDoc, "Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Find Format Search Array for Auto hyphenation.
	_LOWriter_FindFormatModifyHyphenation($atFindFormat, True)
	If @error Then _ERROR($oDoc, "Failed to modify a Find format array. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "As of May 27th 2023, searching for any of these attribute will fail. If you watch closely, once you click ok, and I perform the " & _
			"search, all instances of ""Search"" will disappear. LibreOffice is aware of the bug, but haven't decided how to fix it yet.")

	; Search for the word "search".
	$aoResults = _LOWriter_DocFindAll($oDoc, $oSrchDesc, "search", $atFindFormat)
	If @error Then _ERROR($oDoc, "Failed to perform search in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sResultString = ""

	; Retrieve the Result's string.
	If IsArray($aoResults) Then
		For $i = 0 To UBound($aoResults) - 1
			$sResultString = _LOWriter_CursorGetString($aoResults[$i])
			If @error Then MsgBox($MB_OK + $MB_TOPMOST, Default, "Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		Next

		MsgBox($MB_OK + $MB_TOPMOST, Default, "The search was successful, I searched using a Find Format, looking for any paragraphs that are hyphenated, " & _
				"containing the word ""Search"", and found the following: " & $sResultString)
	EndIf

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
