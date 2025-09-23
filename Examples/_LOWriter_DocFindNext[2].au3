#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc, $oFootNote, $oFootTextCursor, $oResult
	Local $sResultString

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, search." & @CR & "A New Line to searCh.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a search descriptor with all set to False.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc)
	If @error Then _ERROR($oDoc, "Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor right 44 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 44)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a footnote.
	$oFootNote = _LOWriter_FootnoteInsert($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a footnote. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Get a TextCursor for the footnote
	$oFootTextCursor = _LOWriter_FootnoteGetTextCursor($oFootNote)
	If @error Then _ERROR($oDoc, "Failed to create a footnote Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text in the footnote.
	_LOWriter_DocInsertString($oDoc, $oFootTextCursor, "Some text to in the footnote with the word SEarCH.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor right 29 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 29)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select all the words from that spot to the end of the line.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END_OF_LINE, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Search the word "search" within the ViewCursor selection.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "search", Null, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to perform search in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sResultString = ""

	; Retrieve the string for the result.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult) & @CRLF
		If @error Then _ERROR($oDoc, "Failed to retrieve String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	EndIf

	; Search for all matching results in this document, one at a time.
	While IsObj($oResult)
		; Search for the word "Search" using the search descriptor I just created. Starting from my last result.
		$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "Search", Null, $oViewCursor, $oResult)
		If @error Then _ERROR($oDoc, "Failed to perform search in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve the Result's string.
		If IsObj($oResult) Then
			$sResultString &= _LOWriter_DocGetString($oResult) & @CRLF
			If @error Then _ERROR($oDoc, "Failed to retrieve String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		EndIf
	WEnd

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The search was successful, I searched, and found the following words within the selection: " & $sResultString & @CRLF & @CRLF & _
			"Did you notice the search didn't find the word ""SEarCH"" in the footnote? I will search again, this time with Exhaustive set to True.")

	; Search the word "search" within the ViewCursor selection exhaustively.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "search", Null, $oViewCursor, Null, True)
	If @error Then _ERROR($oDoc, "Failed to perform search in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the string for the result.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult) & @CRLF
		If @error Then _ERROR($oDoc, "Failed to retrieve String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	EndIf

	; Search for all matching results in this document, one at a time.
	While IsObj($oResult)
		; Search for the word "Search" using the search descriptor I just created. Starting from my last result.
		$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "Search", Null, $oViewCursor, $oResult, True)
		If @error Then _ERROR($oDoc, "Failed to perform search in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve the Result's string.
		If IsObj($oResult) Then
			$sResultString &= _LOWriter_DocGetString($oResult) & @CRLF
			If @error Then _ERROR($oDoc, "Failed to retrieve String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		EndIf
	WEnd

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The search was successful, I searched, and found the following word within the selection: " & $sResultString)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
