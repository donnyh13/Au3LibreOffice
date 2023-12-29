#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc, $oFootNote, $oFootTextCursor, $oResult
	Local $sResultString
	Local $aAnEmptyArray[0] ;Create an empty array to skip FindFormat parameter.

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, search." & @CR & "A New Line to searCh.")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Create a search descriptor with all set to False.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc)
	If @error Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Move the cursor right 44 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 44)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Insert a footnote.
	$oFootNote = _LOWriter_FootnoteInsert($oDoc, $oViewCursor)
	If @error Then _ERROR("Failed to insert a footnote. Error:" & @error & " Extended:" & @extended)

	; Get a TextCursor for the footnote
	$oFootTextCursor = _LOWriter_FootnoteGetTextCursor($oFootNote)
	If @error Then _ERROR("Failed to create a footnote Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Insert some text in the footnote.
	_LOWriter_DocInsertString($oDoc, $oFootTextCursor, "Some text to in the footnote with the word SEarCH.")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Move the cursor right 29 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 29)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Select all the words from that spot to the end of the line.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END_OF_LINE, 1, True)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Search the word "search" within the viewcursor selection.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "search", $aAnEmptyArray, $oViewCursor)
	If @error Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	$sResultString = ""

	; Retrieve the string for the result.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult) & @CRLF
		If @error Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
	EndIf

	; Search for all matching results in this document, one at a time.
	While IsObj($oResult)

		; Search for the word "Search" using the search descriptor I just created. Starting from my last result.
		$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "Search", $aAnEmptyArray, $oViewCursor, $oResult)
		If @error Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

		; Retrieve the Result's string.
		If IsObj($oResult) Then
			$sResultString &= _LOWriter_DocGetString($oResult) & @CRLF
			If @error Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		EndIf

	WEnd

	MsgBox($MB_OK, "", "The search was successful, I searched, and found the following words within the selection: " & $sResultString & @CRLF & @CRLF & _
			"Did you notice the search didn't find the word ""SEarCH"" in the footnote? I will search again, this time with Exhaustive set to True.")

	; Search the word "search" within the viewcursor selection exhaustively.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "search", $aAnEmptyArray, $oViewCursor, Null, True)
	If @error Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the string for the result.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult) & @CRLF
		If @error Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
	EndIf

	; Search for all matching results in this document, one at a time.
	While IsObj($oResult)

		; Search for the word "Search" using the search descriptor I just created. Starting from my last result.
		$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "Search", $aAnEmptyArray, $oViewCursor, $oResult, True)
		If @error Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

		; Retrieve the Result's string.
		If IsObj($oResult) Then
			$sResultString &= _LOWriter_DocGetString($oResult) & @CRLF
			If @error Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		EndIf

	WEnd

	MsgBox($MB_OK, "", "The search was successful, I searched, and found the following word within the selection: " & $sResultString)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
