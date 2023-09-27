#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc, $oResult
	Local $sResultString
	Local $atFindFormat[0] ;Create an Empty Array to fill.

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, search." & @CR & "A New Line.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the cursor right 29 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 29)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Select the word "SEArch".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 6, True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Strike out the selected text
	_LOWriter_DirFrmtStrikeOut($oViewCursor, Null, True)
	If (@error > 0) Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	;Create a search descriptor for searching with. Set Backward, Match Case, Whole word, Regular Expression, and Search Styles to false, and
	;Search property values to True.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc, False, False, False, False, False, True)
	If (@error > 0) Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	;Create a Find Format Search Array for any struck out text
	_LOWriter_FindFormatModifyStrikeout($atFindFormat, Null, True)
	If (@error > 0) Then _ERROR("Failed to modify a Find format array. Error:" & @error & " Extended:" & @extended)

	;Search for the word "search" that is struck out.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "search", $atFindFormat)
	If (@error > 0) Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	$sResultString = ""

	;Retrieve the Result's string.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult)
		If (@error > 0) Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		MsgBox($MB_OK, "", "The search was successful, I searched using a Find Format, looking for a struck out word, and found the following: " & $sResultString)
	Else
		MsgBox($MB_OK, "", "The search was successful, but returned no results.")
	EndIf

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
