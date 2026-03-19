
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc, $oResult
	Local $sResultString
	Local $aAnEmptyArray[0] ;Create an empty array to skip FindFormat parameter.
	Local $atFindFormat[0] ;Create an Empty Array to fill.
	Local $aoResults

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, search." & @CR & "A New Line.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Create a search descriptor for searching with. All set to False
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc)
	If (@error > 0) Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	;Search for the word "Search using the search descriptor I just created.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "Search", $aAnEmptyArray)
	If (@error > 0) Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	$sResultString = ""

	;Retrieve the Result's string.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult)
		If (@error > 0) Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		MsgBox($MB_OK, "", "The search was successful, I searched for the word ""Search"", and found the following word: " & $sResultString)
	Else
		MsgBox($MB_OK, "", "The search was successful, but returned no results.")
	EndIf

	;Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the cursor right 29 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 29)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Select the word "SEArch".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 6, True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Set the selected text's Font to  and weight (Bold) to $LOW_WEIGHT_BOLD
	_LOWriter_DirFrmtFont($oDoc, $oViewCursor, Null, Null, Null, $LOW_WEIGHT_BOLD)
	If (@error > 0) Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	;Move the cursor right 10 spaces, without selecting
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 10, False)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Select the word "SearcHing".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 9, True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Set the selected text's Font weight to $LOW_WEIGHT_SEMI_BOLD
	_LOWriter_DirFrmtFont($oDoc, $oViewCursor, Null, Null, Null, $LOW_WEIGHT_SEMI_BOLD)
	If (@error > 0) Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	;Create a Find Format Search Array for Bold font.
	_LOWriter_FindFormatModifyFont($oDoc, $atFindFormat, Null, Null, $LOW_WEIGHT_BOLD)
	If (@error > 0) Then _ERROR("Failed to modify a Find format array. Error:" & @error & " Extended:" & @extended)

	;Search for the word "search", I am using a FindFormat Array with the Bold attribute, however I still have $bSearchPropValues set to false,
	;which means instead of searching for formatting that is bold, it will search for any direct formatting involving font weight, so those
	;two words I directly formatted above, one to Bold weight the other to Semi-Bold weight, will both be found.
	$aoResults = _LOWriter_DocFindAll($oDoc, $oSrchDesc, "search", $atFindFormat)
	If (@error > 0) Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	;Clear my result String
	$sResultString = ""

	;Retrieve the Result's string.
	If IsArray($aoResults) Then
		For $i = 0 To UBound($aoResults) - 1
			$sResultString &= _LOWriter_DocGetString($aoResults[$i]) & @CRLF
			If (@error > 0) Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		Next
	EndIf

	MsgBox($MB_OK, "", "The search was successful, I searched using a Find Format, looking for any words with a modified weight, and found the following: " & $sResultString)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



