#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc, $oResult, $oSrchDesc2
	Local $sResultString
	Local $atFindFormat[0] ; Create an Empty Array to fill.

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, search." & @CR & "A New Line.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the last paragraph to Paragraph Style Title using the View Cursor
	_LOWriter_ParStyleSet($oDoc, $oViewCursor, "Title")

	; Create a search descriptor for searching with. Set Backward, Match Case, Whole Word, and Regular Expressions to False, and Search Styles to True.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc, False, False, False, False, True)
	If @error Then _ERROR($oDoc, "Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Search for the Paragraph Style "Title" using the search descriptor I just created.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "Title")
	If @error Then _ERROR($oDoc, "Failed to perform search in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sResultString = ""

	; Retrieve the Result's string.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult)
		If @error Then _ERROR($oDoc, "Failed to retrieve String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The search was successful, I searched for the Paragraph Style ""Title"", and found the following word: " & $sResultString)

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The search was successful, but returned no results.")
	EndIf

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor right 29 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 29)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select the word "SEArch".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 6, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the selected text's Font weight to (Bold) $LOW_WEIGHT_BOLD
	_LOWriter_DirFrmtFont($oViewCursor, Null, Null, Null, $LOW_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new search descriptor for searching with. Set Backward, Match Case, Whole word, Regular Expression, and Search Styles to false, and
	; Search property values to True.
	; I could  have just modified my first search descriptor using the modify function, but since I am demonstrating the Search Descriptor Creation
	; function, I will just make a new one.
	$oSrchDesc2 = _LOWriter_SearchDescriptorCreate($oDoc, False, False, False, False, False, True)
	If @error Then _ERROR($oDoc, "Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Find Format Search Array for Bold font.
	_LOWriter_FindFormatModifyFont($atFindFormat, Null, Null, $LOW_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to modify a Find format array. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Search for the word "search" that is bold.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc2, "search", $atFindFormat)
	If @error Then _ERROR($oDoc, "Failed to perform search in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Result's string.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult)
		If @error Then _ERROR($oDoc, "Failed to retrieve String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The search was successful, I searched using a Find Format, looking for a bold word, and found the following: " & $sResultString)

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The search was successful, but returned no results.")
	EndIf

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
