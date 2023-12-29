
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc
	Local $sResultString
	Local $atFindFormat[0] ;Create an Empty Array to fill.
	Local $aoResults

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search." & @CR & "A New Line to SEARCH.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor Down once
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_Down, 1)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Set the paragraph to Automatic hyphenation.
	_LOWriter_DirFrmtParHyphenation($oViewCursor, True)
	If (@error > 0) Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	;Create a search descriptor for searching with. Set Backward, Match Case, Whole word, Regular Expression, and Search Styles to false, and
	;Search property values to True.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc, False, False, False, False, False, True)
	If (@error > 0) Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	;Create a Find Format Search Array for Auto hyphenation.
	_LOWriter_FindFormatModifyHyphenation($atFindFormat, True)
	If (@error > 0) Then _ERROR("Failed to modify a Find format array. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "As of May 27th 2023, searching for any of these attribute will fail. If you watch closely, once you click ok, and I perform the " & _
			"search, all instances of ""Search"" will disappear. Libre Office is aware of the bug, but haven't decided how to fix it yet.")

	;Search for the word "search".
	$aoResults = _LOWriter_DocFindAll($oDoc, $oSrchDesc, "search", $atFindFormat)
	If (@error > 0) Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	$sResultString = ""

	;Retrieve the Result's string.
	If IsArray($aoResults) Then
		For $i = 0 To UBound($aoResults) - 1
			$sResultString = _LOWriter_DocGetString($aoResults[$i])
			If (@error > 0) Then MsgBox($MB_OK, "", "Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		Next

		MsgBox($MB_OK, "", "The search was successful, I searched using a Find Format, looking for any paragraphs that are hyphenated, " & _
				"containing the word ""Search"", and found the following: " & $sResultString)

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



