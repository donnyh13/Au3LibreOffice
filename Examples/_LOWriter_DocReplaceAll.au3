#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc
	Local $atFindFormat[0], $atReplaceFormat[0] ;Create two Empty Arrays to fill.

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search as an example." & @CR & "A New Line to SEARCH.")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Move the cursor right 35 spaces, selecting as I go.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 35, True)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Set the selected text's Font weight to (Bold) $LOW_WEIGHT_BOLD
	_LOWriter_DirFrmtFont($oDoc, $oViewCursor, Null, Null, Null, $LOW_WEIGHT_BOLD)
	If @error Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	; Create a search descriptor for searching with. Set Backward, Match Case, Whole word, Regular Expression, and Search Styles to false, and
	; Search property values to True.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc, False, False, False, False, False, True)
	If @error Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	; Create a Find Format Search Array for Bold font.
	_LOWriter_FindFormatModifyFont($oDoc, $atFindFormat, Null, Null, $LOW_WEIGHT_BOLD)
	If @error Then _ERROR("Failed to modify a Find format array. Error:" & @error & " Extended:" & @extended)

	; Create a Replace Format Search Array for Italic font.
	_LOWriter_FindFormatModifyFont($oDoc, $atReplaceFormat, Null, Null, Null, $LOW_POSTURE_ITALIC)
	If @error Then _ERROR("Failed to modify a Find format array. Error:" & @error & " Extended:" & @extended)

	; Search and replace all bold letter "a"'s with Italic "@".
	_LOWriter_DocReplaceAll($oDoc, $oSrchDesc, "a", "@", $atFindFormat, $atReplaceFormat)
	If @error Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	ConsoleWrite(@extended & @CRLF)

	MsgBox($MB_OK, "", "The search was successful, I searched using a Find Format, looking for any bold ""a""'s, " & _
			"and replaced all of them with an italic ""@"", I replaced " & @extended & " results.")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
