
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc, $oResult, $oSrchDesc2
	Local $sResultString
	Local $aAnEmptyArray[0] ;Create an empty array to skip FindFormat parameter.

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, search.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Create a search descriptor for searching with. Set Backward to True.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc, True)
	If (@error > 0) Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	;Search for the word "Search" using the search descriptor I just created.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "Search", $aAnEmptyArray)
	If (@error > 0) Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	$sResultString = ""

	;Retrieve the Result's string.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult)
		If (@error > 0) Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		MsgBox($MB_OK, "", "The search was successful, I searched backwards, and found the following word: " & $sResultString)
	Else
		MsgBox($MB_OK, "", "The search was successful, but returned no results.")
	EndIf

	;Create a new search descriptor for searching with. Set Backward to False, and Match Case to True. I could  have just modified
	;my first search descriptor using the modify function, but since I am demonstrating the Search Descriptor Creation function, I will just make a new one.
	$oSrchDesc2 = _LOWriter_SearchDescriptorCreate($oDoc, False, True)
	If (@error > 0) Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	;Search for the word "SearcH" using the search descriptor I just created.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc2, "SearcH", $aAnEmptyArray)
	If (@error > 0) Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Result's string.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult)
		If (@error > 0) Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		MsgBox($MB_OK, "", "The search was successful, I searched, matching case, and found the following word: " & $sResultString)
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



