
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc
	Local $sResultString
	Local $aAnEmptyArray[0] ;Create an empty array to skip FindFormat parameter.
	Local $aoResults

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, seaRch." & @CR & "A New Line.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Create a search descriptor with all set to False.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc)
	If (@error > 0) Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	;Search all the results containing the word "search".
	$aoResults = _LOWriter_DocFindAll($oDoc, $oSrchDesc, "search", $aAnEmptyArray)
	If (@error > 0) Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	$sResultString = ""

	;Retrieve the string for each result.
	If IsArray($aoResults) Then
		For $i = 0 To UBound($aoResults) - 1
			$sResultString &= _LOWriter_DocGetString($aoResults[$i]) & @CRLF
		Next
		If (@error > 0) Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		MsgBox($MB_OK, "", "The search was successful, I searched and found the following words: " & @CRLF & $sResultString)
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



