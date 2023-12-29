
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc, $oResult
	Local $sResultString
	Local $aAnEmptyArray[0] ;Create an empty array to skip FindFormat parameter.
	Local $avSim

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, search." & @CR & "A New Line.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Create a search descriptor with all set to False.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc)
	If (@error > 0) Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	;modify the Search descriptor to set the similarity settings to: Similarity = True, Combine = True, Number of characters to remove = 1,
	;Number of characters to add, = 1, Number of characters to exchange = 2.
	_LOWriter_SearchDescriptorSimilarityModify($oSrchDesc, True, True, 1, 1, 2)
	If (@error > 0) Then _ERROR("Failed to modify the search descriptor. Error:" & @error & " Extended:" & @extended)

	;Search for the word "Szarzhin" using the search descriptor I just created.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "Szarzhin", $aAnEmptyArray)
	If (@error > 0) Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	$sResultString = ""

	;Retrieve the Result's string.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult)
		If (@error > 0) Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		MsgBox($MB_OK, "", "The search was successful, I searched using similarity, and found the following word: " & $sResultString)
	Else
		MsgBox($MB_OK, "", "The search was successful, but returned no results.")
	EndIf

	;Retrieve the current Similarity settings.
	$avSim = _LOWriter_SearchDescriptorSimilarityModify($oSrchDesc)
	If (@error > 0) Then _ERROR("Failed to retrieve settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Search Descriptor's Similarity settings are as follows: " & @CRLF & _
			"Search using Similarity? True/False: " & $avSim[0] & @CRLF & _
			"Combine Similarity values together? True/False: " & $avSim[1] & @CRLF & _
			"The max number of characters that can be removed from the search term is: " & $avSim[2] & @CRLF & _
			"The max number of characters that can be added to the search term is: " & $avSim[3] & @CRLF & _
			"The max number of characters that can be exchanged in the search term is: " & $avSim[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

