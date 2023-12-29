#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc, $oResult
	Local $sResultString
	Local $aAnEmptyArray[0] ;Create an empty array to skip FindFormat parameter.
	Local $abSearch

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, seaRch." & @CR & "A New Line.")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Create a search descriptor with all set to False.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc)
	If @error Then _ERROR("Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended)

	; modify the Search descriptor to set Match Case to True, and Whole Words to True
	_LOWriter_SearchDescriptorModify($oSrchDesc, False, True, True)
	If @error Then _ERROR("Failed to modify the search descriptor. Error:" & @error & " Extended:" & @extended)

	; Search for the word "seaRch" using the search descriptor I just created.
	$oResult = _LOWriter_DocFindNext($oDoc, $oSrchDesc, "seaRch", $aAnEmptyArray)
	If @error Then _ERROR("Failed to perform search in the document. Error:" & @error & " Extended:" & @extended)

	$sResultString = ""

	; Retrieve the Result's string.
	If IsObj($oResult) Then
		$sResultString = _LOWriter_DocGetString($oResult)
		If @error Then _ERROR("Failed to retrieve String. Error:" & @error & " Extended:" & @extended)
		MsgBox($MB_OK, "", "The search was successful, I searched and found the following word: " & $sResultString)
	Else
		MsgBox($MB_OK, "", "The search was successful, but returned no results.")
	EndIf

	; Retrieve the current Search Descriptor settings.
	$abSearch = _LOWriter_SearchDescriptorModify($oSrchDesc)
	If @error Then _ERROR("Failed to retrieve settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Search Descriptor settings are as follows: " & @CRLF & _
			"Search backwards? True/False: " & $abSearch[0] & @CRLF & _
			"Search matching case? True/False: " & $abSearch[1] & @CRLF & _
			"Search for whole words only? True/False: " & $abSearch[2] & @CRLF & _
			"Search using Regular Expressions? True/False: " & $abSearch[3] & @CRLF & _
			"Search for Paragraph Styles or for format settings in Styles? True/False: " & $abSearch[3] & @CRLF & _
			"Search words only with specific paragrapgh format settings?: " & $abSearch[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
