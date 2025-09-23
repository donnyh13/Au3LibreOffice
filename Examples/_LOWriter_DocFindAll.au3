#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSrchDesc
	Local $sResultString
	Local $aoResults

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text, to use for searching later.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to Search, SeArCh, SEArch, SEARCH, SearcHing, seaRched, seaRch." & @CR & "A New Line.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a search descriptor with all set to False.
	$oSrchDesc = _LOWriter_SearchDescriptorCreate($oDoc)
	If @error Then _ERROR($oDoc, "Failed to create a search descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Search all the results containing the word "search".
	$aoResults = _LOWriter_DocFindAll($oDoc, $oSrchDesc, "search")
	If @error Then _ERROR($oDoc, "Failed to perform search in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sResultString = ""

	; Retrieve the string for each result.
	If IsArray($aoResults) Then
		For $i = 0 To UBound($aoResults) - 1
			$sResultString &= _LOWriter_DocGetString($aoResults[$i]) & @CRLF
		Next
		If @error Then _ERROR($oDoc, "Failed to retrieve String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The search was successful, I searched and found the following words: " & @CRLF & $sResultString)

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
