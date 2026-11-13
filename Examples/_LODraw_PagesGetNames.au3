#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc
	Local $sPages = ""
	Local $asPages

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add 5 Pages
	For $i = 1 To 5
		; Insert a new page.
		_LODraw_PageAdd($oDoc)
		If @error Then _ERROR($oDoc, "Failed to Insert a new page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Retrieve an Array of all pages in the Document.
	$asPages = _LODraw_PagesGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Cycle through and list all the pages.
	For $i = 0 To @extended - 1
		$sPages &= $asPages[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "This Document contains the following Pages:" & @CRLF & $sPages)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
