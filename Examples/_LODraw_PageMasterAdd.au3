#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc
	Local $sMasterPages = ""
	Local $asMasterPages

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all Master pages in the Document.
	$asMasterPages = _LODraw_PageMastersGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master Page names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Cycle through and list all the master pages.
	For $i = 0 To @extended - 1
		$sMasterPages &= $asMasterPages[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "This Document contains the following Master Pages:" & @CRLF & $sMasterPages & @CRLF & @CRLF & _
			"Press ok to add a master page.")

	; Add a new Master Page
	_LODraw_PageMasterAdd($oDoc, Null, "Au3 Master")
	If @error Then _ERROR($oDoc, "Failed to Insert a new Master page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all Master pages in the Document.
	$asMasterPages = _LODraw_PageMastersGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master Page names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sMasterPages = ""

	; Cycle through and list all the master pages.
	For $i = 0 To @extended - 1
		$sMasterPages &= $asMasterPages[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "This Document now contains the following Master Pages:" & @CRLF & $sMasterPages)

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
