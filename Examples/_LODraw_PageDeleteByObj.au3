#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current page Object.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Current Page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set page background.
	_LODraw_PageBackColor($oPage, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
	If @error Then _ERROR($oDoc, "Failed to set Page background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new page.
	$oPage = _LODraw_PageAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Insert a new page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set page background.
	_LODraw_PageBackColor($oPage, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
	If @error Then _ERROR($oDoc, "Failed to set Page background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new page.
	$oPage = _LODraw_PageAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Insert a new page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set page background.
	_LODraw_PageBackColor($oPage, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
	If @error Then _ERROR($oDoc, "Failed to set Page background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to delete the middle page.")

	; Get the middle page's object
	$oPage = _LODraw_PageGetObjByIndex($oDoc, 1)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the middle page.
	_LODraw_PageDeleteByObj($oPage)
	If @error Then _ERROR($oDoc, "Failed to delete Page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
