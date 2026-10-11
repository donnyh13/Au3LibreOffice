#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oMaster, $oPage, $oCurrMaster

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a new Master Page
	$oMaster = _LODraw_PageMasterAdd($oDoc, Null, "Au3 Master")
	If @error Then _ERROR($oDoc, "Failed to Insert a new Master page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the background color of the new Master page to purple.
	_LODraw_PageMasterBackColor($oMaster, $LO_COLOR_PURPLE)
	If @error Then _ERROR($oDoc, "Failed to set background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to set the current page's Master page to the new Master page I created.")

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Master page of the current page to the new Master page.
	_LODraw_PageMasterCurrent($oPage, $oMaster)
	If @error Then _ERROR($oDoc, "Failed to set Master page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current master page of the current page.
	$oCurrMaster = _LODraw_PageMasterCurrent($oPage)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Master page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Master Page of the current page is set to: " & _LODraw_PageMasterName($oCurrMaster) & ".")

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
