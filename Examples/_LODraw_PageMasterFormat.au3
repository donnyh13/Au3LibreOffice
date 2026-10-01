#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oMaster
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Master page for the current Page.
	$oMaster = _LODraw_PageMasterCurrent($oPage)
	If @error Then _ERROR($oDoc, "Failed to retrieve current master page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set some of the Master Page's Format settings
	_LODraw_PageMasterFormat($oMaster, $LOD_PAGE_WIDTH_DIA_SLIDE, $LOD_PAGE_HEIGHT_DIA_SLIDE, $LOD_PAGE_ORIENT_LANDSCAPE)
	If @error Then _ERROR($oDoc, "Failed to set Page's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Master Page's Format settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_PageMasterFormat($oMaster)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Master Page's format settings are as follows: " & @CRLF & _
			"The page's width is, in Hundredths of a Millimeter (HMM): " & $avSettings[0] & @CRLF & _
			"The page's height is, in Hundredths of a Millimeter (HMM): " & $avSettings[1] & @CRLF & _
			"The page's orientation is (See UDF Constants): " & $avSettings[2])

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
