#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oMaster
	Local $sName

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the first Master Slide's Object.
	$oMaster = _LOImpress_SlideMasterGetObjByIndex($oDoc, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master Slide Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current name of the Master slide.
	$sName = _LOImpress_SlideMasterName($oMaster)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master slide name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Master Slide's name is currently: " & $sName & @CRLF & @CRLF & "Press ok to rename it to ""Au3 Master"".")

	; Rename the Master slide.
	$sName = _LOImpress_SlideMasterName($oMaster, "Au3 master")
	If @error Then _ERROR($oDoc, "Failed to rename Master slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the name of the Master slide again.
	$sName = _LOImpress_SlideMasterName($oMaster)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master slide name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Master Slide's name now is: " & $sName)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
