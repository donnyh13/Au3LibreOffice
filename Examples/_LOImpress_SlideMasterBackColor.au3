#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oMaster
	Local $iColor

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Master Slide Object.
	$oMaster = _LOImpress_SlideMasterGetObjByIndex($oDoc, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master slide Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Slide's Background Color settings. Background color = $LO_COLOR_GOLD.
	_LOImpress_SlideMasterBackColor($oMaster, $LO_COLOR_GOLD)
	If @error Then _ERROR($oDoc, "Failed to set Master Slide settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current rectangle Shape settings. Return will be an Integer.
	$iColor = _LOImpress_SlideMasterBackColor($oMaster)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master Slide settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Master Slide's Background color settings are as follows: " & @CRLF & _
			"The Background color is (as a RGB Color Integer): " & $iColor)

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
