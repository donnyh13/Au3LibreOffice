#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oMaster, $oDefaultMaster
	Local $bIsInUse

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Object for the default Master Slide
	$oDefaultMaster = _LOImpress_SlideMasterGetObjByIndex($oDoc, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master slide object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a new Master Slide
	$oMaster = _LOImpress_SlideMasterAdd($oDoc, Null, "Au3 Master")
	If @error Then _ERROR($oDoc, "Failed to Insert a new Master slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$bIsInUse = _LOImpress_SlideMasterIsInUse($oDefaultMaster)
	If @error Then _ERROR($oDoc, "Failed to check if a master slide is in use. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the default Master slide currently in use? True/False: " & $bIsInUse)

	$bIsInUse = _LOImpress_SlideMasterIsInUse($oMaster)
	If @error Then _ERROR($oDoc, "Failed to check if a master slide is in use. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the new Master slide I added currently in use? True/False: " & $bIsInUse)

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
