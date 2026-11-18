#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oMaster
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Master Page Object.
	$oMaster = _LODraw_PageMasterGetObjByIndex($oDoc, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master page Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Master Page Background Color settings. Background color = $LO_COLOR_TEAL.
	_LODraw_PageMasterBackColor($oMaster, $LO_COLOR_TEAL)
	If @error Then _ERROR($oDoc, "Failed to set Master Page settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to set the Master page's background transparency gradient level.")

	; Modify the Master Page Transparency Gradient settings to: Gradient Type = $LOD_GRAD_TYPE_ELLIPTICAL, XCenter to 75%, YCenter to 45%, Angle to 180 degrees
	; Border to 16%, Start transparency to 10%, End Transparency to 62%
	_LODraw_PageMasterBackTransparencyGradient($oMaster, $LOD_GRAD_TYPE_ELLIPTICAL, 75, 45, 180, 16, 10, 62)
	If @error Then _ERROR($oDoc, "Failed to set Master Page settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Master Page settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_PageMasterBackTransparencyGradient($oMaster)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master Page settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Master Page's Transparency Gradient settings are as follows: " & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avSettings[1] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avSettings[2] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avSettings[3] & @CRLF & _
			"The percentage of area not covered by the transparency is: " & $avSettings[4] & @CRLF & _
			"The starting transparency percentage is: " & $avSettings[5] & @CRLF & _
			"The ending transparency percentage is: " & $avSettings[6])

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
