#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oHandout
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the handout Page's Object.
	$oHandout = _LOImpress_SlideHandoutGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Handout page's Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set some of the Handout Page's margin settings
	_LOImpress_SlideHandoutMargins($oHandout, 1270, 2540, 2540, 1270)
	If @error Then _ERROR($oDoc, "Failed to set Page's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Handout Page's margin settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_SlideHandoutMargins($oHandout)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Handout Page's margin settings are as follows: " & @CRLF & _
			"The the left page margin is, in Hundredths of a Millimeter (HMM): " & $avSettings[0] & @CRLF & _
			"The the right page margin is, in Hundredths of a Millimeter (HMM): " & $avSettings[1] & @CRLF & _
			"The the top page margin is, in Hundredths of a Millimeter (HMM): " & $avSettings[2] & @CRLF & _
			"The the bottom page margin is, in Hundredths of a Millimeter (HMM): " & $avSettings[3])

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
