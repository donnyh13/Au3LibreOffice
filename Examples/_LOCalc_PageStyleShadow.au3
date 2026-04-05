#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iHMM
	Local $avPageStyleSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOCalc_PageStyleGetObjByName($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/8" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(.125, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Shadow settings to: Location = $LOC_SHADOW_LOCATION_TOP_LEFT, Color = $LO_COLOR_RED, Width = 1/8"
	_LOCalc_PageStyleShadow($oPageStyle, $LOC_SHADOW_LOCATION_TOP_LEFT, $LO_COLOR_RED, $iHMM)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOCalc_PageStyleShadow($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Shadow settings are as follows: " & @CRLF & _
			"The Shadow location is, (see UDF Constants): " & $avPageStyleSettings[0] & @CRLF & _
			"The Shadow color is (as a RGB Color Integer): " & $avPageStyleSettings[1] & @CRLF & _
			"The shadow width is, in Hundredths of a Millimeter (HMM): " & $avPageStyleSettings[2])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
