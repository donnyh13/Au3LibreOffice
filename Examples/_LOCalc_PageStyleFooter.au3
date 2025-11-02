#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iMicrometers, $iMicrometers2
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOCalc_PageStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Border Width settings to: $LOC_BORDERWIDTH_MEDIUM on all four sides.
	_LOCalc_PageStyleBorderWidth($oPageStyle, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2" to Micrometers
	$iMicrometers2 = _LO_ConvertToMicrometer(.5)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Footer settings to: Footer on = True, Same content on left and right pages = False, Same content on the first page = True,
	; Left & Right margins = 1/4", Spacing between Footer content and Page content = 1/2", Skip Height and set AutoHeight to True.
	_LOCalc_PageStyleFooter($oPageStyle, True, False, True, $iMicrometers, $iMicrometers, $iMicrometers2, Null, True)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOCalc_PageStyleFooter($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Footer settings are as follows: " & @CRLF & _
			"Is the Footer on for this Page Style? True/False: " & $avPageStyleSettings[0] & @CRLF & _
			"Is the content on Left and Right pages the same? True/False: " & $avPageStyleSettings[1] & @CRLF & _
			"Is the content on the first page the same? True/False: " & $avPageStyleSettings[2] & @CRLF & _
			"The Left Margin Width is, in Micrometers: " & $avPageStyleSettings[3] & @CRLF & _
			"The Right Margin Width is, in Micrometers: " & $avPageStyleSettings[4] & @CRLF & _
			"The Spacing between the Footer contents and the Page contents, in Micrometers: " & $avPageStyleSettings[5] & @CRLF & _
			"The height of the Footer, in Micrometers: " & $avPageStyleSettings[6] & @CRLF & _
			"Is the height of the Footer automatically adjusted? True/False: " & $avPageStyleSettings[7])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
