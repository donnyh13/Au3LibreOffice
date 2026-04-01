#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOCalc_PageStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Sheet Printing inclusion settings to: Print Headings? = False, Print a Grid? = True, Print comments? True, Print Images/Objects? False,
	; Print Charts? True, Print Drawings? True, Print Formulas instead of results? True, Print Zero Values? False.
	_LOCalc_PageStyleSheetPrint($oPageStyle, False, True, True, False, True, True, True, False)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOCalc_PageStyleSheetPrint($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Sheet printing settings are as follows: " & @CRLF & _
			"Will Row and Column Headings be included when the sheet is printed? True/False: " & $avPageStyleSettings[0] & @CRLF & _
			"Will a Grid be included around each cell when the sheet is printed? True/False: " & $avPageStyleSettings[1] & @CRLF & _
			"Will Comments be included when the sheet is printed? True/False: " & $avPageStyleSettings[2] & @CRLF & _
			"Will Images and Objects be included when the sheet is printed? True/False: " & $avPageStyleSettings[3] & @CRLF & _
			"Will Charts be included when the sheet is printed? True/False: " & $avPageStyleSettings[4] & @CRLF & _
			"Will Drawings be included when the sheet is printed? True/False: " & $avPageStyleSettings[5] & @CRLF & _
			"Will Formulas be shown instead of results when the sheet is printed? True/False: " & $avPageStyleSettings[6] & @CRLF & _
			"Will values that equal 0 be included when this sheet is printed? True/False: " & $avPageStyleSettings[7])

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
