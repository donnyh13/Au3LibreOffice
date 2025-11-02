#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings[0], $avOrigSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Document View Display Settings
	$avOrigSettings = _LOCalc_DocViewDisplaySettings($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Document Display settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Document display settings are: " & @CRLF & _
			"Are Formulas shown instead of the results? True/False: " & $avOrigSettings[0] & @CRLF & _
			"Are Zero Values shown? True/False: " & $avOrigSettings[1] & @CRLF & _
			"Is a indicator displayed indicate a Cell contains a comment? True/False: " & $avOrigSettings[2] & @CRLF & _
			"Are Page Breaks displayed? True/False: " & $avOrigSettings[3] & @CRLF & _
			"Are Help lines shown when dragging a Shape etc.? True/False: " & $avOrigSettings[4] & @CRLF & _
			"Are Cell contents colored differently depending on content type? True/False: " & $avOrigSettings[5] & @CRLF & _
			"Are Anchor icons displayed for Images, shapes etc.? True/False: " & $avOrigSettings[6] & @CRLF & _
			"Are Gridlines shown? True/False: " & $avOrigSettings[7] & @CRLF & _
			"The Gridline color is (in Long Integer Color format): " & $avOrigSettings[8])

	If $IDYES = MsgBox($MB_YESNO + $MB_ICONQUESTION + $MB_DEFBUTTON1, "", "Would you like a demonstration of modifying these settings?" & @CRLF & _
			"Warning, if the script fails to complete these settings may remain changed from your current setting values. Proceed with caution.") Then

		; Set Document Display settings, Set Display Formulas instead of Results to True, Display zero values = False, display comment indicator = False,
		; Display Page breaks = False, Display Helplines = True, Color Cell Contents differently = False, display anchors = True, Display a Grid = True,
		; Grid color = $LO_COLOR_GOLD.
		_LOCalc_DocViewDisplaySettings($oDoc, True, False, False, False, True, False, True, True, $LO_COLOR_GOLD)
		If @error Then _ERROR($oDoc, "Failed to set Document Display settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve the current Document View Display Settings
		$avSettings = _LOCalc_DocViewDisplaySettings($oDoc)
		If @error Then _ERROR($oDoc, "Failed to retrieve current Document Display settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Document display settings are: " & @CRLF & _
				"Are Formulas shown instead of the results? True/False: " & $avSettings[0] & @CRLF & _
				"Are Zero Values shown? True/False: " & $avSettings[1] & @CRLF & _
				"Is a indicator displayed indicate a Cell contains a comment? True/False: " & $avSettings[2] & @CRLF & _
				"Are Page Breaks displayed? True/False: " & $avSettings[3] & @CRLF & _
				"Are Help lines shown when dragging a Shape etc.? True/False: " & $avSettings[4] & @CRLF & _
				"Are Cell contents colored differently depending on content type? True/False: " & $avSettings[5] & @CRLF & _
				"Are Anchor icons displayed for Images, shapes etc.? True/False: " & $avSettings[6] & @CRLF & _
				"Are Grid lines shown? True/False: " & $avSettings[7] & @CRLF & _
				"The Grid line color is (in Long Integer Color format): " & $avSettings[8])

		; Return settings to User's old settings.
		_LOCalc_DocViewDisplaySettings($oDoc, $avOrigSettings[0], $avOrigSettings[1], $avOrigSettings[2], $avOrigSettings[3], $avOrigSettings[4], $avOrigSettings[5], _
				$avOrigSettings[6], $avOrigSettings[7], $avOrigSettings[8])
		If @error Then _ERROR($oDoc, "Failed to re-set Document Display settings to user's previous values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	EndIf

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
