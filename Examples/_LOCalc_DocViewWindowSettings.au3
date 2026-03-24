#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings[0], $avOrigSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Document View Window Settings
	$avOrigSettings = _LOCalc_DocViewWindowSettings($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Document Window View settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Document Window view settings are: " & @CRLF & _
			"Are Column/Row headers shown? True/False: " & $avOrigSettings[0] & @CRLF & _
			"Is the Vertical Scroll bar visible? True/False: " & $avOrigSettings[1] & @CRLF & _
			"Is the Horizontal scroll bar visible? True/False: " & $avOrigSettings[2] & @CRLF & _
			"Is the Sheet Tabs bar displayed? True/False: " & $avOrigSettings[3] & @CRLF & _
			"Are Outline symbols displayed? True/False: " & $avOrigSettings[4] & @CRLF & _
			"Are Charts displayed? True/False: " & $avOrigSettings[5] & @CRLF & _
			"Are Drawings displayed? True/False: " & $avOrigSettings[6] & @CRLF & _
			"Are Objects/Graphics visible? True/False: " & $avOrigSettings[7])

	If $IDYES = MsgBox($MB_YESNO + $MB_ICONQUESTION + $MB_DEFBUTTON1, "", "Would you like a demonstration of modifying these settings?" & @CRLF & _
			"Warning, if the script fails to complete these settings may remain changed from your current setting values. Proceed with caution.") Then

		; Set Document Window settings, Set Display Headers = False, Display vertical Scroll bar = False, display horizontal scroll bar = False,
		; Display Sheet Tabs = False, Display Outline Symbols = True, Display Charts = False, display Drawings = False, Display Graphics/Objects = True
		_LOCalc_DocViewWindowSettings($oDoc, False, False, False, False, True, False, True, True)
		If @error Then _ERROR($oDoc, "Failed to set Document Window settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve the current Document View Window Settings
		$avSettings = _LOCalc_DocViewWindowSettings($oDoc)
		If @error Then _ERROR($oDoc, "Failed to retrieve current Document Window settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Document Window view settings are: " & @CRLF & _
				"Are Column/Row headers shown? True/False: " & $avSettings[0] & @CRLF & _
				"Is the Vertical Scroll bar visible? True/False: " & $avSettings[1] & @CRLF & _
				"Is the Horizontal scroll bar visible? True/False: " & $avSettings[2] & @CRLF & _
				"Is the Sheet Tabs bar displayed? True/False: " & $avSettings[3] & @CRLF & _
				"Are Outline symbols displayed? True/False: " & $avSettings[4] & @CRLF & _
				"Are Charts displayed? True/False: " & $avSettings[5] & @CRLF & _
				"Are Drawings displayed? True/False: " & $avSettings[6] & @CRLF & _
				"Are Objects/Graphics visible? True/False: " & $avSettings[7])

		; Return settings to User's old settings.
		_LOCalc_DocViewWindowSettings($oDoc, $avOrigSettings[0], $avOrigSettings[1], $avOrigSettings[2], $avOrigSettings[3], $avOrigSettings[4], $avOrigSettings[5], _
				$avOrigSettings[6], $avOrigSettings[7])
		If @error Then _ERROR($oDoc, "Failed to re-set Document Window settings to user's previous values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	EndIf

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
