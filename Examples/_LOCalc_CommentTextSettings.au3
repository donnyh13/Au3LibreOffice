#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oComment
	Local $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell B3
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a comment to Cell B3
	$oComment = _LOCalc_CommentAdd($oCell, "AutoIt Comment!")
	If @error Then _ERROR($oDoc, "Failed to add a comment. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Make the comment always visible
	_LOCalc_CommentVisible($oComment, True)
	If @error Then _ERROR($oDoc, "Failed to set comment visibility. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Comment's Text's settings to: Fit Width to Text = False, Fit Height to Text = False, Fit Text to Comment box size = True, Skip setting all spacing,
	; Set left and right spacing to 170 Hundredths of a Millimeter (HMM), set top and bottom spacing to 120 Hundredths of a Millimeter (HMM).
	_LOCalc_CommentTextSettings($oComment, False, False, True, Null, 170, 170, 120, 120)
	If @error Then _ERROR($oDoc, "Failed to set Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Comment settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_CommentTextSettings($oComment)
	If @error Then _ERROR($oDoc, "Failed to retrieve Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Comment's Text settings are as follows: " & @CRLF & _
			"Is the Comment Box's width adjusted to fit the Text's width? True/False: " & $avSettings[0] & @CRLF & _
			"Is the Comment Box's height adjusted to fit the Text's height? True/False: " & $avSettings[1] & @CRLF & _
			"Is the Text's font size adjusted to fit the Comment Box? True/False: " & $avSettings[2] & @CRLF & _
			"The Spacing for all of the borders is, in Hundredths of a Millimeter (HMM) (Will be 0 if they are not all equal): " & $avSettings[3] & @CRLF & _
			"The Left side Spacing between the text and the Comment box border is, in Hundredths of a Millimeter (HMM): " & $avSettings[4] & @CRLF & _
			"The Right side Spacing between the text and the Comment box border is, in Hundredths of a Millimeter (HMM): " & $avSettings[5] & @CRLF & _
			"The Top Spacing between the text and the Comment box border is, in Hundredths of a Millimeter (HMM): " & $avSettings[6] & @CRLF & _
			"The Bottom Spacing between the text and the Comment box border is, in Hundredths of a Millimeter (HMM): " & $avSettings[7])

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
