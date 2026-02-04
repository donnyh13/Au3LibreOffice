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
	$oComment = _LOCalc_CommentAdd($oCell, "This is a Comment added by AutoIt! ")
	If @error Then _ERROR($oDoc, "Failed to add a comment. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Make the comment always visible
	_LOCalc_CommentVisible($oComment, True)
	If @error Then _ERROR($oDoc, "Failed to set comment visibility. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Comment Arrow Style settings to: Starting Arrow Style = $LOC_COMMENT_LINE_ARROW_TYPE_CF_MANY_ONE, Starting Arrow Width = 550 Hundredths of a Millimeter (HMM),
	; Center the starting arrow head on the line, Skip synchronizing the Start and end arrows. Set the Ending Arrow to $LOC_COMMENT_LINE_ARROW_TYPE_HALF_CIRCLE
	; Set the end Arrow width to 350 Hundredths of a Millimeter (HMM), and don't center the arrow head on the line.
	_LOCalc_CommentLineArrowStyles($oComment, $LOC_COMMENT_LINE_ARROW_TYPE_CF_MANY_ONE, 550, True, Null, $LOC_COMMENT_LINE_ARROW_TYPE_HALF_CIRCLE, 350, False)
	If @error Then _ERROR($oDoc, "Failed to set Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Comment settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_CommentLineArrowStyles($oComment)
	If @error Then _ERROR($oDoc, "Failed to retrieve Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Comment's Line Arrow Style settings are as follows: " & @CRLF & _
			"The Start Line Arrow style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Start Line Arrow Width is, in Hundredths of a Millimeter (HMM): " & $avSettings[1] & @CRLF & _
			"Is the Start Line Arrow head centered on the end of the line? True/False: " & $avSettings[2] & @CRLF & _
			"Is the Start and End Arrow values Synchronized? True/False: " & $avSettings[3] & @CRLF & _
			"The End Line Arrow style is (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"The End Line Arrow Width is, in Hundredths of a Millimeter (HMM): " & $avSettings[5] & @CRLF & _
			"Is the End Line Arrow head centered on the end of the line? True/False: " & $avSettings[6])

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
