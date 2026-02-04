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
	$oComment = _LOCalc_CommentAdd($oCell, "This is a Comment added by AutoIt!")
	If @error Then _ERROR($oDoc, "Failed to add a comment. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Make the comment always visible
	_LOCalc_CommentVisible($oComment, True)
	If @error Then _ERROR($oDoc, "Failed to set comment visibility. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Comment Callout settings to: Callout connector line style = Angled, Line connection spacing from Comment box = 100 Hundredths of a Millimeter (HMM),
	; Callout connector line joint position = Vertical, Callout Connector line alignment on the Comment = Right.
	_LOCalc_CommentCallout($oComment, $LOC_COMMENT_CALLOUT_STYLE_ANGLED, 100, $LOC_COMMENT_CALLOUT_EXT_VERT, $LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_RIGHT)
	If @error Then _ERROR($oDoc, "Failed to set Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Comment settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_CommentCallout($oComment)
	If @error Then _ERROR($oDoc, "Failed to retrieve Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Comment's Callout settings are as follows: " & @CRLF & _
			"The Callout Connector Style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Callout Connector line spacing from the comment box is, in Hundredths of a Millimeter (HMM): " & $avSettings[1] & @CRLF & _
			"The Callout Connector line position on the Comment box is (See UDF Constants): " & $avSettings[2] & @CRLF & _
			"The Callout Connector line position Alignment or distance, depending on the Position setting, is (See UDF Constants in this case): " & $avSettings[3] & @CRLF & _
			"Is the Callout Connector line Optimally sized? True/False (Only available for $LOC_COMMENT_CALLOUT_STYLE_ANGLED_CONNECTOR): " & $avSettings[4] & @CRLF & _
			"The length of the Callout line is, in Hundredths of a Millimeter (HMM) (Only used if Optimal sizing is false): " & $avSettings[5])

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
