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

	; Modify the Comment Line Style settings to: Set the line style to $LOC_COMMENT_LINE_STYLE_DASH_DOT_ROUNDED, Line color = $LO_COLOR_INDIGO, Line width = 80 Hundredths of a Millimeter (HMM)
	; Line Transparency = 0, Line corner style = $LOC_COMMENT_LINE_JOINT_MITER, and Line end style = $LOC_COMMENT_LINE_CAP_FLAT
	_LOCalc_CommentLineProperties($oComment, $LOC_COMMENT_LINE_STYLE_DASH_DOT_ROUNDED, $LO_COLOR_INDIGO, 80, 0, $LOC_COMMENT_LINE_JOINT_MITER, $LOC_COMMENT_LINE_CAP_FLAT)
	If @error Then _ERROR($oDoc, "Failed to set Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Comment settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_CommentLineProperties($oComment)
	If @error Then _ERROR($oDoc, "Failed to retrieve Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Comment's Line Style settings are as follows: " & @CRLF & _
			"The Line style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Line Color is (as a RGB Color Integer): " & $avSettings[1] & @CRLF & _
			"The Line Width is, in Hundredths of a Millimeter (HMM): " & $avSettings[2] & @CRLF & _
			"The Percentage of Line Transparency is: " & $avSettings[3] & @CRLF & _
			"The Line corner style is (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"The Line end style is (See UDF Constants): " & $avSettings[5])

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
