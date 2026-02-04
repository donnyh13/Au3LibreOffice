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

	; Modify the Comment's Text Animation settings to: Set the Animation kind to $LOC_COMMENT_ANIMATION_KIND_SCROLL_THROUGH, Animation direction to Up,
	; Start inside the Comment Box = False, Finish with text visible = True, Repeat the Animation 25 times, Increment by 2 pixels,
	; and delay 20 Milliseconds before the next cycle.
	_LOCalc_CommentTextAnimation($oComment, $LOC_COMMENT_ANIMATION_KIND_SCROLL_THROUGH, $LOC_COMMENT_ANIMATION_DIR_UP, False, True, 25, 2, Null, 20)
	If @error Then _ERROR($oDoc, "Failed to set Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Comment settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_CommentTextAnimation($oComment)
	If @error Then _ERROR($oDoc, "Failed to retrieve Comment settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Comment's Text Animation settings are as follows: " & @CRLF & _
			"The Type of Text Animation is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Direction of the Text Animation is (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"Will the Animation begin with the Text visible inside the Comment Box? True/False: " & $avSettings[2] & @CRLF & _
			"Will the Animation end with the Text visible inside the Comment Box? True/False: " & $avSettings[3] & @CRLF & _
			"The Animation will be repeated: " & $avSettings[4] & " times." & @CRLF & _
			"The Animation will be Incremented by: " & $avSettings[5] & " pixels. (May be 0 if the current setting is in Hundredths of a Millimeter (HMM).)" & @CRLF & _
			"The Animation will be Incremented by: " & $avSettings[6] & " Hundredths of a Millimeter (HMM). (May be 0 if the current setting is in Pixels.)" & @CRLF & _
			"The Delay between animation cycles is, in Milliseconds: " & $avSettings[7])

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
