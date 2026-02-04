#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iHMM, $iHMM2
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the paragraph at the current cursor's location Border Width to $LOW_BORDERWIDTH_THICK.
	_LOWriter_DirFrmtParBorderWidth($oViewCursor, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(0.25, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the selected text's Border padding to 1/4"
	_LOWriter_DirFrmtParBorderPadding($oViewCursor, $iHMM)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_DirFrmtParBorderPadding($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Paragraph Border padding settings are as follows: " & @CRLF & _
			"All Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[1] & @CRLF & _
			"Bottom Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[2] & @CRLF & _
			"Left Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[3] & @CRLF & @CRLF & _
			"Right Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[4] & @CRLF & @CRLF & _
			"Press Ok, and I will demonstrate setting individual border padding settings.")

	; Convert 1/2" to Hundredths of a Millimeter (HMM)
	$iHMM2 = _LO_UnitConvert(0.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the selected text's Border padding to, Top and Right, 1/4", Bottom and left, 1/2".
	_LOWriter_DirFrmtParBorderPadding($oViewCursor, Null, $iHMM, $iHMM2, $iHMM2, $iHMM)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_DirFrmtParBorderPadding($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current paragraph Border padding settings are as follows: " & @CRLF & _
			"All Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[1] & @CRLF & _
			"Bottom Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[2] & @CRLF & _
			"Left Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[3] & @CRLF & @CRLF & _
			"Right Padding distance, in Hundredths of a Millimeter (HMM): " & $avSettings[4] & @CRLF & @CRLF & _
			"Press ok to remove direct formatting.")

	; Remove Direct Formatting.
	_LOWriter_DirFrmtParBorderPadding($oViewCursor, Null, Null, Null, Null, Null, True)
	If @error Then _ERROR($oDoc, "Failed to clear the selected text's direct formatting settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
