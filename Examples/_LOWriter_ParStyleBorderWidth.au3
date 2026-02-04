#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oParStyle
	Local $avParStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the Default Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style." & @CR & "Next Line")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Paragraph Style object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Default Paragraph Style Border Width to $LOW_BORDERWIDTH_THICK, And Connect border to True.
	_LOWriter_ParStyleBorderWidth($oParStyle, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, True)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avParStyleSettings = _LOWriter_ParStyleBorderWidth($oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Border width settings are as follows: " & @CRLF & _
			"Top width, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[0] & @CRLF & _
			"Bottom width, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[1] & @CRLF & _
			"Left width, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[2] & @CRLF & _
			"Right width, in Hundredths of a Millimeter (HMM): " & $avParStyleSettings[3] & @CRLF & _
			"Connect Paragraph Borders if the settings are identical? True/False: " & $avParStyleSettings[4])

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
