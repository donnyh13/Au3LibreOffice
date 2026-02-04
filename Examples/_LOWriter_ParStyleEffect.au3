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
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Paragraph Style object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Default Paragraph Style Font effects to $LOW_RELIEF_EMBOSSED relief type.
	_LOWriter_ParStyleEffect($oParStyle, $LOW_RELIEF_EMBOSSED)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avParStyleSettings = _LOWriter_ParStyleEffect($oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current font Effects settings are as follows: " & @CRLF & _
			"Relief style (See UDF Constants): " & $avParStyleSettings[0] & @CRLF & _
			"Case style (See UDF Constants): " & $avParStyleSettings[1] & @CRLF & _
			"Are the words hidden? True/False: " & $avParStyleSettings[2] & @CRLF & _
			"Are the words outlined? True/False: " & $avParStyleSettings[3] & @CRLF & _
			"Do the words have a shadow? True/False: " & $avParStyleSettings[4] & @CRLF & @CRLF & _
			"I will now set Case to $LOW_CASEMAP_SM_CAPS, and Relief to $LOW_RELIEF_NONE.")

	; Set Default Paragraph Style Font effects to $LOW_RELIEF_NONE relief type, Case to $LOW_CASEMAP_SM_CAPS
	_LOWriter_ParStyleEffect($oParStyle, $LOW_RELIEF_NONE, $LOW_CASEMAP_SM_CAPS)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avParStyleSettings = _LOWriter_ParStyleEffect($oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's new font Effects settings are as follows: " & @CRLF & _
			"Relief style (See UDF Constants): " & $avParStyleSettings[0] & @CRLF & _
			"Case style (See UDF Constants): " & $avParStyleSettings[1] & @CRLF & _
			"Are the words hidden? True/False: " & $avParStyleSettings[2] & @CRLF & _
			"Are the words outlined? True/False: " & $avParStyleSettings[3] & @CRLF & _
			"Do the words have a shadow? True/False: " & $avParStyleSettings[4])

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
