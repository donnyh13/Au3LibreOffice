#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oParStyle
	Local $avParStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I modify the Default Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended)

	; Set "Default Paragraph Style" Font effects to $LOW_RELIEF_EMBOSSED relief type.
	_LOWriter_ParStyleEffect($oParStyle, $LOW_RELIEF_EMBOSSED)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStyleEffect($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's current font Effects settings are as follows: " & @CRLF & _
			"Relief style (See UDF Constants): " & $avParStyleSettings[0] & @CRLF & _
			"Case style (See UDF Constants): " & $avParStyleSettings[1] & @CRLF & _
			"Are the words hidden? True/False: " & $avParStyleSettings[2] & @CRLF & _
			"Are the words outlined? True/False: " & $avParStyleSettings[3] & @CRLF & _
			"Do the words have a shadow? True/False: " & $avParStyleSettings[4] & @CRLF & @CRLF & _
			"I will now set Case to $LOW_CASEMAP_SM_CAPS, and Relief to $LOW_RELIEF_NONE.")

	; Set "Default Paragraph Style" Font effects to $LOW_RELIEF_NONE relief type, Case to $LOW_CASEMAP_SM_CAPS
	_LOWriter_ParStyleEffect($oParStyle, $LOW_RELIEF_NONE, $LOW_CASEMAP_SM_CAPS)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStyleEffect($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's new font Effects settings are as follows: " & @CRLF & _
			"Relief style (See UDF Constants): " & $avParStyleSettings[0] & @CRLF & _
			"Case style (See UDF Constants): " & $avParStyleSettings[1] & @CRLF & _
			"Are the words hidden? True/False: " & $avParStyleSettings[2] & @CRLF & _
			"Are the words outlined? True/False: " & $avParStyleSettings[3] & @CRLF & _
			"Do the words have a shadow? True/False: " & $avParStyleSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
