
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oParStyle
	Local $iMicrometers, $iMicrometers2
	Local $avParStyleSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text before I modify the Default Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style." & @CR & "Next Line")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended)

	;Set "Default Paragraph Style" Border Width to $LOW_BORDERWIDTH_THICK.
	_LOWriter_ParStyleBorderWidth($oParStyle, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set "Default Paragraph Style" Border padding to 1/4"
	_LOWriter_ParStyleBorderPadding($oParStyle, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStyleBorderPadding($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's current Border padding distance settings are as follows: " & @CRLF & _
			"All Padding distance, in Micrometers: " & $avParStyleSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Micrometers: " & $avParStyleSettings[1] & @CRLF & _
			"Bottom Padding distance, in Micrometers: " & $avParStyleSettings[2] & @CRLF & _
			"Left Padding distance, in Micrometers: " & $avParStyleSettings[3] & @CRLF & _
			"Right Padding distance, in Micrometers: " & $avParStyleSettings[4] & @CRLF & @CRLF & _
			"Press Ok, and I will demonstrate setting individual border padding settings.")

	;Convert 1/2" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(0.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set "Default Paragraph Style" Border padding to, Top and Right, 1/4", Bottom and left, 1/2".
	_LOWriter_ParStyleBorderPadding($oParStyle, Null, $iMicrometers, $iMicrometers2, $iMicrometers2, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStyleBorderPadding($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's current Border padding distance settings are as follows: " & @CRLF & _
			"All Padding distance, in Micrometers: " & $avParStyleSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Micrometers: " & $avParStyleSettings[1] & @CRLF & _
			"Bottom Padding distance, in Micrometers: " & $avParStyleSettings[2] & @CRLF & _
			"Left Padding distance, in Micrometers: " & $avParStyleSettings[3] & @CRLF & _
			"Right Padding distance, in Micrometers: " & $avParStyleSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

