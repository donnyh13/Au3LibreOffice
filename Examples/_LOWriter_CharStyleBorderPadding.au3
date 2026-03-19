
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oCharStyle
	Local $iMicrometers, $iMicrometers2
	Local $avCharStyleSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text before I modify the Character style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a Character style.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the cursor right 13 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 13)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Select the word "Demonstrate".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 11, True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Set the Character style to "Example" Character style.
	_LOWriter_CharStyleSet($oDoc, $oViewCursor, "Example")
	If (@error > 0) Then _ERROR("Failed to set the Character style. Error:" & @error & " Extended:" & @extended)

	;Retrieve the "Example" object.
	$oCharStyle = _LOWriter_CharStyleGetObj($oDoc, "Example")
	If (@error > 0) Then _ERROR("Failed to retrieve Character style object. Error:" & @error & " Extended:" & @extended)

	;Set "Example" Character style Border Width to $LOW_BORDERWIDTH_THICK.
	_LOWriter_CharStyleBorderWidth($oCharStyle, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If (@error > 0) Then _ERROR("Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set "Example" Character style Border padding to 1/4"
	_LOWriter_CharStyleBorderPadding($oCharStyle, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avCharStyleSettings = _LOWriter_CharStyleBorderPadding($oCharStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Character style's current Border Padding distance settings are as follows: " & @CRLF & _
			"All Padding distance, in Micrometers: " & $avCharStyleSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Micrometers: " & $avCharStyleSettings[1] & @CRLF & _
			"Bottom Padding distance, in Micrometers: " & $avCharStyleSettings[2] & @CRLF & _
			"Left Padding distance, in Micrometers: " & $avCharStyleSettings[3] & @CRLF & _
			"Right Padding distance, in Micrometers: " & $avCharStyleSettings[4] & @CRLF & @CRLF & _
			"Press Ok, and I will demonstrate setting individual border padding settings.")

	;Convert 1/2" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(0.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set "Example" Character style Border padding to, Top and Right, 1/4", Bottom and left, 1/2".
	_LOWriter_CharStyleBorderPadding($oCharStyle, Null, $iMicrometers, $iMicrometers2, $iMicrometers2, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avCharStyleSettings = _LOWriter_CharStyleBorderPadding($oCharStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Character style's current Border Padding distance settings are as follows: " & @CRLF & _
			"All Padding distance, in Micrometers: " & $avCharStyleSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Micrometers: " & $avCharStyleSettings[1] & @CRLF & _
			"Bottom Padding distance, in Micrometers: " & $avCharStyleSettings[2] & @CRLF & _
			"Left Padding distance, in Micrometers: " & $avCharStyleSettings[3] & @CRLF & _
			"Right Padding distance, in Micrometers: " & $avCharStyleSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



