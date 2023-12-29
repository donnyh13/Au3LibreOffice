
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iMicrometers, $iMicrometers2
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly.")
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

	;Set the selected text's Border Width to $LOW_BORDERWIDTH_THICK.
	_LOWriter_DirFrmtCharBorderWidth($oViewCursor, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If (@error > 0) Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set the selected text's Border padding to 1/4"
	_LOWriter_DirFrmtCharBorderPadding($oViewCursor, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtCharBorderPadding($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The selected text's current Character Border padding settings are as follows: " & @CRLF & _
			"All Padding distance, in Micrometers: " & $avSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Micrometers: " & $avSettings[1] & @CRLF & _
			"Bottom Padding distance, in Micrometers: " & $avSettings[2] & @CRLF & _
			"Left Padding distance, in Micrometers: " & $avSettings[3] & @CRLF & @CRLF & _
			"Right Padding distance, in Micrometers: " & $avSettings[4] & @CRLF & @CRLF & _
			"Press Ok, and I will demonstrate setting individual border padding settings.")

	;Convert 1/2" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(0.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set the selected text's Border padding to, Top and Right, 1/4", Bottom and left, 1/2".
	_LOWriter_DirFrmtCharBorderPadding($oViewCursor, Null, $iMicrometers, $iMicrometers2, $iMicrometers2, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtCharBorderPadding($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The selected text's current Border color settings are as follows: " & @CRLF & _
			"All Padding distance, in Micrometers: " & $avSettings[0] & " This setting is best only used to set the distance, as" & _
			" the value will still be present, even though there are individual settings per side present." & @CRLF & _
			"Top Padding distance, in Micrometers: " & $avSettings[1] & @CRLF & _
			"Bottom Padding distance, in Micrometers: " & $avSettings[2] & @CRLF & _
			"Left Padding distance, in Micrometers: " & $avSettings[3] & @CRLF & @CRLF & _
			"Right Padding distance, in Micrometers: " & $avSettings[4] & @CRLF & @CRLF & _
			"Press ok to remove direct formating.")

	;Remove Direct Formatting.
	_LOWriter_DirFrmtCharBorderPadding($oViewCursor, Null, Null, Null, Null, Null, True)
	If (@error > 0) Then _ERROR("Failed to clear the selected text's direct formatting settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



