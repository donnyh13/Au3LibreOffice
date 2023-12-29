
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
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
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style." & @CR & "Next Line" & @CR & "Next Line" & @LF)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Convert 1/2" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(0.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set "Default Paragraph Style" line space settings to, 1/4" above paragraph, 1/4" below paragraph, AddSpace = True, Line space mode =
	;$LOW_LINE_SPC_MODE_FIX, Line space height = 1/2",  and Page Line Space = False.
	_LOWriter_ParStyleSpace($oParStyle, $iMicrometers, $iMicrometers, True, $LOW_LINE_SPC_MODE_FIX, $iMicrometers2, False)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStyleSpace($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's current line Spacing settings are as follows: " & @CRLF & _
			"Spacing above the paragraph, in Micrometers: " & $avParStyleSettings[0] & @CRLF & _
			"Spacing below the paragraph, in Micrometers: " & $avParStyleSettings[1] & @CRLF & _
			"Do not apply spacing if above and/or below paragraph are the same style? True/False: " & $avParStyleSettings[2] & @CRLF & _
			"Line spacing mode, (See UDF constants): " & $avParStyleSettings[3] & @CRLF & _
			"Line spacing distance, in micrometers except for $LOW_LINE_SPC_MODE_PROP(0), which is a percentage: " & $avParStyleSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



