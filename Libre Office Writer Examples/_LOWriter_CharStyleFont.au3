
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oCharStyle
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

	;Retrieve the "Example" Character Style object.
	$oCharStyle = _LOWriter_CharStyleGetObj($oDoc, "Example")
	If (@error > 0) Then _ERROR("Failed to retrieve Character style object. Error:" & @error & " Extended:" & @extended)

	;Set "Example" Character style Font to "Arial", Font size to 18,Posture (Italic) to $LOW_POSTURE_ITALIC, and weight (Bold)
	;to $LOW_WEIGHT_BOLD
	_LOWriter_CharStyleFont($oDoc, $oCharStyle, "Arial", 18, $LOW_POSTURE_ITALIC, $LOW_WEIGHT_BOLD)
	If (@error > 0) Then _ERROR("Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avCharStyleSettings = _LOWriter_CharStyleFont($oDoc, $oCharStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Character style's current Font settings are as follows: " & @CRLF & _
			"Font name: " & $avCharStyleSettings[0] & @CRLF & _
			"Font size: " & $avCharStyleSettings[1] & @CRLF & _
			"Italic/Posture setting (See UDF Constants): " & $avCharStyleSettings[2] & @CRLF & _
			"Weight/Bold setting (See UDF Constants): " & $avCharStyleSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



