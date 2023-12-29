
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>
#include <Array.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iMicrometers
	Local $aiTabStops

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

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Create a Tab Stop for the demonstration.
	_LOWriter_DirFrmtParTabStopCreate($oViewCursor, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	;Convert 1/2" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Create another Tab Stop for the demonstration.
	_LOWriter_DirFrmtParTabStopCreate($oViewCursor, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	;Convert 1" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(1)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Create another Tab Stop for the demonstration.
	_LOWriter_DirFrmtParTabStopCreate($oViewCursor, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	;Retrieve an array of tab stop positions.
	$aiTabStops = _LOWriter_DirFrmtParTabStopList($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve array of Paragraph tab stop positions. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The number of Tabstops found is " & @extended & @CRLF & @CRLF & "I will now display the array of tab stop positions.")

	;Display the Array.
	_ArrayDisplay($aiTabStops)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



