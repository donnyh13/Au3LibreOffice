
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $iMicrometers
	Local $aiPadding

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create the Table, 5 rows, 3 columns
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Set the outer table borders to medium thickness.
	_LOWriter_TableBorderWidth($oTable, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM)
	If (@error > 0) Then _ERROR("Failed to set Table Border width settings. Error:" & @error & " Extended:" & @extended)

	;Convert 1/2 Inch to Micrometers.
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set the Table padding to 1/2"
	_LOWriter_TableBorderPadding($oTable, $iMicrometers, $iMicrometers, $iMicrometers, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to set Table Border Padding settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve Table Border Padding settings. Return will be an Array, with values in order of function parameters.
	$aiPadding = _LOWriter_TableBorderPadding($oTable)
	If (@error > 0) Then _ERROR("Failed to retrieve Table Border Padding settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Table Border padding values are as follows: " & @CRLF & _
			"Top = " & $aiPadding[0] & " Micrometers" & @CRLF & _
			"Bottom = " & $aiPadding[1] & " Micrometers" & @CRLF & _
			"Left = " & $aiPadding[2] & " Micrometers" & @CRLF & _
			"Right = " & $aiPadding[3] & " Micrometers")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



