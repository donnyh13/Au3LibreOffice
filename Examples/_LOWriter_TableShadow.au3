
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $iMicrometers
	Local $avShadow

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a Table, 5 rows, 3 columns
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Convert 1/2 Inch to Micrometers.
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set the Table shadow to 1/2 an inch wide, the color to $LOW_COLOR_DKGRAY, Transparent to False, and shadow location to
	;$LOW_SHADOW_BOTTOM_LEFT
	_LOWriter_TableShadow($oTable, $iMicrometers, $LOW_COLOR_DKGRAY, False, $LOW_SHADOW_BOTTOM_LEFT)
	If (@error > 0) Then _ERROR("Failed to set Table shadow settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve Table shadow settings. Return will be an Array, with values in order of function parameters.
	$avShadow = _LOWriter_TableShadow($oTable)
	If (@error > 0) Then _ERROR("Failed to retrieve Table shadow settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Table shadow values are as follows: " & @CRLF & _
			"Width = " & $avShadow[0] & " Micrometers." & @CRLF & _
			"Color = " & $avShadow[1] & " Long color format." & @CRLF & _
			"Is Color Transparent? True/False = " & $avShadow[2] & @CRLF & _
			"Shadow Location (See constants) = " & $avShadow[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



