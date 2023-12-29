
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $asCellNames
	Local $aCellBorder
	Local $iMicrometers

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

	;Retrieve Array of Cell names.
	$asCellNames = _LOWriter_TableGetCellNames($oTable)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table Cell names. Error:" & @error & " Extended:" & @extended)

	;Insert Cell names
	For $i = 0 To UBound($asCellNames) - 1
		;Retrieve each cell by name as returned in the table
		$oCell = _LOWriter_TableGetCellObjByName($oTable, $asCellNames[$i])
		If (@error > 0) Then _ERROR("Failed to retrieve Text Table Cell by name. Error:" & @error & " Extended:" & @extended)

		;Set each Cell text String to each Cell's name.
		_LOWriter_CellString($oCell, $asCellNames[$i])
		If (@error > 0) Then _ERROR("Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended)
	Next

	;Retrieve 2nd down. 2nd over ("B2") Table Cell Object
	$oCell = _LOWriter_TableGetCellObjByName($oTable, "B2")
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended)

	;Set the Border width so I can set Border padding.
	_LOWriter_CellBorderWidth($oCell, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If (@error > 0) Then _ERROR("Failed to set Text Table cell Border width settings. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4 Inch to Micrometers.
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set cell Border padding values, 1/4 inch on all sides.
	_LOWriter_CellBorderPadding($oCell, $iMicrometers, $iMicrometers, $iMicrometers, $iMicrometers)

	;Retrieve current Border Padding settings. Return will be an Array, with Array elements in order of function parameters.
	$aCellBorder = _LOWriter_CellBorderPadding($oCell)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table cell Border Padding settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Cell Border padding settings are: " & @CRLF & _
			"Top = " & $aCellBorder[0] & " Micrometers" & @CRLF & _
			"Bottom = " & $aCellBorder[1] & " Micrometers" & @CRLF & _
			"Left = " & $aCellBorder[2] & " Micrometers" & @CRLF & _
			"Right = " & $aCellBorder[3] & " Micrometers")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

