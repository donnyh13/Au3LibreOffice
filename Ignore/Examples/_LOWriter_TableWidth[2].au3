#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $avTableProps

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a Table, 5 rows, 3 columns.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If @error Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	; Insert the Table into the document at the View Cursor's location.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	; Set the Table ALignment to $LOW_ORIENT_HORI_LEFT so I can set Table relative width.
	_LOWriter_TableProperties($oTable, $LOW_ORIENT_HORI_LEFT)
	If @error Then _ERROR("Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended)

	; Set Table Relative width to 50%
	_LOWriter_TableWidth($oTable, Null, 50)
	If @error Then _ERROR("Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve current settings.
	$avTableProps = _LOWriter_TableWidth($oTable)
	If @error Then _ERROR("Failed to retrieve Text Table settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Current Text Table Width settings are: " & @CRLF & _
			"Table Width: " & $avTableProps[0] & @CRLF & _
			"Table Relative Width: " & $avTableProps[1] & "%" & @CRLF & _
			"Is the Table's Width currently Relative?: " & $avTableProps[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
