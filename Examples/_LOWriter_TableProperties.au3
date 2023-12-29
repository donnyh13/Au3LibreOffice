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

	; Create a Table, 5 rows, 3 columns, set Table Split to False, Background color to Teal and set a custom Table name.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3, False, $LOW_COLOR_TEAL, "CustomTableName")
	If @error Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	; Insert the Table into the document at the View Cursor's location.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	; Set the table settings to: Table Alignment = $LOW_ORIENT_HORI_NONE, Keep Table with next Paragraph = False, Set a New Table Name,
	; allow the Table to split across pages, Dont Allow the Rows to Split, and do repeat the Table heading, and include 3 rows as the heading.
	_LOWriter_TableProperties($oTable, $LOW_ORIENT_HORI_NONE, False, "NewName", True, True, True, 3)
	If @error Then _ERROR("Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve current settings.
	$avTableProps = _LOWriter_TableProperties($oTable)
	If @error Then _ERROR("Failed to retrieve Text Table settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Current Text Table settings are as follows: " & @CRLF & _
			"Table Alignment (See UDF Constants): " & $avTableProps[0] & @CRLF & _
			"Keep the Table with next Paragraph? True/False: " & $avTableProps[1] & @CRLF & _
			"Table Name: " & $avTableProps[2] & @CRLF & _
			"Allow table to split across pages? True/False: " & $avTableProps[3] & @CRLF & _
			"Allow Rows to split across pages? True/False: " & $avTableProps[4] & @CRLF & _
			"Repeat Table Heading? True/False: " & $avTableProps[5] & @CRLF & _
			"How many Rows are counted as the heading?: " & $avTableProps[6])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
