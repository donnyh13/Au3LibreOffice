#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $iMicrometers
	Local $avTableProps

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 5 rows, 3 columns.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert the Table into the document at the View Cursor's location.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR($oDoc, "Failed to insert Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Table Alignment to $LOW_ORIENT_HORI_LEFT_AND_WIDTH so I can set Table margins.
	_LOWriter_TableProperties($oTable, $LOW_ORIENT_HORI_LEFT_AND_WIDTH)
	If @error Then _ERROR($oDoc, "Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1 inch to micrometers.
	$iMicrometers = _LO_ConvertToMicrometer(1)
	If @error Then _ERROR($oDoc, "Failed to convert inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set all Table margins to 1 inch except the right.
	_LOWriter_TableMargin($oTable, $iMicrometers, $iMicrometers, $iMicrometers, Null)
	If @error Then _ERROR($oDoc, "Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Margins are set to 1""")
	If @error Then _ERROR($oDoc, "Failed to insert Text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current settings.
	$avTableProps = _LOWriter_TableMargin($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Current Text Table Margin settings are: " & @CRLF & _
			"Top Table Margin: " & $avTableProps[0] & " Micrometers" & @CRLF & _
			"Bottom Table Margin: " & $avTableProps[1] & " Micrometers" & @CRLF & _
			"Left Table Margin: " & $avTableProps[2] & " Micrometers" & @CRLF & _
			"Right Table Margin: " & $avTableProps[3] & " Micrometers")

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
