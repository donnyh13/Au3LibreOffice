#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $avTableProps

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 3 columns, 5 rows.
	$oTable = _LOWriter_TableCreate($oDoc, $oViewCursor, 3, 5)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Table Alignment to $LOW_ORIENT_HORI_LEFT so I can set Table relative width.
	_LOWriter_TableProperties($oDoc, $oTable, $LOW_ORIENT_HORI_LEFT)
	If @error Then _ERROR($oDoc, "Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Table Relative width to 50%
	_LOWriter_TableWidth($oTable, Null, 50)
	If @error Then _ERROR($oDoc, "Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current settings.
	$avTableProps = _LOWriter_TableWidth($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Current Text Table Width settings are: " & @CRLF & _
			"Table Width: " & $avTableProps[0] & @CRLF & _
			"Table Relative Width: " & $avTableProps[1] & "%" & @CRLF & _
			"Is the Table's Width currently Relative?: " & $avTableProps[2])

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
