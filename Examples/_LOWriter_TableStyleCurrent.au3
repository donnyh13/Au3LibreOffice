#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $sStyle

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 3 columns, 5 rows.
	$oTable = _LOWriter_TableCreate($oDoc, $oViewCursor, 3, 5, $LO_COLOR_TEAL, "AutoIt-Table", True, "Simple Grid Columns")
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the currently applied Table Style.
	$sStyle = _LOWriter_TableStyleCurrent($oDoc, $oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve current TableStyle. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current TableStyle applied to the table is: " & $sStyle & @CRLF & @CRLF & _
			"Press ok to apply ""Elegant"" Table Style to the Table.")

	; Apply "Elegant" tablestyle to the Table.
	_LOWriter_TableStyleCurrent($oDoc, $oTable, "Elegant")
	If @error Then _ERROR($oDoc, "Failed to apply TableStyle to Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
