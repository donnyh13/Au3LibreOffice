#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $iHMM
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

	; Set the Table Alignment to $LOW_ORIENT_HORI_LEFT_AND_WIDTH so I can set Table margins.
	_LOWriter_TableProperties($oDoc, $oTable, $LOW_ORIENT_HORI_LEFT_AND_WIDTH)
	If @error Then _ERROR($oDoc, "Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1 inch to Hundredths of a Millimeter (HMM).
	$iHMM = _LO_UnitConvert(1, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set all Table margins to 1 inch except the right.
	_LOWriter_TableMargin($oTable, $iHMM, $iHMM, $iHMM, Null)
	If @error Then _ERROR($oDoc, "Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Margins are set to 1""")
	If @error Then _ERROR($oDoc, "Failed to insert Text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current settings.
	$avTableProps = _LOWriter_TableMargin($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Current Text Table Margin settings are: " & @CRLF & _
			"Top Table Margin: " & $avTableProps[0] & " Hundredths of a Millimeter (HMM)" & @CRLF & _
			"Bottom Table Margin: " & $avTableProps[1] & " Hundredths of a Millimeter (HMM)" & @CRLF & _
			"Left Table Margin: " & $avTableProps[2] & " Hundredths of a Millimeter (HMM)" & @CRLF & _
			"Right Table Margin: " & $avTableProps[3] & " Hundredths of a Millimeter (HMM)")

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
