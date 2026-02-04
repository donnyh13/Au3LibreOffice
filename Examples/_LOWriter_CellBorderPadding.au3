#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $asCellNames
	Local $aCellBorder
	Local $iHMM

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create the Table, 3 columns, 5 rows.
	$oTable = _LOWriter_TableCreate($oDoc, $oViewCursor, 3, 5)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Cell names.
	$asCellNames = _LOWriter_TableCellsGetNames($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Cell names
	For $i = 0 To UBound($asCellNames) - 1
		; Retrieve each cell by name as returned in the table
		$oCell = _LOWriter_TableGetCellObjByName($oTable, $asCellNames[$i])
		If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set each Cell text String to each Cell's name.
		_LOWriter_CellString($oCell, $asCellNames[$i])
		If @error Then _ERROR($oDoc, "Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Retrieve 2nd down. 2nd over ("B2") Table Cell Object
	$oCell = _LOWriter_TableGetCellObjByName($oTable, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Border width so I can set Border padding.
	_LOWriter_CellBorderWidth($oCell, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell Border width settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4 Inch to Hundredths of a Millimeter (HMM).
	$iHMM = _LO_UnitConvert(0.25, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set cell Border padding values, 1/4 inch on all sides.
	_LOWriter_CellBorderPadding($oCell, $iHMM, $iHMM, $iHMM, $iHMM)

	; Retrieve current Border Padding settings. Return will be an Array, with Array elements in order of function parameters.
	$aCellBorder = _LOWriter_CellBorderPadding($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell Border Padding settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Cell Border padding settings are: " & @CRLF & _
			"Top = " & $aCellBorder[0] & " Hundredths of a Millimeter (HMM)" & @CRLF & _
			"Bottom = " & $aCellBorder[1] & " Hundredths of a Millimeter (HMM)" & @CRLF & _
			"Left = " & $aCellBorder[2] & " Hundredths of a Millimeter (HMM)" & @CRLF & _
			"Right = " & $aCellBorder[3] & " Hundredths of a Millimeter (HMM)")

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
