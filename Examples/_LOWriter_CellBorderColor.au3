#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $iColor1, $iColor2, $iColor3, $iColor4
	Local Const $iIntegerFlag = 1
	Local $aCellBorder

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create the Table, 2 rows, 2 columns
	$oTable = _LOWriter_TableCreate($oDoc, 2, 2)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Retrieve top left ("A1") Table Cell Object
	$oCell = _LOWriter_TableGetCellObjByName($oTable, "A1")
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended)

	;Set the Border width so I can set the Border Color.
	_LOWriter_CellBorderWidth($oCell, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If (@error > 0) Then _ERROR("Failed to set Text Table cell Border width settings. Error:" & @error & " Extended:" & @extended)

	$iColor1 = Random(0, 16777215, $iIntegerFlag)
	$iColor2 = Random(0, 16777215, $iIntegerFlag)
	$iColor3 = Random(0, 16777215, $iIntegerFlag)
	$iColor4 = Random(0, 16777215, $iIntegerFlag)

	;Set the Border Color, a Random Color on each side.
	_LOWriter_CellBorderColor($oCell, $iColor1, $iColor2, $iColor3, $iColor4)
	If (@error > 0) Then _ERROR("Failed to set Text Table cell Border Color settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Border Style settings. Return will be an array, with elements in order of function parameters.
	$aCellBorder = _LOWriter_CellBorderColor($oCell)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table cell Border Color settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Border Color settings are: " & @CRLF & "Top = " & $aCellBorder[0] & @CRLF & "Bottom = " & $aCellBorder[1] & @CRLF & _
			"Left = " & $aCellBorder[2] & @CRLF & "Right = " & $aCellBorder[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
