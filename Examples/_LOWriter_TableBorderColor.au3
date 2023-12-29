
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local Const $iIntegerFlag = 1
	Local $aTableBorder

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create the Table, 3 rows, 5 columns
	$oTable = _LOWriter_TableCreate($oDoc, 3, 5)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Set the Border width so I can set the Border Color.
	_LOWriter_TableBorderWidth($oTable, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If (@error > 0) Then _ERROR("Failed to set Text Table cell Border width settings. Error:" & @error & " Extended:" & @extended)

	;Set the Border color to a random number for each side.
	_LOWriter_TableBorderColor($oTable, Random(0, 16777215, $iIntegerFlag), Random(0, 16777215, $iIntegerFlag), Random(0, 16777215, $iIntegerFlag), _
			Random(0, 16777215, $iIntegerFlag), Random(0, 16777215, $iIntegerFlag), Random(0, 16777215, $iIntegerFlag))
	If (@error > 0) Then _ERROR("Failed to set Text Table cell Border Color settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Border Style settings. Array will be 6 elements, in order of function parameters.
	$aTableBorder = _LOWriter_TableBorderColor($oTable)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table Border Color settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Border Color settings are: " & @CRLF & "Top = " & $aTableBorder[0] & @CRLF & "Bottom = " & $aTableBorder[1] & @CRLF & _
			"Left = " & $aTableBorder[2] & @CRLF & "Right = " & $aTableBorder[3] & @CRLF & "Vertical = " & $aTableBorder[4] & @CRLF & "Horizontal = " & $aTableBorder[5])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

