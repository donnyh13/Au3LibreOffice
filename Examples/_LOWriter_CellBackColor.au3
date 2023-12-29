
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $aCellBackGround
	Local Const $iIntegerFlag = 1

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

	;Retrieve current BackGround Color and Back Transparent settings. Return will be an Array with elements in order of function parameters.
	$aCellBackGround = _LOWriter_CellBackColor($oCell)
	If (@error > 0) Then _ERROR("Failed to retrieve current Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now demonstrate modifying a cell's background color settings. The current Background color value is: " & $aCellBackGround[0] & _
			@CRLF & " And the current Background Transparency setting is: " & $aCellBackGround[1])

	;Set the cell Background color to $LOW_COLOR_INDIGO, and BackTransparent to False.
	_LOWriter_CellBackColor($oCell, $LOW_COLOR_INDIGO, False)
	If (@error > 0) Then _ERROR("Failed to set Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended)

	;Set Table Cell's Text.
	_LOWriter_CellString($oCell, "Text with a colorful background.")
	If (@error > 0) Then _ERROR("Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended)

	;Retrieve current BackGround Color and Back Transparent settings.
	$aCellBackGround = _LOWriter_CellBackColor($oCell)
	If (@error > 0) Then _ERROR("Failed to retrieve current Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have set the cell's background color to $LOW_COLOR_INDIGO. The current Background color value is: " & $aCellBackGround[0] & _
			@CRLF & " And the current Background Transparency setting is: " & $aCellBackGround[1])

	;Set the cell BackTransparent to True.
	_LOWriter_CellBackColor($oCell, Null, True)
	If (@error > 0) Then _ERROR("Failed to set Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended)

	;Set Table Cell's Text.
	_LOWriter_CellString($oCell, "Text without a colorful background.")
	If (@error > 0) Then _ERROR("Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended)

	;Retrieve current BackGround Color and Back Transparent settings.
	$aCellBackGround = _LOWriter_CellBackColor($oCell)
	If (@error > 0) Then _ERROR("Failed to retrieve current Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have set the cell's background transparent setting to True, which means the color will not be visible. " & _
			"The current Background color value is: " & $aCellBackGround[0] & _
			@CRLF & " And the current Background Transparency setting is: " & $aCellBackGround[1])

	;Set the cell Background color to a random number, and BackTransparent to False.
	_LOWriter_CellBackColor($oCell, Random(0, 16777215, $iIntegerFlag), False)
	If (@error > 0) Then _ERROR("Failed to set Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended)

	;Set Table Cell's Text.
	_LOWriter_CellString($oCell, "Text with a random colorful background.")
	If (@error > 0) Then _ERROR("Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended)

	;Retrieve current BackGround Color and Back Transparent settings.
	$aCellBackGround = _LOWriter_CellBackColor($oCell)
	If (@error > 0) Then _ERROR("Failed to retrieve current Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have set the cell's background color to a random number, and the background transparent setting to False." & _
			"The current Background color value is: " & $aCellBackGround[0] & _
			@CRLF & " And the current Background Transparency setting is: " & $aCellBackGround[1])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

