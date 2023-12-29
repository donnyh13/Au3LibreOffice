
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell, $oCell2
	Local $iVertOrient

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a Table, 2 rows, 2 columns
	$oTable = _LOWriter_TableCreate($oDoc, 2, 2)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Retrieve top left ("A1") Table Cell Object
	$oCell = _LOWriter_TableGetCellObjByName($oTable, "A1")
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended)

	;Retrieve top 2nd in cell ("B1" Table Cell Object
	$oCell2 = _LOWriter_TableGetCellObjByName($oTable, "B1")
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended)

	;Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "This Test Text.")
	If (@error > 0) Then _ERROR("Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended)

	;Set B1 Table Cell's Text.
	_LOWriter_CellString($oCell2, @CR)
	If (@error > 0) Then _ERROR("Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended)

	$iVertOrient = _LOWriter_CellVertOrient($oCell)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table cell current vertical orientation. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now demonstrate modifying a cell's vertical text orientation. The current setting is: " & $iVertOrient & _
			@CRLF & " The possible settings are: " & @CRLF & "$LOW_ORIENT_VERT_NONE(0)," & @CRLF & "$LOW_ORIENT_VERT_TOP(1)," & @CRLF & _
			"$LOW_ORIENT_VERT_CENTER(2)," & @CRLF & "$LOW_ORIENT_VERT_BOTTOM(3)")

	_LOWriter_CellVertOrient($oCell, $LOW_ORIENT_VERT_BOTTOM)
	If (@error > 0) Then _ERROR("Failed to set Text Table cell vertical orientation. Error:" & @error & " Extended:" & @extended)

	;Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "Text vertically aligned to the bottom.")
	If (@error > 0) Then _ERROR("Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have now set the Vertical Orientation to $LOW_ORIENT_VERT_BOTTOM(3).")

	_LOWriter_CellVertOrient($oCell, $LOW_ORIENT_VERT_CENTER)
	If (@error > 0) Then _ERROR("Failed to set Text Table cell vertical orientation. Error:" & @error & " Extended:" & @extended)

	;Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "Text vertically aligned to the Center.")
	If (@error > 0) Then _ERROR("Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have now set the Vertical Orientation to $LOW_ORIENT_VERT_CENTER(2).")

	_LOWriter_CellVertOrient($oCell, $LOW_ORIENT_VERT_TOP)
	If (@error > 0) Then _ERROR("Failed to set Text Table cell vertical orientation. Error:" & @error & " Extended:" & @extended)

	;Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "Text vertically aligned to the Top.")
	If (@error > 0) Then _ERROR("Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have now set the Vertical Orientation to $LOW_ORIENT_VERT_TOP(1).")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

