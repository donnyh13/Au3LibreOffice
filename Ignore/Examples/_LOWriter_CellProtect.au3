#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $bCellProtected

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a Table, 2 rows, 2 columns
	$oTable = _LOWriter_TableCreate($oDoc, 2, 2)
	If @error Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	; Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	; Retrieve top left ("A1") Table Cell Object
	$oCell = _LOWriter_TableGetCellObjByName($oTable, "A1")
	If @error Then _ERROR("Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended)

	; Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "Try to change this Text.")
	If @error Then _ERROR("Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended)

	$bCellProtected = _LOWriter_CellProtect($oCell)
	If @error Then _ERROR("Failed to retrieve Text Table cell current write protection setting. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now demonstrate modifying a cell's write protection setting. The current setting is: " & $bCellProtected & _
			@CRLF & " The possible settings are: " & @CRLF & "True, which means the cell cannot be edited, or " & @CRLF & _
			"False, which means the cell can be edited.")

	; Set the cell protection to True.
	_LOWriter_CellProtect($oCell, True)
	If @error Then _ERROR("Failed to set Text Table cell write protection setting. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have now set the cell protection to True, attempt to edit the text, and then press ok.")

	; Set the cell protection to False.
	_LOWriter_CellProtect($oCell, False)
	If @error Then _ERROR("Failed to set Text Table cell write protection setting. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have now set the cell protection to False, now try to edit the text, and then press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
