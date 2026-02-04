#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $bCellProtected

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 2 columns, 2 rows.
	$oTable = _LOWriter_TableCreate($oDoc, $oViewCursor, 2, 2)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve top left ("A1") Table Cell Object
	$oCell = _LOWriter_TableGetCellObjByName($oTable, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "Try to change this Text.")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$bCellProtected = _LOWriter_CellProtect($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell current write protection setting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now demonstrate modifying a cell's write protection setting. The current setting is: " & $bCellProtected & _
			@CRLF & " The possible settings are: " & @CRLF & "True, which means the cell cannot be edited, or " & @CRLF & _
			"False, which means the cell can be edited.")

	; Set the cell protection to True.
	_LOWriter_CellProtect($oCell, True)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell write protection setting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have now set the cell protection to True, attempt to edit the text, and then press ok.")

	; Set the cell protection to False.
	_LOWriter_CellProtect($oCell, False)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell write protection setting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have now set the cell protection to False, now try to edit the text, and then press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
