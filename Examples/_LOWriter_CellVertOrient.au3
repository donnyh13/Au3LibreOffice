#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell, $oCell2
	Local $iVertOrient

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

	; Retrieve top 2nd in cell ("B1" Table Cell Object
	$oCell2 = _LOWriter_TableGetCellObjByName($oTable, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "This Test Text.")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set B1 Table Cell's Text.
	_LOWriter_CellString($oCell2, @CR)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iVertOrient = _LOWriter_CellVertOrient($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell current vertical orientation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now demonstrate modifying a cell's vertical text orientation. The current setting is: " & $iVertOrient & _
			@CRLF & " The possible settings are: " & @CRLF & "$LOW_ORIENT_VERT_NONE(0)," & @CRLF & "$LOW_ORIENT_VERT_TOP(1)," & @CRLF & _
			"$LOW_ORIENT_VERT_CENTER(2)," & @CRLF & "$LOW_ORIENT_VERT_BOTTOM(3)")

	_LOWriter_CellVertOrient($oCell, $LOW_ORIENT_VERT_BOTTOM)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell vertical orientation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "Text vertically aligned to the bottom.")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have now set the Vertical Orientation to $LOW_ORIENT_VERT_BOTTOM(3).")

	_LOWriter_CellVertOrient($oCell, $LOW_ORIENT_VERT_CENTER)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell vertical orientation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "Text vertically aligned to the Center.")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have now set the Vertical Orientation to $LOW_ORIENT_VERT_CENTER(2).")

	_LOWriter_CellVertOrient($oCell, $LOW_ORIENT_VERT_TOP)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell vertical orientation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Table Cell's Text.
	_LOWriter_CellString($oCell, "Text vertically aligned to the Top.")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have now set the Vertical Orientation to $LOW_ORIENT_VERT_TOP(1).")

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
