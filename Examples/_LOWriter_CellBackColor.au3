#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $aCellBackGround
	Local Const $iIntegerFlag = 1

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create the Table, 2 rows, 2 columns
	$oTable = _LOWriter_TableCreate($oDoc, 2, 2)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR($oDoc, "Failed to insert Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve top left ("A1") Table Cell Object
	$oCell = _LOWriter_TableGetCellObjByName($oTable, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Background Color and Back Transparent settings. Return will be an Array with elements in order of function parameters.
	$aCellBackGround = _LOWriter_CellBackColor($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now demonstrate modifying a cell's background color settings. The current Background color value is: " & $aCellBackGround[0] & _
			@CRLF & " And the current Background Transparency setting is: " & $aCellBackGround[1])

	; Set the cell Background color to $LO_COLOR_INDIGO, and Background Transparent to False.
	_LOWriter_CellBackColor($oCell, $LO_COLOR_INDIGO, False)
	If @error Then _ERROR($oDoc, "Failed to set Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Table Cell's Text.
	_LOWriter_CellString($oCell, "Text with a colorful background.")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Background Color and Back Transparent settings.
	$aCellBackGround = _LOWriter_CellBackColor($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have set the cell's background color to $LO_COLOR_INDIGO. The current Background color value is: " & $aCellBackGround[0] & _
			@CRLF & " And the current Background Transparency setting is: " & $aCellBackGround[1])

	; Set the cell Background Transparent to True.
	_LOWriter_CellBackColor($oCell, Null, True)
	If @error Then _ERROR($oDoc, "Failed to set Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Table Cell's Text.
	_LOWriter_CellString($oCell, "Text without a colorful background.")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Background Color and Back Transparent settings.
	$aCellBackGround = _LOWriter_CellBackColor($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have set the cell's background transparent setting to True, which means the color will not be visible. " & _
			"The current Background color value is: " & $aCellBackGround[0] & _
			@CRLF & " And the current Background Transparency setting is: " & $aCellBackGround[1])

	; Set the cell Background color to a random number, and Background Transparent to False.
	_LOWriter_CellBackColor($oCell, Random(0, 16777215, $iIntegerFlag), False)
	If @error Then _ERROR($oDoc, "Failed to set Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Table Cell's Text.
	_LOWriter_CellString($oCell, "Text with a random colorful background.")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Background Color and Back Transparent settings.
	$aCellBackGround = _LOWriter_CellBackColor($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Text Table Cell Background settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have set the cell's background color to a random number, and the background transparent setting to False." & _
			"The current Background color value is: " & $aCellBackGround[0] & _
			@CRLF & " And the current Background Transparency setting is: " & $aCellBackGround[1])

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
