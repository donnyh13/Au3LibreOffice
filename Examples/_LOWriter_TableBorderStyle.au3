#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $aTableBorder

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 3 columns, 5 rows.
	$oTable = _LOWriter_TableCreate($oDoc, $oViewCursor, 3, 5)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Border width so I can set the Border Style.
	_LOWriter_TableBorderWidth($oTable, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell Border width settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Border Style, a different style on each side.
	_LOWriter_TableBorderStyle($oTable, $LOW_BORDERSTYLE_DOTTED, $LOW_BORDERSTYLE_DASHED, $LOW_BORDERSTYLE_DASH_DOT_DOT, $LOW_BORDERSTYLE_THICKTHIN_SMALLGAP, $LOW_BORDERSTYLE_EMBOSSED, $LOW_BORDERSTYLE_DOUBLE_THIN)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell Border Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Border Style settings.Array will be 6 elements, in order of function parameters.
	$aTableBorder = _LOWriter_TableBorderStyle($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Border Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Border Style settings are: " & @CRLF & "Top = " & $aTableBorder[0] & @CRLF & "Bottom = " & $aTableBorder[1] & @CRLF & _
			"Left = " & $aTableBorder[2] & @CRLF & "Right = " & $aTableBorder[3] & @CRLF & "Vertical = " & $aTableBorder[4] & @CRLF & "Horizontal = " & $aTableBorder[5] & _
			@CRLF & @CRLF & "see Constants in UDF for value meanings.")

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
