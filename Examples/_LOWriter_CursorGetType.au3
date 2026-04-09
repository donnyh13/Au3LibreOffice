#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iCursorType

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Get the type of Cursor.
	$iCursorType = _LOWriter_CursorGetType($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cursor Object type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The possible cursor type values are: " & @CRLF & _
			"$LOW_CURTYPE_TEXT_CURSOR (1)" & @CRLF & _
			"$LOW_CURTYPE_TABLE_CURSOR (2)" & @CRLF & _
			"$LOW_CURTYPE_VIEW_CURSOR (3)")

	; Display a message depending on what type the cursor is.
	Switch $iCursorType
		Case $LOW_CURTYPE_TEXT_CURSOR
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cursor Type is a Text Cursor, with an integer value of : " & $iCursorType & " — Or $LOW_CURTYPE_TEXT_CURSOR")

		Case $LOW_CURTYPE_TABLE_CURSOR
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cursor Type is a Table Cursor, with an integer value of : " & $iCursorType & " — Or $LOW_CURTYPE_TABLE_CURSOR")

		Case $LOW_CURTYPE_VIEW_CURSOR
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cursor Type is a View Cursor, with an integer value of : " & $iCursorType & " — Or $LOW_CURTYPE_VIEW_CURSOR")

		Case Else
			MsgBox($MB_OK + $MB_TOPMOST, Default, "Something went wrong.")
	EndSwitch

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
