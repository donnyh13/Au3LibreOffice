#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oRefAnchor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a reference Mark at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Reference Mark at the ViewCursor, named "Ref. 1".
	_LOWriter_FieldRefMarkSet($oDoc, $oViewCursor, "Ref. 1", False)
	If @error Then _ERROR($oDoc, "Failed to insert a Reference Mark. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "I have inserted a reference Mark at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Reference Mark at the ViewCursor, named "Ref. 2".
	_LOWriter_FieldRefMarkSet($oDoc, $oViewCursor, "Ref. 2", False)
	If @error Then _ERROR($oDoc, "Failed to insert a Reference Mark. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "I have inserted a reference Mark at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Reference Mark at the ViewCursor, named "Ref. 3".
	_LOWriter_FieldRefMarkSet($oDoc, $oViewCursor, "Ref. 3", False)
	If @error Then _ERROR($oDoc, "Failed to insert a Reference Mark. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Line.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR & "The Reference Mark names contained in this document are: " & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will retrieve the anchor for Ref. 2 and insert some text after it.")

	; Retrieve the Reference Mark's Anchor (Text Cursor).
	$oRefAnchor = _LOWriter_FieldRefMarkGetAnchor($oDoc, "Ref. 2")
	If @error Then _ERROR($oDoc, "Failed to Retrieve a Reference Mark anchor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Anchor (Text Cursor)
	_LOWriter_CursorMove($oRefAnchor, $LOW_TEXTCUR_GO_RIGHT, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move a cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Some text.
	_LOWriter_CursorInsertString($oDoc, $oRefAnchor, " Some new text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

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
