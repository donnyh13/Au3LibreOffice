#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2, $oViewCursor, $oViewCursor2, $oPar
	Local $aoPars[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "This Paragraph Contains some direct formatting that would normally be lost when copying it, unless" & _
			" I used the clipboard." & @CR)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Cursor to the beginning of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START, 1, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to move the View Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Cursor to the right 4 spaces and select.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 4, True)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to move the View Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the font color to $LO_COLOR_ORANGE, and highlight to $LO_COLOR_PURPLE
	_LOWriter_DirFrmtFontColor($oViewCursor, $LO_COLOR_ORANGE, Null, $LO_COLOR_PURPLE)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to set Font Color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Cursor to the right 43 spaces.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 43, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to move the View Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Cursor to the right 45 spaces.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 45, True)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to move the View Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Underline the selected portion in red.
	_LOWriter_DirFrmtUnderLine($oViewCursor, True, $LOW_UNDERLINE_BOLD_DASH_DOT_DOT, True, $LO_COLOR_RED)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to set underline settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a list of Paragraph Objects
	$aoPars = _LOWriter_ParObjCreateList($oViewCursor)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to retrieve array of paragraphs. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select the first Paragraph
	_LOWriter_ParObjSelect($oDoc, $aoPars[0])
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to select the paragraph. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Copy the Paragraph.
	$oPar = _LOWriter_ParObjCopy($oDoc)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to copy the paragraphs. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have now copied the paragraph, and the object pointing to this copy is stored in $oPar." & _
			" I will now open a second document and paste the paragraph in there, and it will have preserved its formatting.")

	; Create a New, visible, Blank Libre Office Document.
	$oDoc2 = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; "Paste" the paragraph object into the new document.
	_LOWriter_ParObjPaste($oDoc2, $oPar)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to paste the paragraph object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I can even paste the same paragraph several times if I want.")

	; Retrieve the document view cursor to insert text with.
	$oViewCursor2 = _LOWriter_DocGetViewCursor($oDoc2)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert 3 new paragraphs.
	_LOWriter_DocInsertString($oDoc2, $oViewCursor2, @CR & @CR & @CR)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; "Paste" the paragraph object into the new document.
	_LOWriter_ParObjPaste($oDoc2, $oPar)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to paste the paragraph object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the documents.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOWriter_DocClose($oDoc2, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $oDoc2, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	If IsObj($oDoc2) Then _LOWriter_DocClose($oDoc2, False)
	Exit
EndFunc
