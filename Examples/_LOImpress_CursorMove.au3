#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTextBox, $oTextCursor
	Local $avShapes

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_TITLE_ONLY
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_TITLE_ONLY)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of Textboxes in the current slide.
	$avShapes = _LOImpress_ShapesGetList($oSlide, BitOR($LOI_SHAPE_TYPE_TEXTBOX, $LOI_SHAPE_TYPE_TEXTBOX_TITLE, $LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE))
	If @error Then _ERROR($oDoc, "Failed to retrieve Shapes in slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Get the Object for the first Textbox found.
	$oTextBox = $avShapes[0][0]

	; Create a Text Cursor in the Textbox.
	$oTextCursor = _LOImpress_ShapeCreateTextCursor($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOImpress_CursorInsertString($oTextCursor, "1. Created by AutoIt!")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOImpress_CursorInsertString($oTextCursor, " The Cursor is here-->")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Each time I move the cursor, I will place a ""#"" mark at its final location.")

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now move the cursor to the start of the paragraph.")

	; Return the cursor back to the start.
	_LOImpress_CursorMove($oTextCursor, $LOI_TEXTCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOImpress_CursorInsertString($oTextCursor, "#")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now move the cursor to the right five characters.")

	; Move the Cursor right 5 characters, without selecting them.
	_LOImpress_CursorMove($oTextCursor, $LOI_TEXTCUR_GO_RIGHT, 5, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOImpress_CursorInsertString($oTextCursor, "#")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next I will move the cursor to the end of the paragraph.")

	; Move the Cursor to the end of the document, don't select text.
	_LOImpress_CursorMove($oTextCursor, $LOI_TEXTCUR_GOTO_END, 1, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOImpress_CursorInsertString($oTextCursor, "#")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next I will move the cursor to the left 3 characters, then left again, selecting 10 characters, then collapse the cursor to the end.")

	; Move the Cursor left 3 characters, without selecting them.
	_LOImpress_CursorMove($oTextCursor, $LOI_TEXTCUR_GO_LEFT, 3, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Cursor left 10 characters, selecting them.
	_LOImpress_CursorMove($oTextCursor, $LOI_TEXTCUR_GO_LEFT, 10, True)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Collapse the selection to the end.
	_LOImpress_CursorMove($oTextCursor, $LOI_TEXTCUR_COLLAPSE_TO_END, 1)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOImpress_CursorInsertString($oTextCursor, "#")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
