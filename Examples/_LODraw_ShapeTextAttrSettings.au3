#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oShape, $oTextCursor
	Local $avSettings[0]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Page, 10000 Wide by 1500 High, 5000X and 5400Y.
	$oShape = _LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 10000, 1500, 5000, 5400)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the Shape.
	$oTextCursor = _LODraw_ShapeCreateTextCursor($oShape)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LODraw_CursorInsertString($oTextCursor, "Hi; This is some text entered using AutoIt!")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set 375 distance left and right, 148 distance at the top, 130 distance at the bottom,
	; Position the text at the top-left, and do not fit the text to the full width.
	_LODraw_ShapeTextAttrSettings($oShape, 375, 375, 148, 130, $LOD_PAR_TEXT_ANCHOR_TOP_LEFT, False)
	If @error Then _ERROR($oDoc, "Failed to modify Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_ShapeTextAttrSettings($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's settings are as follows: " & @CRLF & _
			"The amount of space between the left edge of the text and the left edge of the shape is (in Hundredths of a Millimeter (HMM)): " & $avSettings[0] & @CRLF & _
			"The amount of space between the right edge of the text and the right edge of the shape is (in Hundredths of a Millimeter (HMM)): " & $avSettings[1] & @CRLF & _
			"The amount of space between the top edge of the text and the top edge of the shape is (in Hundredths of a Millimeter (HMM)): " & $avSettings[2] & @CRLF & _
			"The amount of space between the bottom edge of the text and the bottom edge of the shape is (in Hundredths of a Millimeter (HMM)): " & $avSettings[3] & @CRLF & _
			"The position the text is anchored is (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"Will the text be fit to the shape's full width? True/False" & $avSettings[5])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
