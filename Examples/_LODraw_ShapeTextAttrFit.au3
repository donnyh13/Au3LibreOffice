#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oShape, $oDimension, $oTextCursor, $oTextBox
	Local $avSettings[0]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Textbox.
	$oTextBox = _LODraw_ShapeTextBoxInsert($oPage, 10500, 5000, -1, 1500)
	If @error Then _ERROR($oDoc, "Failed to  insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Dimension line Shape into the Page
	$oDimension = _LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_LINE_DIMENSION, 5000, 100, 1000, 500)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Page, 12000 Wide by 1500 High, 5000X and 5400Y.
	$oShape = _LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 12000, 1500, 5000, 5400)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the Shape.
	$oTextCursor = _LODraw_ShapeCreateTextCursor($oShape)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LODraw_CursorInsertString($oTextCursor, "Hi; This is some text entered using AutoIt!")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Text Box's Fit settings, set Fit to Width to True and Fit to Height to True.
	_LODraw_ShapeTextAttrFit($oTextBox, True, True, False)
	If @error Then _ERROR($oDoc, "Failed to modify Shape Text Attribute settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Dimension Line's Fit settings, set Fit to Frame to False, and Adjust Contour to True.
	_LODraw_ShapeTextAttrFit($oDimension, Null, Null, False, True)
	If @error Then _ERROR($oDoc, "Failed to modify Shape Text Attribute settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Rectangle Shape's Fit settings, set Wrap Text to False, and Resize Shape to True.
	_LODraw_ShapeTextAttrFit($oShape, Null, Null, Null, Null, False, True)
	If @error Then _ERROR($oDoc, "Failed to modify Shape Text Attribute settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Text Box's current fit settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_ShapeTextAttrFit($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Text Box's fit settings are as follows (Take note of the index values used in the array): " & @CRLF & _
			"[Index 0] Is the width of the shape adjusted to fit the text? True/False) " & $avSettings[0] & @CRLF & _
			"[Index 1] Is the height of the shape adjusted to fit the text? True/False): " & $avSettings[1] & @CRLF & _
			"[Index 2] Is the text fitted to the Frame? True/False: " & $avSettings[2])

	; Retrieve the Dimension Line's current fit settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_ShapeTextAttrFit($oDimension)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Dimension Line's fit settings are as follows (Take note of the index values used in the array): " & @CRLF & _
			"[Index 2] Is the text fitted to the Frame? True/False: " & $avSettings[2] & @CRLF & _
			"[Index 3] Is the text adjusted to contour to the frame? True/False): " & $avSettings[3])

	; Retrieve the Rectangle's current fit settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_ShapeTextAttrFit($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Rectangle's fit settings are as follows (Take note of the index values used in the array): " & @CRLF & _
			"[Index 4] Is the text wrapped? True/False: " & $avSettings[4] & @CRLF & _
			"[Index 5] Is the shape resized to fit the text? True/False): " & $avSettings[5])

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
