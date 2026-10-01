#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage
	Local $avShapes
	Local $sShapeTypes = ""

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Page, 3000 Wide by 6000 High, 500X, 100Y.
	_LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000, 500, 100)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Connector Shape into the Page, 1000 Wide by 3000 High, 12100X, 4400Y.
	_LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_CONNECTOR_ARROWS, 900, 3000, 12100, 4400)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Line Shape into the Page, 4000 Wide by 0 High, 12000X, 4300Y.
	_LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_LINE_ARROW_LINE_CIRCLE_ARROW, 4000, 0, 12000, 4300)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all drawing Shapes in the current page.
	$avShapes = _LODraw_ShapesGetList($oPage, $LOD_SHAPE_TYPE_DRAWING_SHAPE)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shapes in page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		; Get the Shape type.
		$sShapeTypes &= "Shape #" & $i + 1 & " is the following type: " & _LODraw_DrawShapeGetType($avShapes[$i][0]) & @CRLF
		If @error Then _ERROR($oDoc, "Failed to identify shape type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The shapes in this page are the following types (See UDF Constants)." & @CRLF & $sShapeTypes)

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
