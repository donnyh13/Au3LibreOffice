#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oLayer, $oShapeLayer, $oShape, $oPage
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Page, 3000 Wide by 6000 High, 500X and 100Y.
	$oShape = _LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000, 500, 100)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the layer Object the Shape is inserted in.
	$oShapeLayer = _LODraw_ShapeLayerCurrent($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve shape's current Layer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current properties of the layer, so I can retrieve the name.
	$avSettings = _LODraw_DocLayerModify($oShapeLayer)
	If @error Then _ERROR($oDoc, "Failed to retrieve current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The rectangle is inserted into the layer: " & $avSettings[0] & ". " & @CRLF & "Press ok to switch it to ""Dimension Lines"".")

	; Retrieve the Dimension lines layer Object.
	$oLayer = _LODraw_DocLayerGetObjByName($oDoc, "measurelines")
	If @error Then _ERROR($oDoc, "Failed to retrieve the layer Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the shape's layer to Dimension Lines.
	_LODraw_ShapeLayerCurrent($oShape, $oLayer)
	If @error Then _ERROR($oDoc, "Failed to set shape's layer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the layer Object the Shape is inserted in.
	$oShapeLayer = _LODraw_ShapeLayerCurrent($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve shape's current Layer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current properties of the layer, so I can retrieve the name.
	$avSettings = _LODraw_DocLayerModify($oShapeLayer)
	If @error Then _ERROR($oDoc, "Failed to retrieve current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now the rectangle is in the " & $avSettings[0] & "  layer.")

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
