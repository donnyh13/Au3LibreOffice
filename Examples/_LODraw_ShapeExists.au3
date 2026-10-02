#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oShape
	Local $bReturn

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Page, 3000 Wide by 6000 High.
	$oShape = _LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Rename the shape to "AutoIt-Shape"
	_LODraw_ShapeName($oShape, "AutoIt-Shape")
	If @error Then _ERROR($oDoc, "Failed to rename a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the Document has a Shape by the name of "AutoIt-Shape"
	$bReturn = _LODraw_ShapeExists($oDoc, "AutoIt-Shape")
	If @error Then _ERROR($oDoc, "Failed to look for Shape name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does this Page contain a Shape named ""AutoIt-Shape""? True/ False. " & $bReturn)

	; Delete the Shape.
	_LODraw_ShapeDelete($oShape)
	If @error Then _ERROR($oDoc, "Failed to delete Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check again, if the Doc has a Shape by the name of "AutoIt-Shape"
	$bReturn = _LODraw_ShapeExists($oDoc, "AutoIt-Shape")
	If @error Then _ERROR($oDoc, "Failed to look for Shape name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now does this Page contain a Shape named ""AutoIt-Shape""? True/ False. " & $bReturn)

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
