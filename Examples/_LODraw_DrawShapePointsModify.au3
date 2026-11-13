#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oShape
	Local $avArray
	Local $iNewX, $iNewY

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Polygon Shape into the document, 5000 Wide by 7000 High, 12000X, 4300Y.
	$oShape = _LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_LINE_POLYGON, 5000, 7000, 12000, 4300)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Shape's current settings for its third point.
	$avArray = _LODraw_DrawShapePointsModify($oShape, 3)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of settings for a Shape point. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; I will retrieve the second points current position, and add to its X and Y values to determine my new point's new X and Y values.

	; Minus 1400 Hundredths of a Millimeter (HMM) from the X coordinate
	$iNewX = $avArray[0] - 1400

	; Add 400 Hundredths of a Millimeter (HMM) to the Y coordinate
	$iNewY = $avArray[1] + 400

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to modify the Shape's Point.")

	; Apply the modified X and Y coordinates
	_LODraw_DrawShapePointsModify($oShape, 3, $iNewX, $iNewY)
	If @error Then _ERROR($oDoc, "Failed to modify Shape point. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to modify the Shape's Third Point type.")

	; Modify the Shape's Third point to be a Symmetrical Point Type
	_LODraw_DrawShapePointsModify($oShape, 3, Null, Null, $LOD_DRAWSHAPE_POINT_TYPE_SYMMETRIC)
	If @error Then _ERROR($oDoc, "Failed to modify Shape point. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to modify the Shape's Third Point to no longer be a curve.")

	; Modify the Shape's Third point to be a normal point again
	_LODraw_DrawShapePointsModify($oShape, 3, Null, Null, Null, False)
	If @error Then _ERROR($oDoc, "Failed to modify Shape point. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the Third Point. Return will be an Array in order of Function parameters.
	$avArray = _LODraw_DrawShapePointsModify($oShape, 3)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of settings for a Shape point. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's X Coordinate is, in Hundredths of a Millimeter (HMM): " & $avArray[0] & @CRLF & _
			"The Shape's Y Coordinate is, in Hundredths of a Millimeter (HMM): " & $avArray[1] & @CRLF & _
			"The Shape's Point Type is, (See UDF Constants): " & $avArray[2] & @CRLF & _
			"Is this point a Curve? True/False: " & $avArray[3])

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
