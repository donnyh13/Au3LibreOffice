#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $avArray[0]
	Local $iNewX, $iNewY

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a FreeForm Line Shape into the document, 5000 Wide by 6000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_LINE_FREEFORM_LINE, 5000, 6000)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Shape's current settings for its first point.
	$avArray = _LOWriter_ShapePointsModify($oShape, 1)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of settings for a Shape point. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; I will retrieve the first point's current position, and add to its X and Y values to determine my new point's new X and Y values.

	; Minus 400 Hundredths of a Millimeter (HMM) from the X coordinate
	$iNewX = $avArray[0] - 400

	; Add 1600 Hundredths of a Millimeter (HMM) to the Y coordinate
	$iNewY = $avArray[1] + 1600

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to insert the new Point into the shape.")

	; Add the new Point using the new X and Y coordinates. The new point will be added after the called point (1), The point's type will be Symmetrical.
	_LOWriter_ShapePointsAdd($oShape, 1, $iNewX, $iNewY, $LOW_SHAPE_POINT_TYPE_SYMMETRIC)
	If @error Then _ERROR($oDoc, "Failed to modify Shape point. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
