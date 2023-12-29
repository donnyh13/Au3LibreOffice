#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $avArray[0]
	Local $iNewX, $iNewY

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Polygon Shape into the document, 5000 Wide by 7000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_LINE_POLYGON, 5000, 7000)
	If @error Then _ERROR("Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Shape's current settings for its third point.
	$avArray = _LOWriter_ShapePointsModify($oShape, 3)
	If @error Then _ERROR("Failed to retrieve Array of settings for a Shape point. Error:" & @error & " Extended:" & @extended)

	; I will retrieve the second points current position, and add to its X and Y values to determine my new point's new X and Y values.

	; Minus 1400 Micrometers from the X coordinate
	$iNewX = $avArray[0] - 1400

	; Add 400 Micrometers to the Y coordinate
	$iNewY = $avArray[1] + 400

	MsgBox($MB_OK, "", "Press Ok to modify the Shape's Point.")

	; Apply the modified X and Y coordinates
	_LOWriter_ShapePointsModify($oShape, 3, $iNewX, $iNewY)
	If @error Then _ERROR("Failed to modify Shape point. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Shape's Third Point type.")

	; Modify the Shape's Third point to be a Symmetrical Point Type
	_LOWriter_ShapePointsModify($oShape, 3, Null, Null, $LOW_SHAPE_POINT_TYPE_SYMMETRIC)
	If @error Then _ERROR("Failed to modify Shape point. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Shape's Third Point to no longer be a curve.")

	; Modify the Shape's Third point to be a normal point agin
	_LOWriter_ShapePointsModify($oShape, 3, Null, Null, Null, False)
	If @error Then _ERROR("Failed to modify Shape point. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings for the Third Point. Return will be an Array in order of Function parameters.
	$avArray = _LOWriter_ShapePointsModify($oShape, 3)
	If @error Then _ERROR("Failed to retrieve Array of settings for a Shape point. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Shape's X Coordinate is, in Micrometers: " & $avArray[0] & @CRLF & _
			"The Shape's Y Coordinate is, in Micrometers: " & $avArray[1] & @CRLF & _
			"The Shape's Point Type is, (See UDF Constants): " & $avArray[2] & @CRLF & _
			"Is this point a Curve? True/False: " & $avArray[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
