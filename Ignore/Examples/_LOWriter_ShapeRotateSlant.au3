#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Rectangle Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR("Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	;Rotate the Shape 45 degrees.
	_LOWriter_ShapeRotateSlant($oShape, 45)
	If @error Then _ERROR("Failed to Rotate the Shape. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ShapeRotateSlant($oShape)
	If @error Then _ERROR("Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Shape's Rotation and Slant settings are as follows: " & @CRLF & _
			"The Shape is rotated " & $avSettings[0] & " degrees" & @CRLF & _
			"The Shape has a " & $avSettings[1] & " degree slant applied to it")

	;Rotate the Shape back to 0 degrees and apply a 23 degree slant to it.
	_LOWriter_ShapeRotateSlant($oShape, 0, 23)
	If @error Then _ERROR("Failed to Rotate the Shape and apply a slant. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ShapeRotateSlant($oShape)
	If @error Then _ERROR("Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Shape's Rotation and Slant settings are as follows: " & @CRLF & _
			"The Shape is rotated " & $avSettings[0] & " degrees" & @CRLF & _
			"The Shape has a " & $avSettings[1] & " degree slant applied to it")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
