#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $iMicrometers, $iMicrometers2
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

	; Convert 1" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(1)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Convert 2" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(2)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Modify the Shape's Position to 1" over, 2" down, and set Protect Position to True.
	_LOWriter_ShapePosition($oShape, $iMicrometers, $iMicrometers2, True)
	If @error Then _ERROR("Failed to modify Shape's position. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ShapePosition($oShape)
	If @error Then _ERROR("Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Shape's Position settings are as follows: " & @CRLF & _
			"The Shape's X coordinate is, in Micrometers: " & $avSettings[0] & @CRLF & _
			"The Shape's Y coordinate is, in Micrometers: " & $avSettings[1] & @CRLF & _
			"Is the Shape's position protected against changes? True/False: " & $avSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
