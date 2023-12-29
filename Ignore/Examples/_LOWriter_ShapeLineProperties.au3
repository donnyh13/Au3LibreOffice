#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $avSettings[0]
	Local $iMicrometers

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Folded Corner Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_BASIC_FOLDED_CORNER, 3000, 6000)
	If @error Then _ERROR("Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	; Convert 1/8" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.125)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Modify the Shape Line Properties settings to: Set the Line Style to $LOW_SHAPE_LINE_STYLE_3_DASHES_3_DOTS, Line Color to $LOW_COLOR_MAGENTA,
	; Width = 1/8", Transparency = 50%, Corner Style = $LOW_SHAPE_LINE_JOINT_BEVEL, Cap Style = $LOW_SHAPE_LINE_CAP_SQUARE
	_LOWriter_ShapeLineProperties($oShape, $LOW_SHAPE_LINE_STYLE_3_DASHES_3_DOTS, $LOW_COLOR_MAGENTA, $iMicrometers, 50, $LOW_SHAPE_LINE_JOINT_BEVEL, $LOW_SHAPE_LINE_CAP_SQUARE)
	If @error Then _ERROR("Failed to set Shape settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ShapeLineProperties($oShape)
	If @error Then _ERROR("Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Shape's Line Properties settings are as follows: " & @CRLF & _
			"The Line Style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Line color is, is long Color format: " & $avSettings[1] & @CRLF & _
			"The Line's Width is, in Micrometers: " & $avSettings[2] & @CRLF & _
			"The Line's transparency percentage is: " & $avSettings[3] & @CRLF & _
			"The Line Corner Style is, (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"The Line Cap Style is, (See UDF Constants): " & $avSettings[5])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
