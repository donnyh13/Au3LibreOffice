#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $iMicrometers
	Local $avSettings

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

	; Modify the Shape position settings. Horizontal Alignment = $LOW_ORIENT_HORI_CENTER, Skip Horizontal position,
	; Horizontal relation = $LOW_RELATIVE_PAGE, Mirror = True, Vertical align = $LOW_ORIENT_VERT_NONE, Vertical position = 1",
	; Vertical  relation = $LOW_RELATIVE_PAGE_PRINT, Keep inside = True, Anchor = $LOW_ANCHOR_AT_CHARACTER
	_LOWriter_ShapeTypePosition($oShape, $LOW_ORIENT_HORI_CENTER, Null, $LOW_RELATIVE_PAGE, True, $LOW_ORIENT_VERT_NONE, $iMicrometers, _
			$LOW_RELATIVE_PAGE_PRINT, True, $LOW_ANCHOR_AT_CHARACTER)
	If @error Then _ERROR("Failed to set Shape settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ShapeTypePosition($oShape)
	If @error Then _ERROR("Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Shape's position settings are as follows: " & @CRLF & _
			"The Shape's Horizontal alignment setting is (see UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Shape's Horizontal position is, in Micrometers: " & $avSettings[1] & @CRLF & _
			"The Shape's Horizontal relation setting is (see UDF Constants): " & $avSettings[2] & @CRLF & _
			"Mirror Shape position? True/False: " & $avSettings[3] & @CRLF & _
			"The Shape's Vertical alignment setting is (see UDF Constants): " & $avSettings[4] & @CRLF & _
			"The Shape's Vertical position is, in Micrometers: " & $avSettings[5] & @CRLF & _
			"The Shape's Vertical relation setting is (see UDF Constants): " & $avSettings[6] & @CRLF & _
			"Keep Shape within Text boundaries? True/False: " & $avSettings[7] & @CRLF & _
			"The Shape's anchor position is, (see UDF Constants): " & $avSettings[8])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
