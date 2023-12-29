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

	; Insert a Line Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_LINE_LINE, 3000, 6000)
	If @error Then _ERROR("Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.25)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Modify the Shape Arrow Style settings to: Set the Start Arrowhead to $LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, Start width = 1/4",
	; Start Center = True, Syncronize Start and End = True.
	_LOWriter_ShapeLineArrowStyles($oShape, $LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, $iMicrometers, True, True)
	If @error Then _ERROR("Failed to set Shape settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ShapeLineArrowStyles($oShape)
	If @error Then _ERROR("Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Shape's Arrow Style settings are as follows: " & @CRLF & _
			"The Start Arrowhead Style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Start Arrowhead Width is, in Micrometers: " & $avSettings[1] & @CRLF & _
			"Is the Start Arrowhead centered on the line end? True/False: " & $avSettings[2] & @CRLF & _
			"Is the Starting and Ending Arrowhead settings synchronized? True/False: " & $avSettings[3] & @CRLF & _
			"The End Arrowhead Style is (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"The End Arrowhead Width is, in Micrometers: " & $avSettings[5] & @CRLF & _
			"Is the Start Arrowhead centered on the line end? True/False: " & $avSettings[6])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
