#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $iColor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Rectangle Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR("Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	; Modify the Shape Background Color settings. Background color = $LOW_COLOR_TEAL, Background color is transparent = False
	_LOWriter_ShapeAreaColor($oShape, $LOW_COLOR_TEAL)
	If @error Then _ERROR("Failed to set Shape settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Shape settings. Return will be an Integer.
	$iColor = _LOWriter_ShapeAreaColor($oShape)
	If @error Then _ERROR("Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Shape's Background color settings are as follows: " & @CRLF & _
			"The Shape's Background color is, in Long color format: " & $iColor)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
