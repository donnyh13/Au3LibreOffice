#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Rectangle Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR("Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	; Check if the document has a Shape by the name of "Shape 1"
	$bReturn = _LOWriter_DocHasShapeName($oDoc, "Shape 1")
	If @error Then _ERROR("Failed to look for Shape name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does this document contain a Shape named ""Shape 1""? True/ False. " & $bReturn)

	; Delete the Shape.
	_LOWriter_ShapeDelete($oDoc, $oShape)
	If @error Then _ERROR("Failed to delete Shape. Error:" & @error & " Extended:" & @extended)

	; Check again, if the document has a Shape by the name of "Shape 1"
	$bReturn = _LOWriter_DocHasShapeName($oDoc, "Shape 1")
	If @error Then _ERROR("Failed to look for Shape name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Now does this document contain a Shape named ""Shape 1""? True/ False. " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
