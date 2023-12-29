#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $asShapes

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Rectangle Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR("Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of Shape names currently in the document.
	$asShapes = _LOWriter_ShapesGetNames($oDoc)
	If @error Then _ERROR("Failed to retrieve a list of Shape names. Error:" & @error & " Extended:" & @extended)

	If (UBound($asShapes) > 0) Then

		; Retrieve the object for the first Shape listed in the Array.
		$oShape = _LOWriter_ShapeGetObjByName($oDoc, $asShapes[0][0])
		If @error Then _ERROR("Failed to retrieve a Shape Object. Error:" & @error & " Extended:" & @extended)

		MsgBox($MB_OK, "", "Press ok to delete the Shape.")

		; Delete the Shape.
		_LOWriter_ShapeDelete($oDoc, $oShape)
		If @error Then _ERROR("Failed to delete a Shape. Error:" & @error & " Extended:" & @extended)

	Else
		_ERROR("Something went wrong, and no Shapes were found.")
	EndIf

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
