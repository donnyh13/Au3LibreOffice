#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oStyle, $oSlide, $oShape
	Local $asStyles
	Local $sStyles = ""

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_TITLE_ONLY
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_TITLE_ONLY)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Slide, 3000 Wide by 6000 High.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000, 2000, 3500)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Shape Style named "AutoIt-Test"
	$oStyle = _LOImpress_ShapeStyleCreate($oDoc, "AutoIt-Test")
	If @error Then _ERROR($oDoc, "Failed to create a Shape Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the "AutoIt-Test" style to the shape.
	_LOImpress_ShapeStyleCurrent($oDoc, $oShape, "AutoIt-Test")
	If @error Then _ERROR($oDoc, "Failed to set the Shape Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of all user-created Shape Style names.
	$asStyles = _LOImpress_ShapeStylesGetNames($oDoc, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($asStyles) - 1
		$sStyles &= $asStyles[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I found " & UBound($asStyles) & " User-Created Shape Style(s). With the following name(s):" & @CRLF & @CRLF & _
			$sStyles & @CRLF & @CRLF & "Press ok to delete it.")

	; Delete the Shape Style named "AutoIt-Test", Force it to be deleted replacing anywhere it's used with "Filled Green"
	_LOImpress_ShapeStyleDelete($oDoc, $oStyle, True, "Filled Green")
	If @error Then _ERROR($oDoc, "Failed to delete a shape style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of all user-created Shape Style names.
	$asStyles = _LOImpress_ShapeStylesGetNames($oDoc, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStyles = ""

	For $i = 0 To UBound($asStyles) - 1
		$sStyles &= $asStyles[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I found " & UBound($asStyles) & " User-Created Shape Style(s). With the following name(s):" & @CRLF & @CRLF & _
			$sStyles)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
