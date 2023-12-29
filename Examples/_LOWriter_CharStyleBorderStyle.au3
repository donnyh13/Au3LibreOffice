#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oCharStyle
	Local $avCharStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I modify the Character style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a Character style.")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Move the cursor right 13 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 13)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Select the word "Demonstrate".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 11, True)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Set the Character style to "Example" Character style.
	_LOWriter_CharStyleSet($oDoc, $oViewCursor, "Example")
	If @error Then _ERROR("Failed to set the Character style. Error:" & @error & " Extended:" & @extended)

	; Retrieve the "Example" object.
	$oCharStyle = _LOWriter_CharStyleGetObj($oDoc, "Example")
	If @error Then _ERROR("Failed to retrieve Character style object. Error:" & @error & " Extended:" & @extended)

	; Set "Example" Character style Border Width to $LOW_BORDERWIDTH_THICK.
	_LOWriter_CharStyleBorderWidth($oCharStyle, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK, $LOW_BORDERWIDTH_THICK)
	If @error Then _ERROR("Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended)

	; Set "Example" Character style Border Style to, Top =  $LOW_BORDERSTYLE_DASH_DOT_DOT, Bottom = $LOW_BORDERSTYLE_FINE_DASHED,
	; Left = $LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP, Right = $LOW_BORDERSTYLE_SOLID
	_LOWriter_CharStyleBorderStyle($oCharStyle, $LOW_BORDERSTYLE_DASH_DOT_DOT, $LOW_BORDERSTYLE_FINE_DASHED, $LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP, $LOW_BORDERSTYLE_SOLID)
	If @error Then _ERROR("Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avCharStyleSettings = _LOWriter_CharStyleBorderStyle($oCharStyle)
	If @error Then _ERROR("Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Character style's current Border Style settings are as follows: " & @CRLF & _
			"Top style, (see UDF constants): " & $avCharStyleSettings[0] & @CRLF & _
			"Bottom style, (see UDF constants): " & $avCharStyleSettings[1] & @CRLF & _
			"Left style, (see UDF constants): " & $avCharStyleSettings[2] & @CRLF & _
			"Right style, (see UDF constants): " & $avCharStyleSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc