#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iFillStyle

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly." & @CR & _
			"A second paragraph to demonstrate modifying formatting settings directly." & @CR & _
			"A third paragraph to demonstrate modifying formatting settings directly.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor down one paragraph
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_DOWN, 1)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Paragraph Fill Style
	$iFillStyle = _LOWriter_DirFrmtParAreaFillStyle($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph Fill Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Fill Style is: " & $iFillStyle & @CRLF & _
			"The result will be one of the following Constants:" & @CRLF & _
			"$LOW_AREA_FILL_STYLE_OFF, 0 Fill Style is off/ no background is applied." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_SOLID, 1 Fill Style is a solid color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_GRADIENT, 2 Fill Style is a gradient color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_HATCH, 3 Fill Style is a Hatch style color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_BITMAP, 4 Fill Style is a Bitmap.")

	; Set the paragraph at the current cursor's location background color to $LO_COLOR_GRAY.
	_LOWriter_DirFrmtParAreaColor($oViewCursor, $LO_COLOR_GRAY)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Paragraph Fill Style
	$iFillStyle = _LOWriter_DirFrmtParAreaFillStyle($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph Fill Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Fill Style is: " & $iFillStyle & @CRLF & _
			"The result will be one of the following Constants:" & @CRLF & _
			"$LOW_AREA_FILL_STYLE_OFF, 0 Fill Style is off/ no background is applied." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_SOLID, 1 Fill Style is a solid color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_GRADIENT, 2 Fill Style is a gradient color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_HATCH, 3 Fill Style is a Hatch style color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_BITMAP, 4 Fill Style is a Bitmap.")

	; Modify the Paragraph Gradient settings to: Preset Gradient name = $LOW_GRAD_NAME_TEAL_TO_BLUE
	_LOWriter_DirFrmtParAreaGradient($oDoc, $oViewCursor, $LOW_GRAD_NAME_TEAL_TO_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set Paragraph settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Paragraph Fill Style
	$iFillStyle = _LOWriter_DirFrmtParAreaFillStyle($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph Fill Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Paragraph's current Fill Style is: " & $iFillStyle & @CRLF & _
			"The result will be one of the following Constants:" & @CRLF & _
			"$LOW_AREA_FILL_STYLE_OFF, 0 Fill Style is off/ no background is applied." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_SOLID, 1 Fill Style is a solid color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_GRADIENT, 2 Fill Style is a gradient color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_HATCH, 3 Fill Style is a Hatch style color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_BITMAP, 4 Fill Style is a Bitmap.")

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
