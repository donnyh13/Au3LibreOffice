#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTextBox, $oTextCursor, $oDateField, $oTimeField
	Local $tDate, $tTime
	Local $avDate, $avTime, $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_BLANK
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_BLANK)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure.
	$tDate = _LOImpress_DateStructCreate(2026, 09, 03)
	If @error Then _ERROR($oDoc, "Failed to create a Date structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Time Structure.
	$tTime = _LOImpress_DateStructCreate(Null, Null, Null, 15, 25, 56, 12)
	If @error Then _ERROR($oDoc, "Failed to create a Time structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Text Box.
	$oTextBox = _LOImpress_ShapeTextBoxInsert($oSlide, $LOI_SHAPE_TEXTBOX_TYPE_TEXTBOX, 15000, 13000)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a TextCursor in the TextBox.
	$oTextCursor = _LOImpress_ShapeCreateTextCursor($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to create a TextCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Date field in the TextBox using the custom date.
	$oDateField = _LOImpress_FieldDateTimeInsert($oDoc, $oTextCursor, True, True, $tDate, $LOI_FIELD_DATE_FMT_DOW_MMM_DD_YYYY)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a newline after the Field.
	_LOImpress_CursorInsertString($oTextCursor, @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Time field in the TextBox using the custom Time.
	$oTimeField = _LOImpress_FieldDateTimeInsert($oDoc, $oTextCursor, False, True, $tTime, $LOI_FIELD_TIME_FMT_12H_HMS_MS_AMPM)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOImpress_FieldDateTimeModify($oDateField)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Field's current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current values used in the Date Field. The Date Field's structure will be in the #1 array element.
	; The Time values will be present, but we will ignore them since we know this is a Date field.
	$avDate = _LOImpress_DateStructModify($avSettings[1])
	If @error Then _ERROR($oDoc, "Failed to retrieve Date Structure values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Date Structure's values are as follows: " & @CRLF & _
			"[Element 0] The year is: " & $avDate[0] & @CRLF & _
			"[Element 1] The month is: " & $avDate[1] & @CRLF & _
			"[Element 2] The day is: " & $avDate[2])

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOImpress_FieldDateTimeModify($oTimeField)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Field's current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current values used in the Time Field. The Time Field's structure will be in the #1 array element.
	; The Date values will be present, but we will ignore them since we know this is a Time field.
	$avTime = _LOImpress_DateStructModify($avSettings[1])
	If @error Then _ERROR($oDoc, "Failed to retrieve Time Structure values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Time Structure's values are as follows (Note the array element positions.): " & @CRLF & _
			"[Element 3] The Hour is: " & $avTime[3] & @CRLF & _
			"[Element 4] The minute is: " & $avTime[4] & @CRLF & _
			"[Element 5] The second is: " & $avTime[5] & @CRLF & _
			"[Element 6] The millisecond is: " & $avTime[6])

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
