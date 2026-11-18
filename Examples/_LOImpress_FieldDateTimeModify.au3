#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTextBox, $oTextCursor, $oField
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_BLANK
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_BLANK)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Text Box.
	$oTextBox = _LOImpress_ShapeTextBoxInsert($oSlide, $LOI_SHAPE_TEXTBOX_TYPE_TEXTBOX, 15000, 13000)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a TextCursor in the TextBox.
	$oTextCursor = _LOImpress_ShapeCreateTextCursor($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to create a TextCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Time field in the TextBox, Set it to be variable, don't set a custom time, and set the Time format to 12 hour, Hour, Minute, Second, AM/PM.
	$oField = _LOImpress_FieldDateTimeInsert($oDoc, $oTextCursor, False, False, Null, $LOI_FIELD_TIME_FMT_12H_HMS_AMPM)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOImpress_FieldDateTimeModify($oField)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Field's current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If (@extended = 1) Then ; The Field is a Date Field.
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Date Field's settings are as follows: " & @CRLF & _
				"Is the field fixed? True/False: " & $avSettings[0] & @CRLF & _
				"The current Date structure is returned here, we'll just see if it is an Object: " & IsObj($avSettings[1]) & @CRLF & _
				"The Display format is (See UDF Constants): " & $avSettings[2])

	ElseIf (@extended = 2) Then ; The Field is a Time Field.
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Time Field's settings are as follows: " & @CRLF & _
				"Is the field fixed? True/False: " & $avSettings[0] & @CRLF & _
				"The current Time structure is returned here, we'll just see if it is an Object: " & IsObj($avSettings[1]) & @CRLF & _
				"The Display format is (See UDF Constants): " & $avSettings[2])
	EndIf

	; Insert a Date field in the TextBox, Set it to be variable.
	$oField = _LOImpress_FieldDateTimeInsert($oDoc, $oTextCursor, True, False, Null, $LOI_FIELD_DATE_FMT_DOW_MMMM_DD_YYYY)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOImpress_FieldDateTimeModify($oField)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Field's current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If (@extended = 1) Then ; The Field is a Date Field.
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Date Field's settings are as follows: " & @CRLF & _
				"Is the field fixed? True/False: " & $avSettings[0] & @CRLF & _
				"The current Date structure is returned here, we'll just see if it is an Object: " & IsObj($avSettings[1]) & @CRLF & _
				"The Display format is (See UDF Constants): " & $avSettings[2])

	ElseIf (@extended = 2) Then ; The Field is a Time Field.
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Time Field's settings are as follows: " & @CRLF & _
				"Is the field fixed? True/False: " & $avSettings[0] & @CRLF & _
				"The current Time structure is returned here, we'll just see if it is an Object: " & IsObj($avSettings[1]) & @CRLF & _
				"The Display format is (See UDF Constants): " & $avSettings[2])
	EndIf

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
