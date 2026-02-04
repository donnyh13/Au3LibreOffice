#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oLabel
	Local $mFont
	Local $avControl[0], $avFont[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Label Form Control
	$oLabel = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_LABEL, 3500, 2300, 10000, 2000, "AutoIt_Form_Label_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Font Descriptor.
	$mFont = _LOWriter_FontDescCreate("Times New Roman", $LOW_WEIGHT_BOLD, $LOW_POSTURE_ITALIC, 18, $LO_COLOR_BRICK, $LOW_UNDERLINE_BOLD, $LO_COLOR_GREEN, $LOW_STRIKEOUT_NONE, True, $LOW_RELIEF_NONE)
	If @error Then _ERROR($oDoc, "Failed to create a Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's General properties.
	_LOWriter_FormConLabelGeneral($oLabel, Null, "A Label control inserted by AutoIt!", Null, Null, Null, Null, $mFont)
	If @error Then _ERROR($oDoc, "Failed to modify a Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the properties of the control to modify the font.
	$avControl = _LOWriter_FormConLabelGeneral($oLabel)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Font's current settings. Return will be an Array in order of function parameters.
	$avFont = _LOWriter_FontDescEdit($avControl[6])
	If @error Then _ERROR($oDoc, "Failed to retrieve the Font Descriptor's current values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Font Descriptor's current settings are: " & @CRLF & _
			"The Name of the font used is: " & $avFont[0] & @CRLF & _
			"The Font weight is (See UDF Constants): " & $avFont[1] & @CRLF & _
			"The Font Italic setting is (See UDF Constants): " & $avFont[2] & @CRLF & _
			"The Font size is: " & $avFont[3] & @CRLF & _
			"The Font color is (as a RGB Color Integer): " & $avFont[4] & @CRLF & _
			"The Font underline style is (See UDF Constants): " & $avFont[5] & @CRLF & _
			"The Font Underline color is  (as a RGB Color Integer): " & $avFont[6] & @CRLF & _
			"The Strikeout line style is (See UDF Constants): " & $avFont[7] & @CRLF & _
			"Are individual words underlined? True/False: " & $avFont[8] & @CRLF & _
			"The Relief style is: (See UDF Constants) " & $avFont[9] & @CRLF & @CRLF & _
			"Press ok to modify the Font for this Label control.")

	; Modify the Font Descriptor.
	_LOWriter_FontDescEdit($avControl[6], "Arial", $LOW_WEIGHT_NORMAL, $LOW_POSTURE_NONE, 16, $LO_COLOR_LIME, $LOW_UNDERLINE_DBL_WAVE, $LO_COLOR_PURPLE, Null, False, $LOW_RELIEF_ENGRAVED)
	If @error Then _ERROR($oDoc, "Failed to modify the Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the new Font descriptor to the Label.
	_LOWriter_FormConLabelGeneral($oLabel, Null, Null, Null, Null, Null, Null, $avControl[6])
	If @error Then _ERROR($oDoc, "Failed to modify the Label control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
