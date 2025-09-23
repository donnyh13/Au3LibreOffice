#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl, $oGroupBox
	Local $mFont
	Local $avControl

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_OPTION_BUTTON, 500, 300, 6000, 2000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Group Box Control
	$oGroupBox = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_GROUP_BOX, 2000, 2300, 6000, 2000, "AutoIt_Form_Control2")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Font Descriptor.
	$mFont = _LOWriter_FontDescCreate("Times New Roman", $LOW_WEIGHT_BOLD, $LOW_POSTURE_ITALIC, 16, $LO_COLOR_INDIGO, $LOW_UNDERLINE_BOLD, $LO_COLOR_GREEN, $LOW_STRIKEOUT_NONE, False, $LOW_RELIEF_NONE)
	If @error Then _ERROR($oDoc, "Failed to create a Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's General properties.
	_LOWriter_FormConOptionButtonGeneral($oControl, "Renamed_AutoIt_Control", "An AutoIt Labeled Control", $oGroupBox, Null, Null, True, True, True, True, 1, $LOW_FORM_CON_CHKBX_STATE_SELECTED, $mFont, $LOW_FORM_CON_BORDER_FLAT, $LOW_ALIGN_HORI_CENTER, $LOW_ALIGN_VERT_BOTTOM, $LO_COLOR_GOLD, False, Null, Null, "Some Additional Information", "This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to modify the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the control. Return will be an Array in order of function parameters.
	$avControl = _LOWriter_FormConOptionButtonGeneral($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Control's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current settings are: " & @CRLF & _
			"The Control's name is: " & $avControl[0] & @CRLF & _
			"The control's label is: " & $avControl[1] & @CRLF & _
			"If there is a Group Box Control set for this control, this will be its Object. I'll just check IsObj: " & IsObj($avControl[2]) & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avControl[3] & @CRLF & _
			"The Group name is (if any): " & $avControl[4] & @CRLF & _
			"Is the Control currently enabled? True/False: " & $avControl[5] & @CRLF & _
			"Is the Control currently visible? True/False: " & $avControl[6] & @CRLF & _
			"Is the Control currently Printable? True/False: " & $avControl[7] & @CRLF & _
			"Is the Control a Tab Stop position? True/False: " & $avControl[8] & @CRLF & _
			"If the Control is a Tab Stop position, what order position is it? " & $avControl[9] & @CRLF & _
			"The control's default state is (See UDF Constants): " & $avControl[10] & @CRLF & _
			"This is the current Font settings. I'll just check if it is a Map. " & IsMap($avControl[11]) & @CRLF & _
			"The control Style is (See UDF Constants): " & $avControl[12] & @CRLF & _
			"The Horizontal Alignment is: (See UDF Constants) " & $avControl[13] & @CRLF & _
			"The Vertical Alignment is: (See UDF Constants) " & $avControl[14] & @CRLF & _
			"The Long Integer background color is: " & $avControl[15] & @CRLF & _
			"Are line breaks allowed? True/False: " & $avControl[16] & @CRLF & _
			"If there is a Graphic used for the control, this would be its Object. I'll just check if it is an Object. " & IsObj($avControl[17]) & @CRLF & _
			"The Graphic Alignment is: (See UDF Constants) " & $avControl[18] & @CRLF & _
			"The Additional Information text is: " & $avControl[19] & @CRLF & _
			"The Help text is: " & $avControl[20] & @CRLF & _
			"The Help URL is: " & $avControl[21])

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
