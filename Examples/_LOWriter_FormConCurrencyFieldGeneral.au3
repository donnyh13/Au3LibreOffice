#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl, $oLabel
	Local $mFont
	Local $avControl

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_CURRENCY_FIELD, 500, 300, 6000, 2000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Label Form Control
	$oLabel = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_LABEL, 3500, 2300, 3000, 1000, "AutoIt_Form_Label_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Font Descriptor.
	$mFont = _LOWriter_FontDescCreate("Times New Roman", $LOW_WEIGHT_BOLD, $LOW_POSTURE_ITALIC, 18, $LO_COLOR_INDIGO, $LOW_UNDERLINE_BOLD, $LO_COLOR_GREEN, $LOW_STRIKEOUT_NONE, True, $LOW_RELIEF_NONE)
	If @error Then _ERROR($oDoc, "Failed to create a Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's General properties.
	_LOWriter_FormConCurrencyFieldGeneral($oControl, "Renamed_AutoIt_Control", $oLabel, Null, True, True, True, True, True, $LOW_FORM_CON_MOUSE_SCROLL_FOCUS, True, 1, _
			12.50, 100.75, 1, 13.50, 2, False, "$", True, True, False, Null, $mFont, $LOW_ALIGN_HORI_RIGHT, $LOW_ALIGN_VERT_TOP, $LO_COLOR_GREEN, _
			$LOW_FORM_CON_BORDER_WITHOUT, $LO_COLOR_GOLD, True, "Some Additional Information", "This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to modify the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the control. Return will be an Array in order of function parameters.
	$avControl = _LOWriter_FormConCurrencyFieldGeneral($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Control's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current settings are: " & @CRLF & _
			"The Control's name is: " & $avControl[0] & @CRLF & _
			"Is there a Label Control set for this control? (IsObj result): " & IsObj($avControl[1]) & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avControl[2] & @CRLF & _
			"Is Strict formatting enabled? True/False: " & $avControl[3] & @CRLF & _
			"Is the Control currently enabled? True/False: " & $avControl[4] & @CRLF & _
			"Is the Control currently visible? True/False: " & $avControl[5] & @CRLF & _
			"Is the Control currently Read-Only? True/False: " & $avControl[6] & @CRLF & _
			"Is the Control currently Printable? True/False: " & $avControl[7] & @CRLF & _
			"What happens when the mouse scroll wheel is used over the control (See UDF Constants): " & $avControl[8] & @CRLF & _
			"Is the Control a Tab Stop position? True/False: " & $avControl[9] & @CRLF & _
			"If the Control is a Tab Stop position, what order position is it? " & $avControl[10] & @CRLF & _
			"What is the minimum value allowed to be entered? " & $avControl[11] & @CRLF & _
			"What is the maximum value allowed to be entered? " & $avControl[12] & @CRLF & _
			"What is the value incremented by? " & $avControl[13] & @CRLF & _
			"What is the Default value? " & $avControl[14] & @CRLF & _
			"How many decimal places follow the value? " & $avControl[15] & @CRLF & _
			"Is there a Thousands separator? " & $avControl[16] & @CRLF & _
			"What is the currency symbol used (if any)? " & $avControl[17] & @CRLF & _
			"If there is a currency symbol, does it prefix the value? True/False: " & $avControl[18] & @CRLF & _
			"Does this control act as a spin button? True/False: " & $avControl[19] & @CRLF & _
			"Does the button action repeat if clicked and held? True/False: " & $avControl[20] & @CRLF & _
			"If the button action repeats when clicked and held, how much delay is between each repeat? (In Milliseconds): " & $avControl[21] & @CRLF & _
			"This is the current Font settings. I'll just check if it is a Map. " & IsMap($avControl[22]) & @CRLF & _
			"The Horizontal Alignment is: (See UDF Constants) " & $avControl[23] & @CRLF & _
			"The Vertical Alignment is: (See UDF Constants) " & $avControl[24] & @CRLF & _
			"The background color is (as a RGB Color Integer): " & $avControl[25] & @CRLF & _
			"The Border Style is (See UDF Constants): " & $avControl[26] & @CRLF & _
			"The Border color is (as a RGB Color Integer): " & $avControl[27] & @CRLF & _
			"Will selections be hidden when losing focus? True/False: " & $avControl[28] & @CRLF & _
			"The Additional Information text is: " & $avControl[29] & @CRLF & _
			"The Help text is: " & $avControl[30] & @CRLF & _
			"The Help URL is: " & $avControl[31])

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
