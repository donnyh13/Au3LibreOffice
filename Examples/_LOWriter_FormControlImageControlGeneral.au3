#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl, $oLabel
	Local $avControl

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormControlInsert($oForm, $LOW_FORM_CONTROL_TYPE_IMAGE_CONTROL, 500, 300, 6000, 2000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Label Form Control
	$oLabel = _LOWriter_FormControlInsert($oForm, $LOW_FORM_CONTROL_TYPE_LABEL, 3500, 2300, 3000, 1000, "AutoIt_Form_Label_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's General properties.
	_LOWriter_FormControlImageControlGeneral($oControl, "Renamed_AutoIt_Control", $oLabel, Null, True, True, False, True, True, 1, $LOW_COLOR_BLUE, $LOW_FORM_CONTROL_BORDER_FLAT, $LOW_COLOR_GREEN, @ScriptDir & "\Extras\Plain.png", $LOW_FORM_CONTROL_IMG_BTN_SCALE_KEEP_ASPECT, "Some Additional Information", "This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to modify the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the control. Return will be an Array in order of function parameters.
	$avControl = _LOWriter_FormControlImageControlGeneral($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Control's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current settings are: " & @CRLF & _
			"The Control's name is: " & $avControl[0] & @CRLF & _
			"If there is a Label Control set for this control, this will be its Object. I'll just check IsObj: " & IsObj($avControl[1]) & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avControl[2] & @CRLF & _
			"Is the Control currently enabled? True/False: " & $avControl[3] & @CRLF & _
			"Is the Control currently visible? True/False: " & $avControl[4] & @CRLF & _
			"Is the Control currently Read-Only? True/False: " & $avControl[5] & @CRLF & _
			"Is the Control currently Printable? True/False: " & $avControl[6] & @CRLF & _
			"Is the Control a Tab Stop position? True/False: " & $avControl[7] & @CRLF & _
			"If the Control is a Tab Stop position, what order position is it? " & $avControl[8] & @CRLF & _
			"The Long Integer background color is: " & $avControl[9] & @CRLF & _
			"The Border Style is (See UDF Constants): " & $avControl[10] & @CRLF & _
			"The Border color is, in Long Integer format: " & $avControl[11] & @CRLF & _
			"The Graphic used (if any) will be here as an Graphic Object. I'll test if it is an Object: " & IsObj($avControl[12]) & @CRLF & _
			"The Scaling of the Graphic is (See UDF Constants): " & $avControl[13] & @CRLF & _
			"The Additional Information text is: " & $avControl[14] & @CRLF & _
			"The Help text is: " & $avControl[15] & @CRLF & _
			"The Help URL is: " & $avControl[16])

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
