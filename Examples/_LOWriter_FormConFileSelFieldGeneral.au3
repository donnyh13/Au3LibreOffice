#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl
	Local $mFont
	Local $avControl

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_FILE_SELECTION, 500, 300, 6000, 2000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Font Descriptor.
	$mFont = _LOWriter_FontDescCreate("Times New Roman", $LOW_WEIGHT_BOLD, $LOW_POSTURE_ITALIC, 16, $LO_COLOR_INDIGO, $LOW_UNDERLINE_BOLD, $LO_COLOR_GREEN, $LOW_STRIKEOUT_NONE, False, $LOW_RELIEF_NONE)
	If @error Then _ERROR($oDoc, "Failed to create a Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's General properties.
	_LOWriter_FormConFileSelFieldGeneral($oControl, "Renamed_AutoIt_Control", Null, True, True, False, True, True, 1, "Please select a file.", $mFont, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_VERT_MIDDLE, $LO_COLOR_GRAY, $LOW_FORM_CON_BORDER_FLAT, $LO_COLOR_GREEN, False, "Some Additional Information", "This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to modify the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the control. Return will be an Array in order of function parameters.
	$avControl = _LOWriter_FormConFileSelFieldGeneral($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Control's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current settings are: " & @CRLF & _
			"The Control's name is: " & $avControl[0] & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avControl[1] & @CRLF & _
			"Is the Control currently enabled? True/False: " & $avControl[2] & @CRLF & _
			"Is the Control currently visible? True/False: " & $avControl[3] & @CRLF & _
			"Is the Control currently Read-Only? True/False: " & $avControl[4] & @CRLF & _
			"Is the Control currently Printable? True/False: " & $avControl[5] & @CRLF & _
			"Is the Control a Tab Stop position? True/False: " & $avControl[6] & @CRLF & _
			"If the Control is a Tab Stop position, what order position is it? " & $avControl[7] & @CRLF & _
			"The Default text of the control is: " & $avControl[8] & @CRLF & _
			"This is the current Font settings. I'll just check if it is a Map. " & IsMap($avControl[9]) & @CRLF & _
			"The Horizontal Alignment is: (See UDF Constants) " & $avControl[10] & @CRLF & _
			"The Vertical Alignment is: (See UDF Constants) " & $avControl[11] & @CRLF & _
			"The background color is (as a RGB Color Integer): " & $avControl[12] & @CRLF & _
			"The Border Style is (See UDF Constants): " & $avControl[13] & @CRLF & _
			"The Border color is (as a RGB Color Integer): " & $avControl[14] & @CRLF & _
			"Will selections be hidden when losing focus? True/False: " & $avControl[15] & @CRLF & _
			"The Additional Information text is: " & $avControl[16] & @CRLF & _
			"The Help text is: " & $avControl[17] & @CRLF & _
			"The Help URL is: " & $avControl[18])

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
