#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl, $oColumn
	Local $avColumn
	Local $iWidth

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_TABLE_CONTROL, 500, 300, 6000, 3000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Column
	$oColumn = _LOWriter_FormConTableConColumnAdd($oControl, $LOW_FORM_CON_TYPE_NUMERIC_FIELD)
	If @error Then _ERROR($oDoc, "Failed to insert a Table control column. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1.25 inches to Hundredths of a Millimeter (HMM)
	$iWidth = _LO_UnitConvert(1.25, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Column's General properties.
	_LOWriter_FormConTableConNumericFieldGeneral($oColumn, "Renamed_AutoIt_Control", "Numeric Field 12", Null, False, True, True, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS, 1.25, _
			50.00, 2, 18.44, 2, False, True, True, 50, $iWidth, $LOW_ALIGN_HORI_CENTER, False, "Some Additional Information", "This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to set column's general properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the column. Return will be an Array in order of function parameters.
	$avColumn = _LOWriter_FormConTableConNumericFieldGeneral($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Column's current settings are: " & @CRLF & _
			"The Column's name is: " & $avColumn[0] & @CRLF & _
			"The Column's Label is: " & $avColumn[1] & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avColumn[2] & @CRLF & _
			"Is formatting strict? True/False: " & $avColumn[3] & @CRLF & _
			"Is the Column currently enabled? True/False: " & $avColumn[4] & @CRLF & _
			"Is the Column currently Read-Only? True/False: " & $avColumn[5] & @CRLF & _
			"What happens when the mouse scroll wheel is used over the Column (See UDF Constants): " & $avColumn[6] & @CRLF & _
			"The minimum value allowed to be entered is: " & $avColumn[7] & @CRLF & _
			"The maximum value allowed to be entered is: " & $avColumn[8] & @CRLF & _
			"The increment to increase the value is: " & $avColumn[9] & @CRLF & _
			"The default value is: " & $avColumn[10] & @CRLF & _
			"How many decimal places are there?: " & $avColumn[11] & @CRLF & _
			"Is a Thousands separator used? True/False: " & $avColumn[12] & @CRLF & _
			"Does the field have a Spin button? True/False: " & $avColumn[13] & @CRLF & _
			"If the spin button is clicked and held, does the action repeat? True/False: " & $avColumn[14] & @CRLF & _
			"What is the delay, in Milliseconds, that the action repeats?: " & $avColumn[15] & @CRLF & _
			"The Column's width is, in Hundredths of a Millimeter (HMM): " & $avColumn[16] & @CRLF & _
			"The Horizontal Alignment is: (See UDF Constants) " & $avColumn[17] & @CRLF & _
			"Are selections hidden when the Column loses focus? True/False: " & $avColumn[18] & @CRLF & _
			"The Additional Information text is: " & $avColumn[19] & @CRLF & _
			"The Help text is: " & $avColumn[20] & @CRLF & _
			"The Help URL is: " & $avColumn[21])

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
