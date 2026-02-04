#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl, $oColumn
	Local $avColumn
	Local $iWidth
	Local $asComboItems[3] = ["One", "Two", "Three"]

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
	$oColumn = _LOWriter_FormConTableConColumnAdd($oControl, $LOW_FORM_CON_TYPE_COMBO_BOX)
	If @error Then _ERROR($oDoc, "Failed to insert a Table control column. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1.25 inches to Hundredths of a Millimeter (HMM)
	$iWidth = _LO_UnitConvert(1.25, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Column's General properties.
	_LOWriter_FormConTableConComboBoxGeneral($oColumn, "Renamed_AutoIt_Control", "ComboBox 5", Null, 35, True, False, $LOW_FORM_CON_MOUSE_SCROLL_FOCUS, $iWidth, _
			$asComboItems, "Select an Option", $LOW_ALIGN_HORI_CENTER, 2, True, "Some Additional Information", "This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to set column's general properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the column. Return will be an Array in order of function parameters.
	$avColumn = _LOWriter_FormConTableConComboBoxGeneral($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Column's current settings are: " & @CRLF & _
			"The Column's name is: " & $avColumn[0] & @CRLF & _
			"The Column's Label is: " & $avColumn[1] & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avColumn[2] & @CRLF & _
			"The Max text length allowed in the entry field is: " & $avColumn[3] & @CRLF & _
			"Is the Column currently enabled? True/False: " & $avColumn[4] & @CRLF & _
			"Is the Column currently Read-Only? True/False: " & $avColumn[5] & @CRLF & _
			"What happens when the mouse scroll wheel is used over the Column (See UDF Constants): " & $avColumn[6] & @CRLF & _
			"The Column's width is, in Hundredths of a Millimeter (HMM): " & $avColumn[7] & @CRLF & _
			"This is the Array of entries in the Column. I will only see how many elements the array has, which is: " & UBound($avColumn[8]) & @CRLF & _
			"The Column's Default Text is: " & $avColumn[9] & @CRLF & _
			"The Horizontal Alignment is: (See UDF Constants) " & $avColumn[10] & @CRLF & _
			"How many lines are displayed in the drop down? " & $avColumn[11] & @CRLF & _
			"Will selections be hidden when losing focus? True/False: " & $avColumn[12] & @CRLF & _
			"The Additional Information text is: " & $avColumn[13] & @CRLF & _
			"The Help text is: " & $avColumn[14] & @CRLF & _
			"The Help URL is: " & $avColumn[15])

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
