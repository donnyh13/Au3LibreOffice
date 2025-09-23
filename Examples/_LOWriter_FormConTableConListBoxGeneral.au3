#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl, $oColumn
	Local $avColumn
	Local $iWidth
	Local $asListEntries[6] = ["Jon", "Jos", "MLipok", "donnyh13", "Argumentum", "GMK"] ; An array of entries, one entry per element.
	Local $aiSelected[2] = [2, 4] ; Set entry 3 (element 2 in the array), and entry 5 (element 4 in the Array) to be selected by default.

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
	$oColumn = _LOWriter_FormConTableConColumnAdd($oControl, $LOW_FORM_CON_TYPE_LIST_BOX)
	If @error Then _ERROR($oDoc, "Failed to insert a Table control column. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1.25 inches to Micrometers
	$iWidth = _LO_ConvertToMicrometer(1.25)
	If @error Then _ERROR($oDoc, "Failed to convert inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Column's General properties.
	_LOWriter_FormConTableConListBoxGeneral($oColumn, "Renamed_AutoIt_Control", "List Box 3", Null, True, True, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $iWidth, _
			$asListEntries, $LOW_ALIGN_HORI_LEFT, 3, $aiSelected, "Some Additional Information", "This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to set column's general properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the column. Return will be an Array in order of function parameters.
	$avColumn = _LOWriter_FormConTableConListBoxGeneral($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Column's current settings are: " & @CRLF & _
			"The Column's name is: " & $avColumn[0] & @CRLF & _
			"The Column's Label is: " & $avColumn[1] & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avColumn[2] & @CRLF & _
			"Is the Column currently enabled? True/False: " & $avColumn[3] & @CRLF & _
			"Is the Column currently Read-Only? True/False: " & $avColumn[4] & @CRLF & _
			"What happens when the mouse scroll wheel is used over the Column (See UDF Constants): " & $avColumn[5] & @CRLF & _
			"The Column's width is, in Micrometers: " & $avColumn[6] & @CRLF & _
			"This is the Array of entries in the Column. I will only see how many elements the array has, which is: " & UBound($avColumn[7]) & @CRLF & _
			"The Horizontal Alignment is: (See UDF Constants) " & $avColumn[8] & @CRLF & _
			"If this Column behaves as a drop down, this is how many lines will be shown: " & $avColumn[9] & @CRLF & _
			"This is the Array of Integers indicating which entries are selected by default. I will only see how many elements the array has, which is: " & UBound($avColumn[10]) & @CRLF & _
			"The Additional Information text is: " & $avColumn[11] & @CRLF & _
			"The Help text is: " & $avColumn[12] & @CRLF & _
			"The Help URL is: " & $avColumn[13])

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
