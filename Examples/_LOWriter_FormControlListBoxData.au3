#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl
	Local $avControl
	Local $asValues[3] = ["Test1", "Test2", "Test3"]
	Local $asTables[1] = ["Table_Name"]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormControlInsert($oForm, $LOW_FORM_CONTROL_TYPE_LIST_BOX, 500, 300, 6000, 2000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's Data properties.
	_LOWriter_FormControlListBoxData($oControl, "Datafield1", False, $LOW_FORM_CONTROL_SOURCE_TYPE_VALUE_LIST, $asValues)
	If @error Then _ERROR($oDoc, "Failed to modify the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the control. Return will be an Array in order of function parameters.
	$avControl = _LOWriter_FormControlListBoxData($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Control's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current settings are: " & @CRLF & _
			"The Database datafield or content source name is: " & $avControl[0] & @CRLF & _
			"Is user input required? True/False: " & $avControl[1] & @CRLF & _
			"What type of content is the List box filled with? (See UDF Constants): " & $avControl[2] & @CRLF & _
			"This is an array of entries. I'll just see the size of the array: " & UBound($avControl[3]) & @CRLF & _
			"The bound field value is: " & $avControl[4] & @CRLF & @CRLF & _
			"Press ok to set values for a Table now.")

	; Modify the Control's Data properties.
	_LOWriter_FormControlListBoxData($oControl, "Datafield1", False, $LOW_FORM_CONTROL_SOURCE_TYPE_TABLE, $asTables, 4)
	If @error Then _ERROR($oDoc, "Failed to modify the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the control. Return will be an Array in order of function parameters.
	$avControl = _LOWriter_FormControlListBoxData($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Control's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current settings are: " & @CRLF & _
			"The Database datafield or content source name is: " & $avControl[0] & @CRLF & _
			"Is user input required? True/False: " & $avControl[1] & @CRLF & _
			"What type of content is the List box filled with? (See UDF Constants): " & $avControl[2] & @CRLF & _
			"This is an array of entries. I'll just see the size of the array: " & UBound($avControl[3]) & @CRLF & _
			"The bound field value is: " & $avControl[4])

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
