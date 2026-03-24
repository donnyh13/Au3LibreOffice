#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm
	Local $avProps

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Form's General properties.
	_LOWriter_FormPropertiesGeneral($oForm, "New_Form", Null, Null, $LOW_FORM_SUBMIT_ENCODING_TEXT, $LOW_FORM_SUBMIT_METHOD_POST)
	If @error Then _ERROR($oDoc, "Failed to set form general properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the Form. Return will be an Array in order of function parameters.
	$avProps = _LOWriter_FormPropertiesGeneral($oForm)
	If @error Then _ERROR($oDoc, "Failed to retrieve Form's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Form's current settings are: " & @CRLF & _
			"The Form's name is: " & $avProps[0] & @CRLF & _
			"The URL or Path to open is (if any): " & $avProps[1] & @CRLF & _
			"The Frame to use to open the URL is (See UDF Constants): " & $avProps[2] & @CRLF & _
			"The Form encoding type is (See UDF Constants): " & $avProps[3] & @CRLF & _
			"The Form submit method is (See UDF Constants): " & $avProps[4])

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
