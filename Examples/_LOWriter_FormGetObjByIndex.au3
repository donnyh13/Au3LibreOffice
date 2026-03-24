#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oFormObj
	Local $iCount
	Local $avProps

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	_LOWriter_FormAdd($oDoc, "AutoIt_Form2")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a sub-Form in the form.
	_LOWriter_FormAdd($oForm, "AutoIt_SubForm")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of forms in the document.
	$iCount = _LOWriter_FormsGetCount($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of forms in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To $iCount - 1
		; Retrieve the Form Object.
		$oFormObj = _LOWriter_FormGetObjByIndex($oDoc, $i)
		If @error Then _ERROR($oDoc, "Failed to retrieve form Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve the properties for the Form to obtain the name.
		$avProps = _LOWriter_FormPropertiesGeneral($oFormObj)
		If @error Then _ERROR($oDoc, "Failed to retrieve form Properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "This form's name is: " & $avProps[0])
	Next

	; Retrieve a count of sub-forms in AutoIt_Form.
	$iCount = _LOWriter_FormsGetCount($oForm)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of sub-forms in AutoIt_Form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To $iCount - 1
		; Retrieve the Form Object.
		$oFormObj = _LOWriter_FormGetObjByIndex($oForm, $i)
		If @error Then _ERROR($oDoc, "Failed to retrieve form Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve the properties for the Form to obtain the name.
		$avProps = _LOWriter_FormPropertiesGeneral($oFormObj)
		If @error Then _ERROR($oDoc, "Failed to retrieve form Properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "This sub-form's name is: " & $avProps[0])
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

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
