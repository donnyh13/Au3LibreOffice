#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm
	Local $iForms

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a sub-form in the Form.
	_LOWriter_FormAdd($oForm, "AutoIt_Sub_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create another sub-form in the Form.
	_LOWriter_FormAdd($oForm, "AutoIt_Sub_Form2")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a Count of Forms contained in the Document.
	$iForms = _LOWriter_FormsGetCount($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of forms in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The number of Forms currently contained in the Document is: " & $iForms)

	; Retrieve a Count of Forms contained in the Form.
	$iForms = _LOWriter_FormsGetCount($oForm)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of forms in the Form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The number of sub-forms currently contained in the form is: " & $iForms)

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
