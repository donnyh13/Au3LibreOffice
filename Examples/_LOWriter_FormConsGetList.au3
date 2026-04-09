#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm
	Local $avControls[0][0]
	Local $iCount = 0

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	_LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_TEXT_BOX, 500, 300, 2000, 3000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	_LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_PUSH_BUTTON, 2500, 1300, 2000, 1000, "AutoIt_Form_Control1")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	_LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_FILE_SELECTION, 1500, 4300, 4000, 2000, "AutoIt_Form_Control2")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of controls.
	$avControls = _LOWriter_FormConsGetList($oForm)
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of controls. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iCount = @extended

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There was found " & $iCount & " controls." & @CRLF & _
			"Press ok to cycle through the types of the controls.")

	For $i = 0 To $iCount - 1
		MsgBox($MB_OK + $MB_TOPMOST, Default, "This control's type is (See UDF Constants): " & $avControls[$i][1])
		If @error Then _ERROR($oDoc, "Failed to delete the control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Retrieve an array of controls, but only Text Boxes and Push buttons.
	$avControls = _LOWriter_FormConsGetList($oForm, BitOR($LOW_FORM_CON_TYPE_PUSH_BUTTON, $LOW_FORM_CON_TYPE_TEXT_BOX))
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of controls. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iCount = @extended

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There was found " & $iCount & " controls, filtering for only Text boxes and Push buttons." & @CRLF & _
			"Press ok to cycle through the types of the controls.")

	For $i = 0 To $iCount - 1
		MsgBox($MB_OK + $MB_TOPMOST, Default, "This control's type is (See UDF Constants): " & $avControls[$i][1] & @CRLF & _
				"Press ok to delete the control.")
		_LOWriter_FormConDelete($avControls[$i][0])
		If @error Then _ERROR($oDoc, "Failed to delete the control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
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
