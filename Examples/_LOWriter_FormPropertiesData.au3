#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm
	Local $avProps
	Local $asMaster[1] = ["Master_Field"], $asSlave[2] = ["Slave_Field1", "Slave_Field2"]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Form's Data properties.
	_LOWriter_FormPropertiesData($oForm, "Bibliography", $LOW_FORM_CONTENT_TYPE_SQL, 'SELECT "Year" FROM "Table_Name"', True, '( "Table"."Year" = "12" )', '"Year" ASC', $asMaster, $asSlave, True, False, False, False, $LOW_FORM_CYCLE_MODE_ACTIVE_RECORD, $LOW_FORM_NAV_BAR_MODE_YES)
	If @error Then _ERROR($oDoc, "Failed to set form Data properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the Form. Return will be an Array in order of function parameters.
	$avProps = _LOWriter_FormPropertiesData($oForm)
	If @error Then _ERROR($oDoc, "Failed to retrieve Form's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Form's current settings are: " & @CRLF & _
			"The Form's data source name is (if any): " & $avProps[0] & @CRLF & _
			"The Data source type is (See UDF Constants): " & $avProps[1] & @CRLF & _
			"The content to be used, in this case a SQL statement: " & $avProps[2] & @CRLF & _
			"Are SQL statements analyzed? True/False: " & $avProps[3] & @CRLF & _
			"The filter condition is (if any): " & $avProps[4] & @CRLF & _
			"The sort method is (if any): " & $avProps[5] & @CRLF & _
			"Are any master fields linked? (I will just do UBound to see array size): " & UBound($avProps[6]) & @CRLF & _
			"Are any slave fields linked? (I will just do UBound to see array size): " & UBound($avProps[7]) & @CRLF & _
			"Are Additions allowed? True/False: " & $avProps[8] & @CRLF & _
			"Are Modifications allowed? True/False: " & $avProps[9] & @CRLF & _
			"Are Deletions allowed? True/False: " & $avProps[10] & @CRLF & _
			"Is Data allowed to be added only? True/False: " & $avProps[11] & @CRLF & _
			"The Navigation bar mode is (See UDF Constants): " & $avProps[12] & @CRLF & _
			"The Tab cycle mode is (See UDF Constants): " & $avProps[13])

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
