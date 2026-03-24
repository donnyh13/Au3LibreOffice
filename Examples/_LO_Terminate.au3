#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"
#include "..\LibreOffice_Helper.au3"

Example()

Func Example()
	Local $oDoc

	If ProcessExists("soffice.bin") Then Exit MsgBox($MB_OK + $MB_TOPMOST, Default, "This example will only work if all LibreOffice instances are closed first.")

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "A New Writer Document was successfully opened. Press ""OK"" to close it.")

	; Close the document, don't save changes.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ProcessExists("soffice.bin") Then
		MsgBox($MB_OK + $MB_TOPMOST, Default, "All LibreOffice documents are closed, but the ""soffice.bin"" process still remains. Press ok to Terminate it.")

		; Terminate the remaining soffice.bin process.
		_LO_Terminate()
		If @error Then _ERROR($oDoc, "Failed to terminate the LibreOffice process. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "Does the ""soffice.bin"" process still exist? True/False: " & ((ProcessExists("soffice.bin") > 0) ? (True) : (False)) & @CRLF & _
				"Press ok to exit the script.")

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "All LibreOffice documents are closed, and also the ""soffice.bin"" process. Press ok to exit the script.")
	EndIf

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
