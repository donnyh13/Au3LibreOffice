#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $sSaveName, $sSavepath
	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "A New Calc Document was successfully opened. Press OK to close and save it.")

	; Create a Temporary Unique File name.
	$sSaveName = "TestCloseDocument_" & @YEAR & "_" & @MON & "_" & @YDAY & "_" & @HOUR & "_" & @MIN & "_" & @SEC

	; Close the document, save changes.
	$sSavepath = _LOCalc_DocClose($oDoc, True, $sSaveName)
	If @error Then _ERROR($oDoc, "Failed to close and save opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Calc Document was successfully saved to the following path: " & $sSavepath & @CRLF & _
			"Press OK to Delete it.")

	FileDelete($sSavepath)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
