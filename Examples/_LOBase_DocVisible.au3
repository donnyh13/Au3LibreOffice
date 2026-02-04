#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok and I will make the new document I just opened invisible.")

	; Make the document invisible by setting visible to False
	_LOBase_DocVisible($oDoc, False)
	If (@error > 0) Then _ERROR($oDoc, "Failed to change Document visibility settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test if the document is Visible
	$bReturn = _LOBase_DocVisible($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document status. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the document currently visible? True/False: " & $bReturn & @CRLF & @CRLF & _
			"Press Ok to make the document visible again.")

	; Make the document visible by setting visible to True
	_LOBase_DocVisible($oDoc, True)
	If (@error > 0) Then _ERROR($oDoc, "Failed to change Document visibility settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test if the document is Visible
	$bReturn = _LOBase_DocVisible($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document status. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the document now visible? True/False: " & $bReturn)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	Exit
EndFunc
