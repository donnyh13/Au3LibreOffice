#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2
	Local $iUserChoice
	Local $sDocName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created a blank L.O. Base Doc, I will now Connect to it and use the new Object returned to close it.")

	; Connect to the Current Document.
	$oDoc2 = _LOBase_DocConnect("", True)
	If (@error > 0) Or Not IsObj($oDoc2) Then _ERROR($oDoc, $oDoc2, "Failed to Connect to Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Doc Name.
	$sDocName = _LOBase_DocGetName($oDoc, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to retrieve Base Document name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iUserChoice = MsgBox($MB_YESNO, "Close?", "I have connected to the Document with the following title: " & $sDocName & _
			@CRLF & "Would you like to close it now?")

	If ($iUserChoice = $IDYES) Then
		; Close the document.
		_LOBase_DocClose($oDoc2, False)
		If @error Then _ERROR($oDoc, $oDoc2, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	EndIf
EndFunc

Func _ERROR($oDoc, $oDoc2, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsObj($oDoc2) Then _LOBase_DocClose($oDoc2, False)
	Exit
EndFunc
