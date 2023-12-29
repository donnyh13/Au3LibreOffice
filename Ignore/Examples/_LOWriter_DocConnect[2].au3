#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2
	Local $iUserChoice
	Local $sDocName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have created a blank L.O. Writer Doc, I will now Connect to it and use the new Object returned to close it.")

	; Connect to the Current Document.
	$oDoc2 = _LOWriter_DocConnect("", True)
	If (@error > 0) Or Not IsObj($oDoc2) Then _ERROR("Failed to Connect to Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve Doc Name.
	$sDocName = _LOWriter_DocGetName($oDoc, False)
	If @error Then _ERROR("Failed to retrieve Writer Document name. Error:" & @error & " Extended:" & @extended)

	$iUserChoice = MsgBox($MB_YESNO, "Close?", "I have connected to the Document with the following title: " & $sDocName & _
			@CRLF & "Would you like to close it now?")

	If ($iUserChoice = $IDYES) Then
		; Close the document.
		_LOWriter_DocClose($oDoc2, False)
		If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
	EndIf

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
