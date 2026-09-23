#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oHandout
	Local $iLayout

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the handout Page's Object.
	$oHandout = _LOImpress_SlideHandoutGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Handout page's Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current handout page layout.
	$iLayout = _LOImpress_SlideHandoutLayout($oHandout)
	If @error Then _ERROR($oDoc, "Failed to retrieve page Layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page's current layout is (See UDF Constants): " & $iLayout & @CRLF & @CRLF & _
			"Press ok to change the Page's layout to $LOI_HANDOUT_LAYOUT_THREE_SLIDES. You can switch to that view to see that it worked.")

	; Change the Handout Page's layout to $LOI_HANDOUT_LAYOUT_THREE_SLIDES
	_LOImpress_SlideLayout($oHandout, $LOI_HANDOUT_LAYOUT_THREE_SLIDES)
	If @error Then _ERROR($oDoc, "Failed to modify Handout Page layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
