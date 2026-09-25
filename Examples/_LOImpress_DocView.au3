#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $iView

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current view mode.
	$iView = _LOImpress_DocView($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current view mode. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your current view mode is (See UDF Constants): " & $iView & @CRLF & _
			"I will now set the view mode to: $LOI_PAGE_VIEW_SLIDE_SORTER.")

	; View mode to $LOI_PAGE_VIEW_SLIDE_SORTER.
	_LOImpress_DocView($oDoc, $LOI_PAGE_VIEW_SLIDE_SORTER)
	If @error Then _ERROR($oDoc, "Failed to set view mode. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now set view mode to $LOI_PAGE_VIEW_MASTER_NOTES.")

	; Set the zoom to the Zoom type of $LOI_PAGE_VIEW_MASTER_NOTES.
	_LOImpress_DocView($oDoc, $LOI_PAGE_VIEW_MASTER_NOTES)
	If @error Then _ERROR($oDoc, "Failed to set zoom value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
