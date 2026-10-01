#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc
	Local $iZoom
	Local $aiArray[0]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current zoom settings. Return value will be in order of function parameters.
	$aiArray = _LODraw_DocZoom($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current zoom settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iZoom = Int($aiArray[1] * .75) ; Set my new zoom value to 75% of the current zoom value.

	; Zoom cannot be less than 20% or greater than 600%, if my value is outside of this, set it to 140%
	If ($iZoom < 20) Or ($iZoom > 600) Then $iZoom = 140

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your current zoom value is: " & $aiArray[1] & "%. The Zoom type currently is: " & $aiArray[0] & @CRLF & _
			". I will now set the zoom value to: " & $iZoom & "%.")

	; Skip zoom type and set the zoom to my new value.
	_LODraw_DocZoom($oDoc, Null, $iZoom)
	If @error Then _ERROR($oDoc, "Failed to set zoom value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current zoom value again.
	$aiArray = _LODraw_DocZoom($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current zoom value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your new zoom value is: " & $aiArray[1] & "%. And the Zoom type is now: " & $aiArray[0] & @CRLF & _
			" I will now set zoom type to $LOD_ZOOMTYPE_ENTIRE_PAGE.")

	; Set the zoom to the Zoom type of $LOD_ZOOMTYPE_ENTIRE_PAGE.
	_LODraw_DocZoom($oDoc, $LOD_ZOOMTYPE_ENTIRE_PAGE)
	If @error Then _ERROR($oDoc, "Failed to set zoom value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
