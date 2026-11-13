#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oLayer
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new layer.
	$oLayer = _LODraw_DocLayerInsert($oDoc, "AutoIt Testing", "", "", False)
	If @error Then _ERROR($oDoc, "Failed to insert a new layer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Layer's settings. Name = "AutoIt_Shapes", Title = "Test Layer", Description = "This is a layer inserted for testing.", Visibility to True, and
	; Printable to False, and Locked = True.
	_LODraw_DocLayerModify($oLayer, "AutoIt_Shapes", "Test Layer", "This is a layer inserted for testing.", True, False, True)
	If @error Then _ERROR($oDoc, "Failed to set Layer settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Layer settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_DocLayerModify($oLayer)
	If @error Then _ERROR($oDoc, "Failed to retrieve Layer settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Layer's Gradient settings are as follows: " & @CRLF & _
			"The Layer's name is: " & $avSettings[0] & @CRLF & _
			"The Layer's title is: " & $avSettings[1] & @CRLF & _
			"The Layer's description is: " & $avSettings[2] & @CRLF & _
			"Is the Layer visible? True/False: " & $avSettings[3] & @CRLF & _
			"Is the Layer printable? True/False: " & $avSettings[4] & @CRLF & _
			"Is the Layer locked? True/False: " & $avSettings[5])

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
