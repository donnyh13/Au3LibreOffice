#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc
	Local $sUndo

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; For some reason actions done using COM/UNO commands doesn't register in the Undo stack, so I have to simulate a user action using a Dispatch.
	; Insert a new page.
	_LODraw_DocExecuteDispatch($oDoc, "uno:InsertPageQuick")
	If @error Then _ERROR($oDoc, "Failed to execute a dispatch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Fixed Date Field.
	_LODraw_DocExecuteDispatch($oDoc, "uno:InsertDateFieldFix")
	If @error Then _ERROR($oDoc, "Failed to execute a dispatch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the next available Undo action title.
	$sUndo = _LODraw_DocUndoCurActionTitle($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve next undo action title. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The next undo action title is: " & $sUndo & @CRLF & "Press ok to perform it.")

	; Perform one undo action.
	_LODraw_DocUndo($oDoc)
	If @error Then _ERROR($oDoc, "Failed to perform an undo action. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
