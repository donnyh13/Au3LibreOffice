#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $sRedos
	Local $asRedo

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; For some reason actions done using COM/UNO commands doesn't register in the Undo stack, so I have to simulate a user action using a Dispatch.
	; Insert a new page.
	_LOImpress_DocExecuteDispatch($oDoc, "uno:InsertPageQuick")
	If @error Then _ERROR($oDoc, "Failed to execute a dispatch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Fixed Date Field.
	_LOImpress_DocExecuteDispatch($oDoc, "uno:InsertDateFieldFix")
	If @error Then _ERROR($oDoc, "Failed to execute a dispatch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Perform one undo action.
	_LOImpress_DocUndo($oDoc)
	If @error Then _ERROR($oDoc, "Failed to perform an undo action. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of available Redo action titles.
	$asRedo = _LOImpress_DocRedoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Redo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $sRedo In $asRedo
		$sRedos &= $sRedo & @CRLF
	Next

	; Display the available Redo action titles.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Redo Actions are:" & @CRLF & $sRedos)

	; Clear the Redo Action list.
	_LOImpress_DocRedoClear($oDoc)
	If @error Then _ERROR($oDoc, "Failed to clear Redo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have cleared the Redo Actions list. I will retrieve the available Redo Actions list again and show that it is now empty.")

	; Retrieve an array of available Redo action titles again.
	$asRedo = _LOImpress_DocRedoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Redo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sRedos = ""

	For $sRedo In $asRedo
		$sRedos &= $sRedo & @CRLF
	Next

	; Display the available Redo action titles again, if any.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Redo Actions are:" & @CRLF & $sRedos)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
