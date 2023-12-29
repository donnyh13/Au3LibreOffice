#include <GUIConstants.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $hExampleGUI, $hExecute, $hClose, $hExecuteCombo
	Local $sExecuteCommand

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "This is to demonstrate using _LOWriter_DocExecuteDispatch. Try clicking ""Execute"" on the small GUI " & _
			"window present on your screen. You can also click the drop down and try any other command. If you press ""Execute"" with ""uno:SwBackspace"" selected " & _
			"this smiley face will be backspaced.☻")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	$hExampleGUI = GUICreate("Doc Execute Example", 200, 60, -1, -1, -1, $WS_EX_TOPMOST)

	$hExecute = GUICtrlCreateButton("Execute", 5, 10, 60, 20)
	$hClose = GUICtrlCreateButton("Close", 5, 30, 60, 20)

	$hExecuteCombo = GUICtrlCreateCombo("", 70, 10, 120)

	; Fill the Combo with possible commands.
	GUICtrlSetData($hExecuteCombo, "uno:FullScreen|uno:ChangeCaseToUpper|uno:ChangeCaseToLower|uno:ResetAttributes|uno:SwBackspace|uno:Delete|" & _
			"uno:Paste|uno:PasteUnformatted|uno:Copy|uno:Cut|uno:SelectAll|uno:ZoomPlus|uno:ZoomMinus", "uno:SwBackspace")

	GUISetState(@SW_SHOW, $hExampleGUI)

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $hClose

				; Close the document.
				_LOWriter_DocClose($oDoc, False)
				If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

				GUIDelete($hExampleGUI)
				Exit

			Case $hExecute
				$sExecuteCommand = GUICtrlRead($hExecuteCombo)

				; Perform the requested Execute command.
				_LOWriter_DocExecuteDispatch($oDoc, $sExecuteCommand)
				If @error Then _ERROR("Failed to execute a dispatch command. Error:" & @error & " Extended:" & @extended)

				MsgBox($MB_OK, "Executed Command", "The command """ & $sExecuteCommand & """ was succesfully performed.")
		EndSwitch
	WEnd

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
