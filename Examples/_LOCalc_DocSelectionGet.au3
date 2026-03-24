#include <GUIConstants.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"
Example()

Func Example()
	Local $oDoc
	Local $vReturn
	Local $hGui, $hTest, $hClose, $hEdit

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$hGui = GUICreate("Calc Selection Example.", 200, 400, -1, -1, -1, $WS_EX_TOPMOST)
	$hTest = GUICtrlCreateButton("Test", 5, 70, 50, 30)
	GUICtrlSetFont($hTest, 10, 600)
	$hClose = GUICtrlCreateButton("Close", 145, 70, 50, 30)
	GUICtrlSetFont($hClose, 10, 600)
	$hEdit = GUICtrlCreateEdit("", 1, 105, 198, 290, $ES_READONLY)
	GUICtrlSetFont($hEdit, 12, 600)
	GUICtrlCreateLabel("Select something in the Calc document, and then press ""Test"".", 1, 1, 195, 65)
	GUICtrlSetFont(-1, 12, 600)

	GUISetState(@SW_SHOW, $hGui)

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $hClose
				ExitLoop

			Case $hTest
				; Retrieve the current selection in the Document.
				$vReturn = _LOCalc_DocSelectionGet($oDoc)
				If @error Then _ERROR($oDoc, "Failed to retrieve the current selection. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

				Switch @extended
					Case 0     ; Single Cell
						GUICtrlSetData($hEdit, "The current selection is a single Cell." & @CRLF & @CRLF & _
								"The Address of the selected Cell is: " & @CRLF & _LOCalc_RangeGetAddressAsName($vReturn))

					Case 1     ; Single Cell Range
						GUICtrlSetData($hEdit, "The current selection is a single Cell Range." & @CRLF & @CRLF & _
								"The Address of the Range selected is: " & @CRLF & _LOCalc_RangeGetAddressAsName($vReturn))

					Case Else     ; Multiple Cell Ranges
						GUICtrlSetData($hEdit, "The current selection is several Cell Ranges." & @CRLF & @CRLF & _
								"The Addresses of the Ranges currently selected are: " & @CRLF)

						For $I = 0 To UBound($vReturn) - 1
							GUICtrlSetData($hEdit, _LOCalc_RangeGetAddressAsName($vReturn[$I]) & @CRLF, 1)
						Next
				EndSwitch
		EndSwitch
	WEnd

	GUIDelete($hGui)

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
