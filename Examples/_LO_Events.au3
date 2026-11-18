#include <Array.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"

Example()

Func Example()
	Local $asEvents

	; Retrieve an Array of All events
	$asEvents = _LO_Events()
	If @error Then _ERROR("Failed to retrieve Array of events. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & @extended & " events found. I will now display the results using _ArrayDisplay. There will be two columns, " & @CRLF & _
			"-the first column contains the event name, " & @CRLF & _
			"-the second column contains the macro assigned to it, if any.")

	_ArrayDisplay($asEvents, Default, Default, 0, Default, "Event Name|Macro")

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR("Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
