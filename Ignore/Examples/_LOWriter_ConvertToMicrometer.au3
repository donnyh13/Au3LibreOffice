#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $iInch_To_MicroM, $iCM_To_MicroM, $iMM_To_MicroM, $iPt_To_MicroM

	; Convert 1 Inch to Micrometers.
	$iInch_To_MicroM = _LOWriter_ConvertToMicrometer(1)
	If @error Then _ERROR("Failed to convert to Micrometers from Inch. Error:" & @error & " Extended:" & @extended)

	; Convert 2.54 Centimeters to Micrometers.
	$iCM_To_MicroM = _LOWriter_ConvertToMicrometer(Null, 2.54)
	If @error Then _ERROR("Failed to convert to Micrometers from Centimeter. Error:" & @error & " Extended:" & @extended)

	; Convert 25.40 Millimeters to Micrometers.
	$iMM_To_MicroM = _LOWriter_ConvertToMicrometer(Null, Null, 25.4)
	If @error Then _ERROR("Failed to convert to Micrometers from Millimeter. Error:" & @error & " Extended:" & @extended)

	; Convert 72 Printer's Points to Micrometers.
	$iPt_To_MicroM = _LOWriter_ConvertToMicrometer(Null, Null, Null, 72)
	If @error Then _ERROR("Failed to convert to Micrometers from Millimeter. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "1 Inch converted to Micrometers = " & $iInch_To_MicroM & @CRLF & _
			"2.54 Cm to Micrometers = " & $iCM_To_MicroM & @CRLF & _
			"25.40 MM to Micrometers = " & $iMM_To_MicroM & @CRLF & _
			"72 Printer's Points to Micrometers = " & $iPt_To_MicroM & @CRLF & @CRLF & _
			"a Micrometer is 1000th of a centimeter.")

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
