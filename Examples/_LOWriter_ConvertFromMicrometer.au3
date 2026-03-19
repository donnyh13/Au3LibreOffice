
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $iInch_From_MicroM, $iCM_From_MicroM, $iMM_From_MicroM, $iPt_From_MicroM

	;Convert 2540 Micrometers to Inches.
	$iInch_From_MicroM = _LOWriter_ConvertFromMicrometer(2540)
	If (@error > 0) Then _ERROR("Failed to convert from Micrometers to Inch. Error:" & @error & " Extended:" & @extended)

	;Convert 2540 Micrometers to Centimeters.
	$iCM_From_MicroM = _LOWriter_ConvertFromMicrometer(Null, 2540)
	If (@error > 0) Then _ERROR("Failed to convert to Micrometers from Centimeter. Error:" & @error & " Extended:" & @extended)

	;Convert 2540 Micrometers to Millimeters.
	$iMM_From_MicroM = _LOWriter_ConvertFromMicrometer(Null, Null, 2540)
	If (@error > 0) Then _ERROR("Failed to convert to Micrometers from Millimeter. Error:" & @error & " Extended:" & @extended)

	;Convert 2540 Micrometers to Printer's Points.
	$iPt_From_MicroM = _LOWriter_ConvertFromMicrometer(Null, Null, Null, 2540)
	If (@error > 0) Then _ERROR("Failed to convert to Micrometers from Millimeter. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "2540 Micrometers converted to Inches = " & $iInch_From_MicroM & @CRLF & _
			"2540 Micrometers to CM = " & $iCM_From_MicroM & @CRLF & _
			"2540 Micrometers to MM = " & $iMM_From_MicroM & @CRLF & _
			"2540 Micrometers to Printer's Points = " & $iPt_From_MicroM & @CRLF & @CRLF & _
			"a Micrometer is 1000th of a centimeter.")

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



