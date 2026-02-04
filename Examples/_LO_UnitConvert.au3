#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"

Example()

Func Example()
	Local $iInch_To_HMM, $iCM_To_HMM, $iMM_To_HMM, $iPt_To_HMM
	Local $iInch_From_HMM, $iCM_From_HMM, $iMM_From_HMM, $iPt_From_HMM

	; Convert 1 Inch to Hundredth Millimeters.
	$iInch_To_HMM = _LO_UnitConvert(1, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR("Failed to convert to HMM from Inch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2.54 Centimeters to Hundredth Millimeters.
	$iCM_To_HMM = _LO_UnitConvert(2.54, $LO_CONVERT_UNIT_CM_HMM)
	If @error Then _ERROR("Failed to convert to HMM from Centimeter. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 25.40 Millimeters to Hundredth Millimeters.
	$iMM_To_HMM = _LO_UnitConvert(25.4, $LO_CONVERT_UNIT_MM_HMM)
	If @error Then _ERROR("Failed to convert to HMM from Millimeter. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 72 Printer's Points to Hundredth Millimeters.
	$iPt_To_HMM = _LO_UnitConvert(72, $LO_CONVERT_UNIT_PT_HMM)
	If @error Then _ERROR("Failed to convert to HMM from Printer's Points. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "1 Inch converted to Hundredths of a Millimeter (HMM) = " & $iInch_To_HMM & @CRLF & _
			"2.54 CM to Hundredths of a Millimeter (HMM) = " & $iCM_To_HMM & @CRLF & _
			"25.40 MM to Hundredths of a Millimeter (HMM) = " & $iMM_To_HMM & @CRLF & _
			"72 Printer's Points to Hundredths of a Millimeter (HMM) = " & $iPt_To_HMM)

	; Convert 2540 Hundredth Millimeters to Inches.
	$iInch_From_HMM = _LO_UnitConvert(2540, $LO_CONVERT_UNIT_HMM_INCH)
	If @error Then _ERROR("Failed to convert from HMM to Inch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2540 Hundredth Millimeters to Centimeters.
	$iCM_From_HMM = _LO_UnitConvert(2540, $LO_CONVERT_UNIT_HMM_CM)
	If @error Then _ERROR("Failed to convert from HMM to Centimeter. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2540 Hundredth Millimeters to Millimeters.
	$iMM_From_HMM = _LO_UnitConvert(2540, $LO_CONVERT_UNIT_HMM_MM)
	If @error Then _ERROR("Failed to convert from HMM to Millimeter. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2540 Hundredth Millimeters to Printer's Points.
	$iPt_From_HMM = _LO_UnitConvert(2540, $LO_CONVERT_UNIT_HMM_PT)
	If @error Then _ERROR("Failed to convert from HMM to Printer's Points. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "2540 HMM (100th Millimeters) converted to Inches = " & $iInch_From_HMM & @CRLF & _
			"2540 HMM (100th Millimeters) to CM = " & $iCM_From_HMM & @CRLF & _
			"2540 HMM (100th Millimeters) to MM = " & $iMM_From_HMM & @CRLF & _
			"2540 HMM (100th Millimeters) to Printer's Points = " & $iPt_From_HMM)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
