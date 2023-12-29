#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Constants for the Libre Office UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the Libre Office UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

#Tidy_ILC_Pos=65
; Error Codes
Global Enum _
		$__LO_STATUS_SUCCESS = 0, _                            ; 0 Function finished successfully.
		$__LO_STATUS_INPUT_ERROR, _                            ; 1 Function encountered a input error.
		$__LO_STATUS_INIT_ERROR, _                             ; 2 Function encountered a Initialization error.
		$__LO_STATUS_PROCESSING_ERROR, _                       ; 3 Function encountered a Processing error.
		$__LO_STATUS_PROP_SETTING_ERROR, _                     ; 4 Function encountered a Property setting error.
		$__LO_STATUS_DOC_ERROR, _                              ; 5 Function encountered a Document related error.
		$__LO_STATUS_PRINTER_RELATED_ERROR, _                  ; 6 Function encountered a Printer related error.
		$__LO_STATUS_VER_ERROR                                 ; 7 Function encountered a Version error.

; Conversion Constants.
Global Enum _
		$__LOCONST_CONVERT_TWIPS_CM, _                         ; 0 Convert from TWIPS (Twentieth of a Printer Point) To Centimeters.
		$__LOCONST_CONVERT_TWIPS_INCH, _                       ; 1 Convert from TWIPS (Twentieth of a Printer Point) To Inches.
		$__LOCONST_CONVERT_TWIPS_UM, _                         ; 2 Convert from TWIPS(Twentieth of a Printer Point) To Micrometer(100th of a millimeter).
		$__LOCONST_CONVERT_MM_UM, _                            ; 3 Convert from Millimeters To Micrometer (100th of a millimeter).
		$__LOCONST_CONVERT_UM_MM, _                            ; 4 Convert from Micrometer (100th of a millimeter) To Millimeters.
		$__LOCONST_CONVERT_CM_UM, _                            ; 5 Convert from Centimeters To Micrometer (100th of a millimeter).
		$__LOCONST_CONVERT_UM_CM, _                            ; 6 Convert from Micrometer (100th of a millimeter) To Centimeters.
		$__LOCONST_CONVERT_INCH_UM, _                          ; 7 Convert from Inches To Micrometer (100th of a millimeter).
		$__LOCONST_CONVERT_UM_INCH, _                          ; 8 Convert from Micrometer (100th of a millimeter) To Inches.
		$__LOCONST_CONVERT_PT_UM, _                            ; 9 Convert from Printers Point to Micrometers.
		$__LOCONST_CONVERT_UM_PT                               ; 10 Convert from Micrometers to Printers Point.
