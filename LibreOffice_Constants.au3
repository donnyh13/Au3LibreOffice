#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/reel /tcl=1
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

#Tidy_ILC_Pos=50

; Sleep Divisor $__LOCONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LOCONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LOCONST_SLEEP_DIV = 0

; RGB Color Integers
Global Const _
		$LO_COLOR_OFF = -1, _                    ; Turn Color off, or to automatic mode.
		$LO_COLOR_BLACK = 0, _                   ; Black color.
		$LO_COLOR_GREEN = 43315, _               ; Green color.
		$LO_COLOR_TEAL = 1410150, _              ; Teal color.
		$LO_COLOR_BLUE = 2777241, _              ; Blue color.
		$LO_COLOR_DKGRAY = 3355443, _            ; Dark Gray color.
		$LO_COLOR_INDIGO = 5582989, _            ; Indigo color.
		$LO_COLOR_PURPLE = 8388736, _            ; Purple color.
		$LO_COLOR_GRAY = 8421504, _              ; Gray color.
		$LO_COLOR_LIME = 8508442, _              ; Lime color.
		$LO_COLOR_BROWN = 9127187, _             ; Brown color.
		$LO_COLOR_LGRAY = 11711154, _            ; Light Gray color.
		$LO_COLOR_MAGENTA = 12517441, _          ; Magenta color.
		$LO_COLOR_RED = 16711680, _              ; Red color.
		$LO_COLOR_BRICK = 16728064, _            ; Brick color.
		$LO_COLOR_ORANGE = 16744448, _           ; Orange color.
		$LO_COLOR_GOLD = 16760576, _             ; Gold color.
		$LO_COLOR_YELLOW = 16776960, _           ; Yellow color.
		$LO_COLOR_WHITE = 16777215               ; White color.

; Conversion Constants.
Global Enum _
		$LO_CONVERT_UNIT_TWIPS_CM, _             ; 0 Convert from TWIPS (Twentieth of a Printer Point) To Centimeters.
		$LO_CONVERT_UNIT_TWIPS_INCH, _           ; 1 Convert from TWIPS (Twentieth of a Printer Point) To Inches.
		$LO_CONVERT_UNIT_TWIPS_HMM, _            ; 2 Convert from TWIPS(Twentieth of a Printer Point) To Hundredths of a Millimeter (HMM).
		$LO_CONVERT_UNIT_MM_HMM, _               ; 3 Convert from Millimeters To Hundredths of a Millimeter (HMM).
		$LO_CONVERT_UNIT_HMM_MM, _               ; 4 Convert from Hundredths of a Millimeter (HMM) To Millimeters.
		$LO_CONVERT_UNIT_CM_HMM, _               ; 5 Convert from Centimeters To Hundredths of a Millimeter (HMM).
		$LO_CONVERT_UNIT_HMM_CM, _               ; 6 Convert from Hundredths of a Millimeter (HMM) To Centimeters.
		$LO_CONVERT_UNIT_INCH_HMM, _             ; 7 Convert from Inches To Hundredths of a Millimeter (HMM).
		$LO_CONVERT_UNIT_HMM_INCH, _             ; 8 Convert from Hundredths of a Millimeter (HMM) To Inches.
		$LO_CONVERT_UNIT_PT_HMM, _               ; 9 Convert from Printers Point to Hundredths of a Millimeter (HMM).
		$LO_CONVERT_UNIT_HMM_PT                  ; 10 Convert from Hundredths of a Millimeter (HMM) to Printers Point.

; Path Convert Constants.
Global Const _
		$LO_PATHCONV_AUTO_RETURN = 0, _          ; Automatically returns the opposite of the input path, determined by StringInStr search for either "File:///"(L.O.Office URL) or "[A-Z]:\" (Windows File Path).
		$LO_PATHCONV_OFFICE_RETURN = 1, _        ; Returns L.O. Office URL, even if the input is already in that format.
		$LO_PATHCONV_PCPATH_RETURN = 2           ; Returns Windows File Path, even if the input is already in that format.

; Error Codes
Global Enum _
		$__LO_STATUS_SUCCESS, _                  ; 0 Function finished successfully.
		$__LO_STATUS_INPUT_ERROR, _              ; 1 Function encountered an input error.
		$__LO_STATUS_INIT_ERROR, _               ; 2 Function encountered an Initialization error.
		$__LO_STATUS_PROCESSING_ERROR, _         ; 3 Function encountered a Processing error.
		$__LO_STATUS_PROP_SETTING_ERROR, _       ; 4 Function encountered a Property setting error.
		$__LO_STATUS_PRINTER_RELATED_ERROR, _    ; 5 Function encountered a Printer related error.
		$__LO_STATUS_VER_ERROR                   ; 6 Function encountered a Version error.
