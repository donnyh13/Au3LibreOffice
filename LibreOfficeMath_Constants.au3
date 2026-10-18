#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice Math Constants for the LibreOffice UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various Math functions in the LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Note ..........: Descriptions for some Constants are taken from the LibreOffice SDK API documentation.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LOMCONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LOMCONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LOMCONST_SLEEP_DIV = 0

#Tidy_ILC_Pos=55

; Posture/Italic
Global Const _                                        ; com.sun.star.awt.FontSlant
		$LOM_CHAR_POSTURE_NONE = 0, _                 ; Specifies a font without slant.
		$LOM_CHAR_POSTURE_OBLIQUE = 1, _              ; Specifies an oblique font (slant not designed into the font).
		$LOM_CHAR_POSTURE_ITALIC = 2, _               ; Specifies an italic font (slant designed into the font).
		$LOM_CHAR_POSTURE_DONTKNOW = 3, _             ; Specifies a font with an unknown slant. For Read Only.
		$LOM_CHAR_POSTURE_REV_OBLIQUE = 4, _          ; Specifies a reverse oblique font (slant not designed into the font).
		$LOM_CHAR_POSTURE_REV_ITALIC = 5              ; Specifies a reverse italic font (slant designed into the font).

; Weight/Bold
Global Const _                                        ; com.sun.star.awt.FontWeight
		$LOM_CHAR_WEIGHT_DONT_KNOW = 0, _             ; The font weight is not specified/unknown. For Read Only.
		$LOM_CHAR_WEIGHT_THIN = 50, _                 ; A 50% (Thin) font weight.
		$LOM_CHAR_WEIGHT_ULTRA_LIGHT = 60, _          ; A 60% (Ultra Light) font weight.
		$LOM_CHAR_WEIGHT_LIGHT = 75, _                ; A 75% (Light) font weight.
		$LOM_CHAR_WEIGHT_SEMI_LIGHT = 90, _           ; A 90% (Semi-Light) font weight.
		$LOM_CHAR_WEIGHT_NORMAL = 100, _              ; A 100% (Normal) font weight.
		$LOM_CHAR_WEIGHT_SEMI_BOLD = 110, _           ; A 110% (Semi-Bold) font weight.
		$LOM_CHAR_WEIGHT_BOLD = 150, _                ; A 150% (Bold) font weight.
		$LOM_CHAR_WEIGHT_ULTRA_BOLD = 175, _          ; A 175% (Ultra-Bold) font weight.
		$LOM_CHAR_WEIGHT_BLACK = 200                  ; A 200% (Black) font weight.

; Font kind constants
Global Enum _
		$LOM_FMT_FONT_KIND_MATH, _                    ; 0 The Font applied to Mathematical symbols in the formula.
		$LOM_FMT_FONT_KIND_VARIABLES, _               ; 1 The Font applied to Variables in the formula.
		$LOM_FMT_FONT_KIND_FUNCTIONS, _               ; 2 The Font applied to Functions in the formula.
		$LOM_FMT_FONT_KIND_NUMBERS, _                 ; 3 The Font applied to Numbers in the formula.
		$LOM_FMT_FONT_KIND_TEXT, _                    ; 4 The Font applied to Text in the formula.
		$LOM_FMT_FONT_KIND_SERIF, _                   ; 5 The Serif Custom Font.
		$LOM_FMT_FONT_KIND_SANS_SERIF, _              ; 6 The Sans-Serif Custom Font.
		$LOM_FMT_FONT_KIND_FIXED_WIDTH                ; 7 The Fixed Width Custom Font.

; Formula Alignment
Global Const _                                        ; com.sun.star.style.HorizontalAlignment
		$LOM_FORMULA_ALIGN_HORI_LEFT = 0, _           ; The Formula is aligned to the Left.
		$LOM_FORMULA_ALIGN_HORI_CENTER = 1, _         ; The Formula is aligned in the Center.
		$LOM_FORMULA_ALIGN_HORI_RIGHT = 2             ; The Formula is aligned to the Right.
