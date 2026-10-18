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

; Font kind constants
Global Enum _
		$LOM_FMT_FONT_KIND_MATH, _                    ; The Font applied to Mathematical symbols in the formula.
		$LOM_FMT_FONT_KIND_VARIABLES, _               ; The Font applied to Variables in the formula.
		$LOM_FMT_FONT_KIND_FUNCTIONS, _               ; The Font applied to Functions in the formula.
		$LOM_FMT_FONT_KIND_NUMBERS, _                 ; The Font applied to Numbers in the formula.
		$LOM_FMT_FONT_KIND_TEXT, _                    ; The Font applied to Text in the formula.
		$LOM_FMT_FONT_KIND_SERIF, _                   ; The Serif Custom Font.
		$LOM_FMT_FONT_KIND_SANS_SERIF, _              ; The Sans-Serif Custom Font.
		$LOM_FMT_FONT_KIND_FIXED_WIDTH                ; The Fixed Width Custom Font.

; Formula Alignment
Global Const _                                        ; com.sun.star.style.HorizontalAlignment
		$LOM_FORMULA_ALIGN_HORI_LEFT = 0, _           ; The Formula is aligned to the Left.
		$LOM_FORMULA_ALIGN_HORI_CENTER = 1, _         ; The Formula is aligned in the Center.
		$LOM_FORMULA_ALIGN_HORI_RIGHT = 2             ; The Formula is aligned to the Right.
