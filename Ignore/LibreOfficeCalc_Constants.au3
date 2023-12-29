#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Calc Constants for the Libre Office UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the Libre Office UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LOCCONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LOCCONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LOCCONST_SLEEP_DIV = 0

#Tidy_ILC_Pos=65

; Cell Type
Global Const _
		$LOC_CELL_TYPE_EMPTY = 0, _                             ; Cell is empty.
		$LOC_CELL_TYPE_VALUE = 1, _                             ; Cell contains a value.
		$LOC_CELL_TYPE_TEXT = 2, _                              ; Cell contains text.
		$LOC_CELL_TYPE_FORMULA = 3                              ; Cell contains a formula.

; Color in Long Color Format
Global Const _
		$LOC_COLOR_OFF = -1, _                                  ; Turn Color off, or to automatic mode.
		$LOC_COLOR_BLACK = 0, _                                 ; Black color.
		$LOC_COLOR_WHITE = 16777215, _                          ; White color.
		$LOC_COLOR_LGRAY = 11711154, _                          ; Light Gray color.
		$LOC_COLOR_GRAY = 8421504, _                            ; Gray color.
		$LOC_COLOR_DKGRAY = 3355443, _                          ; Dark Gray color.
		$LOC_COLOR_YELLOW = 16776960, _                         ; Yellow color.
		$LOC_COLOR_GOLD = 16760576, _                           ; Gold color.
		$LOC_COLOR_ORANGE = 16744448, _                         ; Orange color.
		$LOC_COLOR_BRICK = 16728064, _                          ; Brick color.
		$LOC_COLOR_RED = 16711680, _                            ; Red color.
		$LOC_COLOR_MAGENTA = 12517441, _                        ; Magenta color.
		$LOC_COLOR_PURPLE = 8388736, _                          ; Purple color.
		$LOC_COLOR_INDIGO = 5582989, _                          ; Indigo color.
		$LOC_COLOR_BLUE = 2777241, _                            ; Blue color.
		$LOC_COLOR_TEAL = 1410150, _                            ; Teal color.
		$LOC_COLOR_GREEN = 43315, _                             ; Green color.
		$LOC_COLOR_LIME = 8508442, _                            ; Lime color.
		$LOC_COLOR_BROWN = 9127187                              ; Brown color.

; Path Convert Constants.
Global Const _
		$LOC_PATHCONV_AUTO_RETURN = 0, _                        ; Automatically returns the opposite of the input path, determined by StringInStr search for either "File:///"(L.O.Office URL) or "[A-Z]:\" (Windows File Path).
		$LOC_PATHCONV_OFFICE_RETURN = 1, _                      ; Returns L.O. Office URL, even if the input is already in that format.
		$LOC_PATHCONV_PCPATH_RETURN = 2                         ; Returns Windows File Path, even if the input is already in that format.

; Zoom Type Constants
Global Const _
		$LOC_ZOOMTYPE_OPTIMAL = 0, _                            ; The page content width (excluding margins) at the current selection is fit into the view.
		$LOC_ZOOMTYPE_PAGE_WIDTH = 1, _                         ; The page width at the current selection is fit into the view.
		$LOC_ZOOMTYPE_ENTIRE_PAGE = 2, _                        ; A complete page of the document is fit into the view.
		$LOC_ZOOMTYPE_BY_VALUE = 3, _                           ; The Zoom property is relative, and set using Zoom Value.
		$LOC_ZOOMTYPE_PAGE_WIDTH_EXACT = 4                      ; The Page width at the current selection is fit into the view with the view ends exactly at the end of the page.
