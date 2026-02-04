#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Calc Constants for the Libre Office UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the Libre Office UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Note ..........: Descriptions for some Constants are taken from the LibreOffice SDK API documentation.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LOCCONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LOCCONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LOCCONST_SLEEP_DIV = 0

#Tidy_ILC_Pos=65

; Fill Style Type Constants
Global Enum _                                                   ; com.sun.star.drawing.FillStyle
		$LOC_AREA_FILL_STYLE_OFF, _                             ; 0 Fill Style is off.
		$LOC_AREA_FILL_STYLE_SOLID, _                           ; 1 Fill Style is a solid color.
		$LOC_AREA_FILL_STYLE_GRADIENT, _                        ; 2 Fill Style is a gradient color.
		$LOC_AREA_FILL_STYLE_HATCH, _                           ; 3 Fill Style is a Hatch style color.
		$LOC_AREA_FILL_STYLE_BITMAP                             ; 4 Fill Style is a Bitmap.

; Border Style
Global Const _                                                  ; com.sun.star.table.BorderLineStyle
		$LOC_BORDERSTYLE_NONE = 0x7FFF, _                       ; No border line.
		$LOC_BORDERSTYLE_SOLID = 0, _                           ; Solid border line.
		$LOC_BORDERSTYLE_DOTTED = 1, _                          ; Dotted border line.
		$LOC_BORDERSTYLE_DASHED = 2, _                          ; Dashed border line.
		$LOC_BORDERSTYLE_DOUBLE = 3, _                          ; Double border line.
		$LOC_BORDERSTYLE_THINTHICK_SMALLGAP = 4, _              ; Double border line with a thin line outside and a thick line inside separated by a small gap.
		$LOC_BORDERSTYLE_THINTHICK_MEDIUMGAP = 5, _             ; Double border line with a thin line outside and a thick line inside separated by a medium gap.
		$LOC_BORDERSTYLE_THINTHICK_LARGEGAP = 6, _              ; Double border line with a thin line outside and a thick line inside separated by a large gap.
		$LOC_BORDERSTYLE_THICKTHIN_SMALLGAP = 7, _              ; Double border line with a thick line outside and a thin line inside separated by a small gap.
		$LOC_BORDERSTYLE_THICKTHIN_MEDIUMGAP = 8, _             ; Double border line with a thick line outside and a thin line inside separated by a medium gap.
		$LOC_BORDERSTYLE_THICKTHIN_LARGEGAP = 9, _              ; Double border line with a thick line outside and a thin line inside separated by a large gap.
		$LOC_BORDERSTYLE_EMBOSSED = 10, _                       ; 3D embossed border line.
		$LOC_BORDERSTYLE_ENGRAVED = 11, _                       ; 3D engraved border line.
		$LOC_BORDERSTYLE_OUTSET = 12, _                         ; Outset border line.
		$LOC_BORDERSTYLE_INSET = 13, _                          ; Inset border line.
		$LOC_BORDERSTYLE_FINE_DASHED = 14, _                    ; Finely dashed border line.
		$LOC_BORDERSTYLE_DOUBLE_THIN = 15, _                    ; Double border line consisting of two fixed thin lines separated by a variable gap.
		$LOC_BORDERSTYLE_DASH_DOT = 16, _                       ; Line consisting of a repetition of one dash and one dot.
		$LOC_BORDERSTYLE_DASH_DOT_DOT = 17                      ; Line consisting of a repetition of one dash and 2 dots.

; Border Width
Global Const _
		$LOC_BORDERWIDTH_HAIRLINE = 2, _                        ; Hairline Border line width.
		$LOC_BORDERWIDTH_VERY_THIN = 18, _                      ; Very Thin Border line width.
		$LOC_BORDERWIDTH_THIN = 26, _                           ; Thin Border line width.
		$LOC_BORDERWIDTH_MEDIUM = 53, _                         ; Medium Border line width.
		$LOC_BORDERWIDTH_THICK = 79, _                          ; Thick Border line width.
		$LOC_BORDERWIDTH_EXTRA_THICK = 159                      ; Extra Thick Border line width.

; Cell Content Horizontal Alignment
Global Const _                                                  ; com.sun.star.table.CellHoriJustify
		$LOC_CELL_ALIGN_HORI_DEFAULT = 0, _                     ; The default alignment is used (left for numbers, right for text).
		$LOC_CELL_ALIGN_HORI_LEFT = 1, _                        ; The contents are printed from left to right.
		$LOC_CELL_ALIGN_HORI_CENTER = 2, _                      ; The contents are horizontally centered.
		$LOC_CELL_ALIGN_HORI_RIGHT = 3, _                       ; The contents are aligned to the right edge of the cell.
		$LOC_CELL_ALIGN_HORI_JUSTIFIED = 4, _                   ; The contents are justified to the cell width.
		$LOC_CELL_ALIGN_HORI_FILLED = 5, _                      ; The contents are repeated to fill the cell.
		$LOC_CELL_ALIGN_HORI_DISTRIBUTED = 6                    ; The contents are evenly aligned across the whole cell. Unlike Justified, it justifies the very last line of text, too.

; Cell Content Vertical Alignment
Global Const _                                                  ; com.sun.star.table.CellVertJustify2
		$LOC_CELL_ALIGN_VERT_DEFAULT = 0, _                     ; The default alignment is used.
		$LOC_CELL_ALIGN_VERT_TOP = 1, _                         ; The contents are aligned with the upper edge of the cell.
		$LOC_CELL_ALIGN_VERT_MIDDLE = 2, _                      ; The contents are aligned to the vertical middle of the cell.
		$LOC_CELL_ALIGN_VERT_BOTTOM = 3, _                      ; The contents are aligned to the lower edge of the cell.
		$LOC_CELL_ALIGN_VERT_JUSTIFIED = 4, _                   ; The contents are justified to the cell height.
		$LOC_CELL_ALIGN_VERT_DISTRIBUTED = 5                    ; The same as Justified, unless the text orientation is vertical. Then it behaves similarly to the horizontal Distributed setting, i.e. the very last line is justified, too.

; Cell Delete Mode Constants
Global Const _                                                  ; com.sun.star.sheet.CellDeleteMode
		$LOC_CELL_DELETE_MODE_NONE = 0, _                       ; No cells are moved -- Nothing happens.
		$LOC_CELL_DELETE_MODE_UP = 1, _                         ; The cells below the inserted Cells are moved up.
		$LOC_CELL_DELETE_MODE_LEFT = 2, _                       ; The cells to the right of the inserted cells are moved left.
		$LOC_CELL_DELETE_MODE_ROWS = 3, _                       ; Entire rows below the inserted cells are moved up.
		$LOC_CELL_DELETE_MODE_COLUMNS = 4                       ; Entire columns to the right of the inserted cells are moved left.

; Cell Content Type Flag Constants
Global Const _                                                  ; com.sun.star.sheet.CellFlags
		$LOC_CELL_FLAG_VALUE = 1, _                             ; Cell Contents that are numeric values but are not formatted as dates or times.
		$LOC_CELL_FLAG_DATE_TIME = 2, _                         ; Cell Contents that are numeric values that have a date or time number format.
		$LOC_CELL_FLAG_STRING = 4, _                            ; Cell Contents that are Strings.
		$LOC_CELL_FLAG_ANNOTATION = 8, _                        ; Cell Annotations
		$LOC_CELL_FLAG_FORMULA = 16, _                          ; Cell Contents that are Formulas.
		$LOC_CELL_FLAG_HARD_ATTR = 32, _                        ; Cells with explicit formatting, but not the formatting which is applied implicitly through styles.
		$LOC_CELL_FLAG_STYLES = 64, _                           ; Cells with Styles applied.
		$LOC_CELL_FLAG_OBJECTS = 128, _                         ; Cell Contents that are Drawing Objects.
		$LOC_CELL_FLAG_EDIT_ATTR = 256, _                       ; Cells containing formatting within parts of the cell contents.
		$LOC_CELL_FLAG_FORMATTED = 512, _                       ; Cells with formatting within the cells or cells with more than one paragraph within the cells.
		$LOC_CELL_FLAG_ALL = 1023                               ; All flags listed above.

; Cell Insert Mode Constants
Global Const _                                                  ; com.sun.star.sheet.CellInsertMode
		$LOC_CELL_INSERT_MODE_NONE = 0, _                       ; No cells are moved -- Nothing happens.
		$LOC_CELL_INSERT_MODE_DOWN = 1, _                       ; The cells below the inserted Cells are moved down.
		$LOC_CELL_INSERT_MODE_RIGHT = 2, _                      ; The cells to the right of the inserted cells are moved right.
		$LOC_CELL_INSERT_MODE_ROWS = 3, _                       ; Entire rows below the inserted cells are moved down.
		$LOC_CELL_INSERT_MODE_COLUMNS = 4                       ; Entire columns to the right of the inserted cells are moved right.

; Cell Content Rotation Reference
Global Const _                                                  ; com.sun.star.table.CellVertJustify2
		$LOC_CELL_ROTATE_REF_LOWER_CELL_BORDER = 0, _           ; Writes the rotated text from the bottom cell edge outwards.
		$LOC_CELL_ROTATE_REF_UPPER_CELL_BORDER = 1, _           ; Writes the rotated text from the top cell edge outwards.
		$LOC_CELL_ROTATE_REF_INSIDE_CELLS = 3                   ; Writes the rotated text only within the cell.

; Cell Type
Global Const _                                                  ; com.sun.star.table.CellContentType
		$LOC_CELL_TYPE_EMPTY = 0, _                             ; Cell is empty.
		$LOC_CELL_TYPE_VALUE = 1, _                             ; Cell contains a value.
		$LOC_CELL_TYPE_TEXT = 2, _                              ; Cell contains text.
		$LOC_CELL_TYPE_FORMULA = 3                              ; Cell contains a formula.

; Comment Text Anchor Position
Global Enum _
		$LOC_COMMENT_ANCHOR_TOP_LEFT, _                         ; The comment text is anchored in the Upper-Left corner of the comment box.
		$LOC_COMMENT_ANCHOR_TOP_CENTER, _                       ; The comment text is anchored in the Upper-Center of the comment box.
		$LOC_COMMENT_ANCHOR_TOP_RIGHT, _                        ; The comment text is anchored in the Upper-Right of the comment box.
		$LOC_COMMENT_ANCHOR_MIDDLE_LEFT, _                      ; The comment text is anchored in the Middle-Left corner of the comment box.
		$LOC_COMMENT_ANCHOR_MIDDLE_CENTER, _                    ; The comment text is anchored in the Middle-Center of the comment box.
		$LOC_COMMENT_ANCHOR_MIDDLE_RIGHT, _                     ; The comment text is anchored in the Middle-Right of the comment box.
		$LOC_COMMENT_ANCHOR_BOTTOM_LEFT, _                      ; The comment text is anchored in the Lower-Left corner of the comment box.
		$LOC_COMMENT_ANCHOR_BOTTOM_CENTER, _                    ; The comment text is anchored in the Lower-Center of the comment box.
		$LOC_COMMENT_ANCHOR_BOTTOM_RIGHT                        ; The comment text is anchored in the Lower-Right of the comment box.

; Comment Animation Direction
Global Const _                                                  ; com.sun.star.drawing.TextAnimationDirection
		$LOC_COMMENT_ANIMATION_DIR_LEFT = 0, _                  ; The Text moves towards the Left.
		$LOC_COMMENT_ANIMATION_DIR_RIGHT = 1, _                 ; The Text moves towards the Right.
		$LOC_COMMENT_ANIMATION_DIR_UP = 2, _                    ; The Text moves towards the Top.
		$LOC_COMMENT_ANIMATION_DIR_DOWN = 3                     ; The Text moves towards the Bottom.

; Comment Animation Kind
Global Const _                                                  ; com.sun.star.drawing.TextAnimationKind
		$LOC_COMMENT_ANIMATION_KIND_NONE = 0, _                 ; The Comment Text is not animated.
		$LOC_COMMENT_ANIMATION_KIND_BLINK = 1, _                ; The Comment Text has a blinking animation.
		$LOC_COMMENT_ANIMATION_KIND_SCROLL_THROUGH = 2, _       ; The Comment Text has a Scrolling animation.
		$LOC_COMMENT_ANIMATION_KIND_SCROLL_ALTERNATE = 3, _     ; The Comment Text has a Scrolling back and forth animation.
		$LOC_COMMENT_ANIMATION_KIND_SCROLL_IN = 4               ; The Comment Text Scrolls in animation.

; Comment Connector Horizontal Alignment
Global Const _
		$LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_TOP = 0, _          ; Align the Connector line to the Top of the Comment's left or right side, when Connector position is set to Horizontal.
		$LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_MIDDLE = 5000, _    ; Align the Connector line to the Middle of the Comment's left or right side, when Connector position is set to Horizontal.
		$LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_BOTTOM = 10000      ; Align the Connector line to the Bottom of the Comment's left or right side, when Connector position is set to Horizontal.

; Comment Connector Vertical Alignment
Global Const _
		$LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_LEFT = 0, _         ; Align the Connector line to the Left of the Comment's top or bottom side, when Connector position is set to Vertical.
		$LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_CENTER = 5000, _    ; Align the Connector line to the Center of the Comment's top or bottom side, when Connector position is set to Vertical.
		$LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_RIGHT = 10000       ; Align the Connector line to the Right of the Comment's top or bottom side, when Connector position is set to Vertical.

; Comment Connector Line Position
Global Const _                                                  ; com.sun.star.drawing.CaptionEscapeDirection
		$LOC_COMMENT_CALLOUT_EXT_HORI = 0, _                    ; The Connector line extends Horizontally from the Comment.
		$LOC_COMMENT_CALLOUT_EXT_VERT = 1, _                    ; The Connector line extends Vertically from the Comment.
		$LOC_COMMENT_CALLOUT_EXT_OPTIMAL = 2, _                 ; The Connector line extends from the optimal position of the Comment.
		$LOC_COMMENT_CALLOUT_EXT_FROM_LEFT = 3, _               ; The Connector line extends from the left of the Comment.
		$LOC_COMMENT_CALLOUT_EXT_FROM_TOP = 4                   ; The Connector line extends from the top of the Comment.

; Comment Connector Line Style
Global Const _                                                  ; com.sun.star.drawing.CaptionType
		$LOC_COMMENT_CALLOUT_STYLE_STRAIGHT = 0, _              ; The connecting line from the comment to the cell is straight.
		$LOC_COMMENT_CALLOUT_STYLE_ANGLED = 1, _                ; The connecting line from the comment to the cell is angled.
		$LOC_COMMENT_CALLOUT_STYLE_ANGLED_CONNECTOR = 2         ; The connecting line from the comment to the cell is angled and connected.

; Arrowhead Type Constants
Global Enum _
		$LOC_COMMENT_LINE_ARROW_TYPE_NONE, _                    ; 0 -- No Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_ARROW_SHORT, _             ; 1 --Short Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CONCAVE_SHORT, _           ; 2 -- Short Concave Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_ARROW, _                   ; 3 -- Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_TRIANGLE, _                ; 4 -- Triangle Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CONCAVE, _                 ; 5 -- Concave Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_ARROW_LARGE, _             ; 6 -- Large Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CIRCLE, _                  ; 7 -- Circle Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_SQUARE, _                  ; 8 -- Square Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_SQUARE_45, _               ; 9 -- Square Arrow head rotated 45 degrees.
		$LOC_COMMENT_LINE_ARROW_TYPE_DIAMOND, _                 ; 10 -- Diamond Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_HALF_CIRCLE, _             ; 11 -- Half Circle Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_DIMENSIONAL_LINES, _       ; 12 -- Dimension Lines head.
		$LOC_COMMENT_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW, _  ; 13 -- Dimension Line Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_DIMENSION_LINE, _          ; 14 -- Dimension Line head.
		$LOC_COMMENT_LINE_ARROW_TYPE_LINE_SHORT, _              ; 15 -- Short Line head.
		$LOC_COMMENT_LINE_ARROW_TYPE_LINE, _                    ; 16 -- Line head.
		$LOC_COMMENT_LINE_ARROW_TYPE_TRIANGLE_UNFILLED, _       ; 17 -- Unfilled Triangle Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_DIAMOND_UNFILLED, _        ; 18 -- Unfilled Diamond Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CIRCLE_UNFILLED, _         ; 19 -- Unfilled Circle Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, _      ; 20 -- Unfilled Square Arrow head, rotated 45 degrees.
		$LOC_COMMENT_LINE_ARROW_TYPE_SQUARE_UNFILLED, _         ; 21 -- Unfilled Square Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED, _    ; 22 -- Unfilled Half Circle Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_HALF_ARROW_LEFT, _         ; 23 -- Half Arrow left Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_HALF_ARROW_RIGHT, _        ; 24 -- Half Arrow right Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_REVERSED_ARROW, _          ; 25 -- Reversed Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_DOUBLE_ARROW, _            ; 26 -- Double Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CF_ONE, _                  ; 27 -- CF One Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CF_ONLY_ONE, _             ; 28 -- CF Only One Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CF_MANY, _                 ; 29 -- CF Many Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CF_MANY_ONE, _             ; 30 -- CF Many One Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CF_ZERO_ONE, _             ; 31 -- CF Zero One Arrow head.
		$LOC_COMMENT_LINE_ARROW_TYPE_CF_ZERO_MANY               ; 32 -- CF Zero Many Arrow head.

; Shape Line End Cap Constants.
Global Const _                                                  ; com.sun.star.drawing.LineCap
		$LOC_COMMENT_LINE_CAP_FLAT = 0, _                       ; Also called Butt, the line will end without any additional shape.
		$LOC_COMMENT_LINE_CAP_ROUND = 1, _                      ; The line will get a half circle as additional cap.
		$LOC_COMMENT_LINE_CAP_SQUARE = 2                        ; The line uses a square for the line end.

; Shape Line Joint Constants.
Global Const _                                                  ; com.sun.star.drawing.LineJoint
		$LOC_COMMENT_LINE_JOINT_NONE = 0, _                     ; The joint between lines will not be connected.
		$LOC_COMMENT_LINE_JOINT_MIDDLE = 1, _                   ; The middle value between the joints is used. ## Note used?
		$LOC_COMMENT_LINE_JOINT_BEVEL = 2, _                    ; The edges of the thick lines will be joined by lines.
		$LOC_COMMENT_LINE_JOINT_MITER = 3, _                    ; The lines join at intersections.
		$LOC_COMMENT_LINE_JOINT_ROUND = 4                       ; The lines join with an arc.

; Shape Line Style Constants.
Global Enum _
		$LOC_COMMENT_LINE_STYLE_NONE, _                         ; 0 -- No Line is applied.
		$LOC_COMMENT_LINE_STYLE_CONTINUOUS, _                   ; 1 -- A Solid Line.
		$LOC_COMMENT_LINE_STYLE_DOT, _                          ; 2 -- A Dotted Line.
		$LOC_COMMENT_LINE_STYLE_DOT_ROUNDED, _                  ; 3 -- A Rounded Dotted Line.
		$LOC_COMMENT_LINE_STYLE_LONG_DOT, _                     ; 4 -- A Long Dotted Line.
		$LOC_COMMENT_LINE_STYLE_LONG_DOT_ROUNDED, _             ; 5 -- A Rounded Long Dotted Line.
		$LOC_COMMENT_LINE_STYLE_DASH, _                         ; 6 -- A Dashed Line.
		$LOC_COMMENT_LINE_STYLE_DASH_ROUNDED, _                 ; 7 -- A Rounded Dashed Line.
		$LOC_COMMENT_LINE_STYLE_LONG_DASH, _                    ; 8 -- A Long Dashed Line.
		$LOC_COMMENT_LINE_STYLE_LONG_DASH_ROUNDED, _            ; 9 -- A Rounded Long Dashed Line.
		$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH, _                  ; 10 -- A Double Dashed Line.
		$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_ROUNDED, _          ; 11 -- A Rounded Double Dash.
		$LOC_COMMENT_LINE_STYLE_DASH_DOT, _                     ; 12 -- A Dashed and Dotted Line.
		$LOC_COMMENT_LINE_STYLE_DASH_DOT_ROUNDED, _             ; 13 -- A Rounded Dashed and Dotted Line.
		$LOC_COMMENT_LINE_STYLE_LONG_DASH_DOT, _                ; 14 -- A Long Dashed and Dotted Line.
		$LOC_COMMENT_LINE_STYLE_LONG_DASH_DOT_ROUNDED, _        ; 15 -- A Rounded Long Dashed and Dotted Line.
		$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_DOT, _              ; 16 -- A Double Dash Dot Line.
		$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED, _      ; 17 -- A Rounded Double Dash Dot Line
		$LOC_COMMENT_LINE_STYLE_DASH_DOT_DOT, _                 ; 18 -- A Dash Dot Dot Line.
		$LOC_COMMENT_LINE_STYLE_DASH_DOT_DOT_ROUNDED, _         ; 19 -- A Rounded Dash Dot Dot Line.
		$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_DOT_DOT, _          ; 20 -- A Double Dash Dot Dot Line.
		$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED, _  ; 21 -- A Rounded Double Dash Dot Dot Line.
		$LOC_COMMENT_LINE_STYLE_ULTRAFINE_DOTTED, _             ; 22 -- A Ultrafine Dotted Line.
		$LOC_COMMENT_LINE_STYLE_FINE_DOTTED, _                  ; 23 -- A Fine Dotted Line.
		$LOC_COMMENT_LINE_STYLE_ULTRAFINE_DASHED, _             ; 24 -- A Ultrafine Dashed Line.
		$LOC_COMMENT_LINE_STYLE_FINE_DASHED, _                  ; 25 -- A Fine Dashed Line.
		$LOC_COMMENT_LINE_STYLE_DASHED, _                       ; 26 -- A Dashed Line.
		$LOC_COMMENT_LINE_STYLE_LINE_STYLE_9, _                 ; 27 -- Line Style 9.
		$LOC_COMMENT_LINE_STYLE_3_DASHES_3_DOTS, _              ; 28 -- A Line consisting of 3 Dashes and 3 Dots.
		$LOC_COMMENT_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES, _    ; 29 -- A Ultrafine Line consisting of 2 Dots and 3 Dashes.
		$LOC_COMMENT_LINE_STYLE_2_DOTS_1_DASH, _                ; 30 -- A Line consisting of 2 Dots and 1 Dash.
		$LOC_COMMENT_LINE_STYLE_LINE_WITH_FINE_DOTS             ; 31 -- A Line with Fine Dots.

; Comment Shadow Position
Global Enum _
		$LOC_COMMENT_SHADOW_TOP_LEFT, _                         ; The comment Shadow is positioned in the Upper-Left corner of the comment box.
		$LOC_COMMENT_SHADOW_TOP_CENTER, _                       ; The comment Shadow is positioned in the Upper-Center of the comment box.
		$LOC_COMMENT_SHADOW_TOP_RIGHT, _                        ; The comment Shadow is positioned in the Upper-Right corner of the comment box.
		$LOC_COMMENT_SHADOW_MIDDLE_LEFT, _                      ; The comment Shadow is positioned in the Middle-Left corner of the comment box.
		$LOC_COMMENT_SHADOW_MIDDLE_CENTER, _                    ; The comment Shadow is positioned in the Middle-Center of the comment box.
		$LOC_COMMENT_SHADOW_MIDDLE_RIGHT, _                     ; The comment Shadow is positioned in the Middle-Right of the comment box.
		$LOC_COMMENT_SHADOW_BOTTOM_LEFT, _                      ; The comment Shadow is positioned in the Lower-Left corner of the comment box.
		$LOC_COMMENT_SHADOW_BOTTOM_CENTER, _                    ; The comment Shadow is positioned in the Lower-Center of the comment box.
		$LOC_COMMENT_SHADOW_BOTTOM_RIGHT                        ; The comment Shadow is positioned in the Lower-Right corner of the comment box.

; General Computation Functions
Global Const _                                                  ; com.sun.star.sheet.GeneralFunction
		$LOC_COMPUTE_FUNC_NONE = 0, _                           ; Nothing is calculated.
		$LOC_COMPUTE_FUNC_AUTO = 1, _                           ; Uses SUM if all values in the range are numbers, else uses COUNT.
		$LOC_COMPUTE_FUNC_SUM = 2, _                            ; Adds all numerical values in the Range.
		$LOC_COMPUTE_FUNC_COUNT = 3, _                          ; Count all cells containing a string or value.
		$LOC_COMPUTE_FUNC_AVERAGE = 4, _                        ; Average all numerical values in a range.
		$LOC_COMPUTE_FUNC_MAX = 5, _                            ; Find the maximum numerical value in the range.
		$LOC_COMPUTE_FUNC_MIN = 6, _                            ; Find the minimum numerical value in the range.
		$LOC_COMPUTE_FUNC_PRODUCT = 7, _                        ; The result of multiplying of all numbers in the range.
		$LOC_COMPUTE_FUNC_COUNTNUMS = 8, _                      ; Count the number of cells containing numerical values in the range.
		$LOC_COMPUTE_FUNC_STDEV = 9, _                          ; Standard deviation based on a sample.
		$LOC_COMPUTE_FUNC_STDEVP = 10, _                        ; Standard deviation based on the entire population.
		$LOC_COMPUTE_FUNC_VAR = 11, _                           ; Variance based on a sample.
		$LOC_COMPUTE_FUNC_VARP = 12                             ; Variance based on the entire population.

; Cursor Type Related Constants
Global Const _
		$LOC_CURTYPE_TEXT_CURSOR = 1, _                         ; Cursor is a Text Cursor type.
		$LOC_CURTYPE_SHEET_CURSOR = 2, _                        ; Cursor is a Sheet Cursor type.
		$LOC_CURTYPE_PARAGRAPH = 3, _                           ; Object is a Paragraph Object.
		$LOC_CURTYPE_TEXT_PORTION = 4                           ; Object is a Paragraph Text Portion Object.

; Printer Duplex Constants.
Global Const _                                                  ; com.sun.star.view.DuplexMode
		$LOC_DUPLEX_UNKNOWN = 0, _                              ; Duplex mode setting is unknown.
		$LOC_DUPLEX_OFF = 1, _                                  ; Duplex mode is off.
		$LOC_DUPLEX_LONG = 2, _                                 ; Duplex mode is on, flip on Long edge.
		$LOC_DUPLEX_SHORT = 3                                   ; Duplex mode is on, flip on Short edge.

; Field Types
Global Enum Step *2 _
		$LOC_FIELD_TYPE_ALL = 1, _                              ; Returns an array of all field types listed below.
		$LOC_FIELD_TYPE_DATE_TIME, _                            ; A Date or Time field. {Cell & Header.}
		$LOC_FIELD_TYPE_DOC_TITLE, _                            ; A Document Title field. {Cell & Header.}
		$LOC_FIELD_TYPE_FILE_NAME, _                            ; A File Name or Path and File Name field. {Header.}
		$LOC_FIELD_TYPE_PAGE_NUM, _                             ; A Page Number field. {Header.}
		$LOC_FIELD_TYPE_PAGE_COUNT, _                           ; A total Page Count field. {Header.}
		$LOC_FIELD_TYPE_SHEET_NAME, _                           ; A Sheet Name field. {Cell & Header.}
		$LOC_FIELD_TYPE_URL                                     ; A Hyperlink/URL field. {Cell.}

; Fill Date Mode
Global Const _                                                  ; com.sun.star.sheet.FillDateMode
		$LOC_FILL_DATE_MODE_DAY = 0, _                          ; For each Cell a single day is added.
		$LOC_FILL_DATE_MODE_WEEKDAY = 1, _                      ; For each Cell a single day is added, skipping weekends.
		$LOC_FILL_DATE_MODE_MONTH = 2, _                        ; For each Cell one month is added without modifying the day.
		$LOC_FILL_DATE_MODE_YEAR = 3                            ; For each Cell a year is added without modifying the day or month.

; Fill Direction
Global Const _                                                  ; com.sun.star.sheet.FillDirection
		$LOC_FILL_DIR_DOWN = 0, _                               ; Rows are filled from top to bottom.
		$LOC_FILL_DIR_RIGHT = 1, _                              ; Columns are filled from left to right.
		$LOC_FILL_DIR_TOP = 2, _                                ; Rows are filled from bottom to top.
		$LOC_FILL_DIR_LEFT = 3                                  ; Columns are filled from right to left.

; Fill Series Mode
Global Const _                                                  ; com.sun.star.sheet.FillMode
		$LOC_FILL_MODE_SIMPLE = 0, _                            ; All cells are filled with the same value.
		$LOC_FILL_MODE_LINEAR = 1, _                            ; The initial value is increased by a specified value, per each cell processed.
		$LOC_FILL_MODE_GROWTH = 2, _                            ; The initial value is multiplied by a specified value, per each cell processed.
		$LOC_FILL_MODE_DATE = 3, _                              ; Any date the Cells is increased by the specified number of days/
		$LOC_FILL_MODE_AUTO = 4                                 ; The cells are filled using a user-defined series.

; Filter Conditions
Global Const _                                                  ; com.sun.star.sheet.FilterOperator2
		$LOC_FILTER_CONDITION_EMPTY = 0, _                      ; Show only Empty cells.
		$LOC_FILTER_CONDITION_NOT_EMPTY = 1, _                  ; Show only non-empty cells.
		$LOC_FILTER_CONDITION_EQUAL = 2, _                      ; Show only cells equal to the value set.
		$LOC_FILTER_CONDITION_NOT_EQUAL = 3, _                  ; Show only cells NOT equal to the value set.
		$LOC_FILTER_CONDITION_GREATER = 4, _                    ; Show only cells greater than the value set.
		$LOC_FILTER_CONDITION_GREATER_EQUAL = 5, _              ; Show only cells greater than or equal to the value set.
		$LOC_FILTER_CONDITION_LESS = 6, _                       ; Show only cells less than the value set.
		$LOC_FILTER_CONDITION_LESS_EQUAL = 7, _                 ; Show only cells less than or equal to the value set.
		$LOC_FILTER_CONDITION_TOP_VALUES = 8, _                 ; Show a specified number of the largest values contained in the range.
		$LOC_FILTER_CONDITION_TOP_PERCENT = 9, _                ; Show a specified percentage of the largest values contained in the range.
		$LOC_FILTER_CONDITION_BOTTOM_VALUES = 10, _             ; Show a specified number of the lowest values contained in the range.
		$LOC_FILTER_CONDITION_BOTTOM_PERCENT = 11, _            ; Show a specified percentage of the lowest values contained in the range.
		$LOC_FILTER_CONDITION_CONTAINS = 12, _                  ; Show only cells containing the specified entry.
		$LOC_FILTER_CONDITION_DOES_NOT_CONTAIN = 13, _          ; Show only cells that do not contain the specified entry.
		$LOC_FILTER_CONDITION_BEGINS_WITH = 14, _               ; Show only cells beginning with the specified entry.
		$LOC_FILTER_CONDITION_DOES_NOT_BEGIN_WITH = 15, _       ; Show only cells not beginning with the specified entry.
		$LOC_FILTER_CONDITION_ENDS_WITH = 16, _                 ; Show only cells ending with the specified entry.
		$LOC_FILTER_CONDITION_DOES_NOT_END_WITH = 17            ; Show only cells not ending with the specified entry.

; Filter Operators
Global Const _                                                  ; com.sun.star.sheet.FilterConnection
		$LOC_FILTER_OPERATOR_AND = 0, _                         ; Both conditions have to be fulfilled.
		$LOC_FILTER_OPERATOR_OR = 1                             ; At least one of the conditions has to be fulfilled.

; Format Key Type
Global Const _                                                  ; com.sun.star.util.NumberFormat
		$LOC_FORMAT_KEYS_ALL = 0, _                             ; Returns All number formats.
		$LOC_FORMAT_KEYS_DEFINED = 1, _                         ; Returns Only user-defined number formats.
		$LOC_FORMAT_KEYS_DATE = 2, _                            ; Returns Date formats.
		$LOC_FORMAT_KEYS_TIME = 4, _                            ; Returns Time formats.
		$LOC_FORMAT_KEYS_DATE_TIME = 6, _                       ; Returns Number formats which contain date and time.
		$LOC_FORMAT_KEYS_CURRENCY = 8, _                        ; Returns Currency formats.
		$LOC_FORMAT_KEYS_NUMBER = 16, _                         ; Returns Decimal number formats.
		$LOC_FORMAT_KEYS_SCIENTIFIC = 32, _                     ; Returns Scientific number formats.
		$LOC_FORMAT_KEYS_FRACTION = 64, _                       ; Returns Number formats for fractions.
		$LOC_FORMAT_KEYS_PERCENT = 128, _                       ; Returns Percentage number formats.
		$LOC_FORMAT_KEYS_TEXT = 256, _                          ; Returns Text number formats.
		$LOC_FORMAT_KEYS_LOGICAL = 1024, _                      ; Returns Boolean number formats.
		$LOC_FORMAT_KEYS_UNDEFINED = 2048, _                    ; Returns Is used as a return value if no format exists.
		$LOC_FORMAT_KEYS_EMPTY = 4096, _                        ; Returns Empty Number formats (?)
		$LOC_FORMAT_KEYS_DURATION = 8196                        ; Returns Duration number formats.

; Formula Result Type Constants
Global Const _                                                  ; com.sun.star.sheet.FormulaResult
		$LOC_FORMULA_RESULT_TYPE_VALUE = 1, _                   ; The formula's result is a number.
		$LOC_FORMULA_RESULT_TYPE_STRING = 2, _                  ; The formula's result is a string.
		$LOC_FORMULA_RESULT_TYPE_ERROR = 4, _                   ; The formula has an error of some form.
		$LOC_FORMULA_RESULT_TYPE_ALL = 7                        ; All of the above types.

; Gradient Names
Global Const _
		$LOC_GRAD_NAME_PASTEL_BOUQUET = "Pastel Bouquet", _     ; The "Pastel Bouquet" Gradient Preset.
		$LOC_GRAD_NAME_PASTEL_DREAM = "Pastel Dream", _         ; The "Pastel Dream" Gradient Preset.
		$LOC_GRAD_NAME_BLUE_TOUCH = "Blue Touch", _             ; The "Blue Touch" Gradient Preset.
		$LOC_GRAD_NAME_BLANK_W_GRAY = "Blank with Gray", _      ; The "Blank with Gray" Gradient Preset.
		$LOC_GRAD_NAME_LONDON_MIST = "London Mist", _           ; The "London Mist" Gradient Preset.
		$LOC_GRAD_NAME_SUBMARINE = "Submarine", _               ; The "Submarine" Gradient Preset.
		$LOC_GRAD_NAME_MIDNIGHT = "Midnight", _                 ; The "Midnight" Gradient Preset.
		$LOC_GRAD_NAME_DEEP_OCEAN = "Deep Ocean", _             ; The "Deep Ocean" Gradient Preset.
		$LOC_GRAD_NAME_MAHOGANY = "Mahogany", _                 ; The "Mahogany" Gradient Preset.
		$LOC_GRAD_NAME_GREEN_GRASS = "Green Grass", _           ; The "Green Grass" Gradient Preset.
		$LOC_GRAD_NAME_NEON_LIGHT = "Neon Light", _             ; The "Neon Light" Gradient Preset.
		$LOC_GRAD_NAME_SUNSHINE = "Sunshine", _                 ; The "Sunshine" Gradient Preset.
		$LOC_GRAD_NAME_RAINBOW = "Rainbow", _                   ; The "Rainbow" Gradient Preset. L.O. 7.6+
		$LOC_GRAD_NAME_SUNRISE = "Sunrise", _                   ; The "Sunrise" Gradient Preset. L.O. 7.6+
		$LOC_GRAD_NAME_SUNDOWN = "Sundown"                      ; The "Sundown" Gradient Preset. L.O. 7.6+

; Gradient Type
Global Const _                                                  ; com.sun.star.awt.GradientStyle
		$LOC_GRAD_TYPE_OFF = -1, _                              ; Turn the Gradient off.
		$LOC_GRAD_TYPE_LINEAR = 0, _                            ; Linear type Gradient
		$LOC_GRAD_TYPE_AXIAL = 1, _                             ; Axial type Gradient
		$LOC_GRAD_TYPE_RADIAL = 2, _                            ; Radial type Gradient
		$LOC_GRAD_TYPE_ELLIPTICAL = 3, _                        ; Elliptical type Gradient
		$LOC_GRAD_TYPE_SQUARE = 4, _                            ; Square type Gradient
		$LOC_GRAD_TYPE_RECT = 5                                 ; Rectangle type Gradient

; Group Orientation
Global Const _                                                  ; com.sun.star.table.TableOrientation
		$LOC_GROUP_ORIENT_COLUMNS = 0, _                        ; Group using Columns.
		$LOC_GROUP_ORIENT_ROWS = 1                              ; Group using Rows.

; Named Range Options
Global Const _                                                  ; com.sun.star.sheet.NamedRangeFlag
		$LOC_NAMED_RANGE_OPT_NONE = 0, _                        ; Normally used for a common Named Range.
		$LOC_NAMED_RANGE_OPT_FILTER = 1, _                      ; The range contains filter criteria.
		$LOC_NAMED_RANGE_OPT_PRINT = 2, _                       ; The range can be used as a print range.
		$LOC_NAMED_RANGE_OPT_COLUMN = 4, _                      ; The range can be used as column headers for printing.
		$LOC_NAMED_RANGE_OPT_ROW = 8                            ; The range can be used as row headers for printing.

; Numbering Style Type
Global Const _                                                  ; com.sun.star.style.NumberingType
		$LOC_NUM_STYLE_CHARS_UPPER_LETTER = 0, _                ; Numbering is put in upper case letters. ("A, B, C, D)
		$LOC_NUM_STYLE_CHARS_LOWER_LETTER = 1, _                ; Numbering is in lower case letters. (a, b, c, d)
		$LOC_NUM_STYLE_ROMAN_UPPER = 2, _                       ; Numbering is in Roman numbers with upper case letters. (I, II, III)
		$LOC_NUM_STYLE_ROMAN_LOWER = 3, _                       ; Numbering is in Roman numbers with lower case letters. (i, ii, iii).
		$LOC_NUM_STYLE_ARABIC = 4, _                            ; Numbering is in Arabic numbers. (1, 2, 3, 4),
		$LOC_NUM_STYLE_NUMBER_NONE = 5, _                       ; Numbering is invisible.
		$LOC_NUM_STYLE_CHAR_SPECIAL = 6, _                      ; Use a character from a specified font.
		$LOC_NUM_STYLE_PAGE_DESCRIPTOR = 7, _                   ; Numbering is specified in the page style.
		$LOC_NUM_STYLE_BITMAP = 8, _                            ; Numbering is displayed as a bitmap graphic.
		$LOC_NUM_STYLE_CHARS_UPPER_LETTER_N = 9, _              ; Numbering is put in upper case letters. (A, B, Y, Z, AA, BB)
		$LOC_NUM_STYLE_CHARS_LOWER_LETTER_N = 10, _             ; Numbering is put in lower case letters. (a, b, y, z, aa, bb)
		$LOC_NUM_STYLE_TRANSLITERATION = 11, _                  ; A transliteration module will be used to produce numbers in Chinese, Japanese, etc.
		$LOC_NUM_STYLE_NATIVE_NUMBERING = 12, _                 ; The NativeNumberSupplier service will be called to produce numbers in native languages.
		$LOC_NUM_STYLE_FULLWIDTH_ARABIC = 13, _                 ; Numbering for full width Arabic number.
		$LOC_NUM_STYLE_CIRCLE_NUMBER = 14, _                    ; Bullet for Circle Number.
		$LOC_NUM_STYLE_NUMBER_LOWER_ZH = 15, _                  ; Numbering for Chinese lower case number.
		$LOC_NUM_STYLE_NUMBER_UPPER_ZH = 16, _                  ; Numbering for Chinese upper case number.
		$LOC_NUM_STYLE_NUMBER_UPPER_ZH_TW = 17, _               ; Numbering for Traditional Chinese upper case number.
		$LOC_NUM_STYLE_TIAN_GAN_ZH = 18, _                      ; Bullet for Chinese Tian Gan.
		$LOC_NUM_STYLE_DI_ZI_ZH = 19, _                         ; Bullet for Chinese Di Zi.
		$LOC_NUM_STYLE_NUMBER_TRADITIONAL_JA = 20, _            ; Numbering for Japanese traditional number.
		$LOC_NUM_STYLE_AIU_FULLWIDTH_JA = 21, _                 ; Bullet for Japanese AIU fullwidth.
		$LOC_NUM_STYLE_AIU_HALFWIDTH_JA = 22, _                 ; Bullet for Japanese AIU halfwidth.
		$LOC_NUM_STYLE_IROHA_FULLWIDTH_JA = 23, _               ; Bullet for Japanese IROHA fullwidth.
		$LOC_NUM_STYLE_IROHA_HALFWIDTH_JA = 24, _               ; Bullet for Japanese IROHA halfwidth.
		$LOC_NUM_STYLE_NUMBER_UPPER_KO = 25, _                  ; Numbering for Korean upper case number.
		$LOC_NUM_STYLE_NUMBER_HANGUL_KO = 26, _                 ; Numbering for Korean Hangul number.
		$LOC_NUM_STYLE_HANGUL_JAMO_KO = 27, _                   ; Bullet for Korean Hangul Jamo.
		$LOC_NUM_STYLE_HANGUL_SYLLABLE_KO = 28, _               ; Bullet for Korean Hangul Syllable.
		$LOC_NUM_STYLE_HANGUL_CIRCLED_JAMO_KO = 29, _           ; Bullet for Korean Hangul Circled Jamo.
		$LOC_NUM_STYLE_HANGUL_CIRCLED_SYLLABLE_KO = 30, _       ; Bullet for Korean Hangul Circled Syllable.
		$LOC_NUM_STYLE_CHARS_ARABIC = 31, _                     ; Numbering in Arabic alphabet letters.
		$LOC_NUM_STYLE_CHARS_THAI = 32, _                       ; Numbering in Thai alphabet letters.
		$LOC_NUM_STYLE_CHARS_HEBREW = 33, _                     ; Numbering in Hebrew alphabet letters.
		$LOC_NUM_STYLE_CHARS_NEPALI = 34, _                     ; Numbering in Nepali alphabet letters.
		$LOC_NUM_STYLE_CHARS_KHMER = 35, _                      ; Numbering in Khmer alphabet letters.
		$LOC_NUM_STYLE_CHARS_LAO = 36, _                        ; Numbering in Lao alphabet letters.
		$LOC_NUM_STYLE_CHARS_TIBETAN = 37, _                    ; Numbering in Tibetan/Dzongkha alphabet letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_BG = 38, _   ; Numbering in Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_BG = 39, _   ; Numbering in Cyrillic alphabet lower case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_BG = 40, _ ; Numbering in Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_BG = 41, _ ; Numbering in Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_RU = 42, _   ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_RU = 43, _   ; Numbering in Russian Cyrillic alphabet lower case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_RU = 44, _ ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_RU = 45, _ ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_PERSIAN = 46, _                    ; Numbering in Persian alphabet letters.
		$LOC_NUM_STYLE_CHARS_MYANMAR = 47, _                    ; Numbering in Myanmar alphabet letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_SR = 48, _   ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_SR = 49, _   ; Numbering in Russian Serbian alphabet lower case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_SR = 50, _ ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_SR = 51, _ ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_GREEK_UPPER_LETTER = 52, _         ; Numbering in Greek alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_GREEK_LOWER_LETTER = 53, _         ; Numbering in Greek alphabet lower case letters.
		$LOC_NUM_STYLE_CHARS_ARABIC_ABJAD = 54, _               ; Numbering in Arabic alphabet using abjad sequence.
		$LOC_NUM_STYLE_CHARS_PERSIAN_WORD = 55, _               ; Numbering in Persian words.
		$LOC_NUM_STYLE_NUMBER_HEBREW = 56, _                    ; Numbering in Hebrew numerals.
		$LOC_NUM_STYLE_NUMBER_ARABIC_INDIC = 57, _              ; Numbering in Arabic-Indic numerals.
		$LOC_NUM_STYLE_NUMBER_EAST_ARABIC_INDIC = 58, _         ; Numbering in East Arabic-Indic numerals.
		$LOC_NUM_STYLE_NUMBER_INDIC_DEVANAGARI = 59, _          ; Numbering in Indic Devanagari numerals.
		$LOC_NUM_STYLE_TEXT_NUMBER = 60, _                      ; Numbering in ordinal numbers of the language of the text node. (1st, 2nd, 3rd)
		$LOC_NUM_STYLE_TEXT_CARDINAL = 61, _                    ; Numbering in cardinal numbers of the language of the text node. (One, Two)
		$LOC_NUM_STYLE_TEXT_ORDINAL = 62, _                     ; Numbering in ordinal numbers of the language of the text node. (First, Second)
		$LOC_NUM_STYLE_SYMBOL_CHICAGO = 63, _                   ; Footnoting symbols according the University of Chicago style.
		$LOC_NUM_STYLE_ARABIC_ZERO = 64, _                      ; Numbering is in Arabic numbers, padded with zero to have a length of at least two. (01, 02)
		$LOC_NUM_STYLE_ARABIC_ZERO3 = 65, _                     ; Numbering is in Arabic numbers, padded with zero to have a length of at least three.
		$LOC_NUM_STYLE_ARABIC_ZERO4 = 66, _                     ; Numbering is in Arabic numbers, padded with zero to have a length of at least four.
		$LOC_NUM_STYLE_ARABIC_ZERO5 = 67, _                     ; Numbering is in Arabic numbers, padded with zero to have a length of at least five.
		$LOC_NUM_STYLE_SZEKELY_ROVAS = 68, _                    ; Numbering is in Szekely rovas (Old Hungarian) numerals.
		$LOC_NUM_STYLE_NUMBER_DIGITAL_KO = 69, _                ; Numbering is in Korean Digital number.
		$LOC_NUM_STYLE_NUMBER_DIGITAL2_KO = 70, _               ; Numbering is in Korean Digital Number, reserved "koreanDigital2".
		$LOC_NUM_STYLE_NUMBER_LEGAL_KO = 71                     ; Numbering is in Korean Legal Number, reserved "koreanLegal".

; Page Layout
Global Const _                                                  ; com.sun.star.style.PageStyleLayout
		$LOC_PAGE_LAYOUT_ALL = 0, _                             ; Page style shows both odd(Right) and even(Left) pages. With left and right margins.
		$LOC_PAGE_LAYOUT_LEFT = 1, _                            ; Page style shows only even(Left) pages. Odd pages are shown as blank pages. With left and right margins.
		$LOC_PAGE_LAYOUT_RIGHT = 2, _                           ; Page style shows only odd(Right) pages. Even pages are shown as blank pages. With left and right margins.
		$LOC_PAGE_LAYOUT_MIRRORED = 3                           ; Page style shows both odd(Right) and even(Left) pages with inner and outer margins.

; Paper Height in Hundredths of a Millimeter
Global Const _
		$LOC_PAPER_HEIGHT_A6 = 14808, _                         ; A6 paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_A5 = 21006, _                         ; A5 paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_A4 = 29693, _                         ; A4 paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_A3 = 42012, _                         ; A3 paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_B6ISO = 17602, _                      ; B6ISO paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_B5ISO = 24994, _                      ; B5ISO paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_B4ISO = 35306, _                      ; B4ISO paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_LETTER = 27940, _                     ; Letter paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_LEGAL = 35560, _                      ; Legal paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_LONG_BOND = 33020, _                  ; Long Bond paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_TABLOID = 43180, _                    ; Tabloid paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_B6JIS = 18200, _                      ; B6JIS paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_B5JIS = 25705, _                      ; B5JIS paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_B4JIS = 36398, _                      ; B4JIS paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_16KAI = 26010, _                      ; 16KAI paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_32KAI = 18390, _                      ; 32KAI paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_BIG_32KAI = 20295, _                  ; Big 32KAI paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_DLENVELOPE = 21996, _                 ; DL Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_C6ENVELOPE = 16205, _                 ; C6 Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_C6_5_ENVELOPE = 22911, _              ; C6/5 Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_C5ENVELOPE = 22911, _                 ; C5 Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_C4ENVELOPE = 32410, _                 ; C4 Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_6_3_4ENVELOPE = 16510, _              ; 6 3/4 Pound Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_7_3_4ENVELOPE = 19050, _              ; 7 3/4 Pound Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_9ENVELOPE = 22543, _                  ; 9 Pound Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_10ENVELOPE = 24130, _                 ; 10 Pound Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_11ENVELOPE = 26365, _                 ; 11 Pound Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_12ENVELOPE = 27940, _                 ; 12 Pound Envelope paper height in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_HEIGHT_JAP_POSTCARD = 14808                  ; Japanese Postcard paper height in Hundredths of a Millimeter (HMM).

; Paper Width in Hundredths of a Millimeter
Global Const _
		$LOC_PAPER_WIDTH_A6 = 10490, _                          ; A6 paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_A5 = 14808, _                          ; A5 paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_A4 = 21006, _                          ; A4 paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_A3 = 29693, _                          ; A3 paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_B6ISO = 12497, _                       ; B6ISO paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_B5ISO = 17602, _                       ; B5ISO paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_B4ISO = 24994, _                       ; B4ISO paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_LETTER = 21590, _                      ; Letter paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_LEGAL = 21590, _                       ; Legal paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_LONG_BOND = 21590, _                   ; Long Bond paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_TABLOID = 27940, _                     ; Tabloid paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_B6JIS = 12801, _                       ; B6JIS paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_B5JIS = 18212, _                       ; B5JIS paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_B4JIS = 25705, _                       ; B4JIS paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_16KAI = 18390, _                       ; 16KAI paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_32KAI = 13005, _                       ; 32KAI paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_BIG_32KAI = 13995, _                   ; Big 32KAI paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_DLENVELOPE = 10998, _                  ; DL Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_C6ENVELOPE = 11405, _                  ; C6 Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_C6_5_ENVELOPE = 11405, _               ; C6/5 Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_C5ENVELOPE = 16205, _                  ; C5 Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_C4ENVELOPE = 22911, _                  ; C4 Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_6_3_4ENVELOPE = 9208, _                ; 6 3/4 Pound Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_7_3_4ENVELOPE = 9855, _                ; 7 3/4 Pound Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_9ENVELOPE = 9843, _                    ; 9 Pound Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_10ENVELOPE = 10490, _                  ; 10 Pound Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_11ENVELOPE = 11430, _                  ; 11 Pound Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_12ENVELOPE = 12065, _                  ; 12 Pound Envelope paper width in Hundredths of a Millimeter (HMM).
		$LOC_PAPER_WIDTH_JAP_POSTCARD = 10008                   ; Japanese Postcard paper width in Hundredths of a Millimeter (HMM).

; Pivot Table Field Base Item Type
Global Const _                                                  ; com.sun.star.sheet.DataPilotFieldReferenceItemType
		$LOC_PIVOT_TBL_FIELD_BASE_ITEM_NAMED = 0, _             ; The reference item is given by a name.
		$LOC_PIVOT_TBL_FIELD_BASE_ITEM_PREV = 1, _              ; The reference item is the previous one.
		$LOC_PIVOT_TBL_FIELD_BASE_ITEM_NEXT = 2                 ; The reference item is the next one.

; Pivot Table Field Display Type
Global Const _                                                  ; com.sun.star.sheet.DataPilotFieldReferenceType
		$LOC_PIVOT_TBL_FIELD_DISP_NONE = 0, _                   ; {Normal}. The results in the data fields are displayed like they are.
		$LOC_PIVOT_TBL_FIELD_DISP_ITEM_DIFF = 1, _              ; {Difference From}. From each result, its reference value (Named, Previous or Next Item) is subtracted, and the difference is shown.
		$LOC_PIVOT_TBL_FIELD_DISP_ITEM_PERCENT = 2, _           ; {% of}. Each result is divided by its reference value.
		$LOC_PIVOT_TBL_FIELD_DISP_ITEM_PERCENT_DIFF = 3, _      ; {% Difference From}. From each result, its reference value is subtracted, and the difference divided by the reference value.
		$LOC_PIVOT_TBL_FIELD_DISP_RUNNING_TOTAL = 4, _          ; Each result is added to the sum of the results for preceding items in the base field, in the base field's sort order, and the total sum is shown.
		$LOC_PIVOT_TBL_FIELD_DISP_ROW_PERCENT = 5, _            ; {% of Row}. Each result is divided by the total result for its row in the Data Pilot table.
		$LOC_PIVOT_TBL_FIELD_DISP_COL_PERCENT = 6, _            ; {% of Column}. Same as $LOC_PIVOT_TBL_FIELD_DISP_ROW_PERCENT, but the total for the result's column is used.
		$LOC_PIVOT_TBL_FIELD_DISP_TOTAL_PERCENT = 7, _          ; {% of Total}. Same as $LOC_PIVOT_TBL_FIELD_DISP_ROW_PERCENT, but the grand total for the result's data field is used.
		$LOC_PIVOT_TBL_FIELD_DISP_INDEX = 8                     ; The row and column totals and the grand total, are used to calculate the following expression. ( original result * grand total ) / ( row total * column total )

; Data Pivot Table Field Orientation
Global Const _                                                  ; com.sun.star.sheet.DataPilotFieldOrientation
		$LOC_PIVOT_TBL_FIELD_TYPE_HIDDEN = 0, _                 ; The field is not used in the table.
		$LOC_PIVOT_TBL_FIELD_TYPE_COLUMN = 1, _                 ; The field is used as a column field.
		$LOC_PIVOT_TBL_FIELD_TYPE_ROW = 2, _                    ; The field is used as a row field.
		$LOC_PIVOT_TBL_FIELD_TYPE_FILTER = 3, _                 ; The field is used as a filter field. Also called "Page" in the constants.
		$LOC_PIVOT_TBL_FIELD_TYPE_DATA = 4                      ; The field is used as a data field.

; Posture/Italic
Global Const _                                                  ; com.sun.star.awt.FontSlant
		$LOC_POSTURE_NONE = 0, _                                ; Specifies a font without slant.
		$LOC_POSTURE_OBLIQUE = 1, _                             ; Specifies an oblique font (slant not designed into the font).
		$LOC_POSTURE_ITALIC = 2, _                              ; Specifies an italic font (slant designed into the font).
		$LOC_POSTURE_DontKnow = 3, _                            ; Specifies a font with an unknown slant. For Read Only.
		$LOC_POSTURE_REV_OBLIQUE = 4, _                         ; Specifies a reverse oblique font (slant not designed into the font).
		$LOC_POSTURE_REV_ITALIC = 5                             ; Specifies a reverse italic font (slant designed into the font).

; Relief
Global Const _                                                  ; com.sun.star.text.FontRelief
		$LOC_RELIEF_NONE = 0, _                                 ; No relief is applied.
		$LOC_RELIEF_EMBOSSED = 1, _                             ; The font relief is embossed.
		$LOC_RELIEF_ENGRAVED = 2                                ; The font relief is engraved.

; Page Print Scale Mode
Global Enum _
		$LOC_SCALE_REDUCE_ENLARGE = 1, _                        ; Specifies a scaling factor to scale all printed pages.
		$LOC_SCALE_FIT_WIDTH_HEIGHT, _                          ; Specifies the maximum number of pages horizontally (width) and vertically (height) on which every sheet is to be printed.
		$LOC_SCALE_FIT_PAGES                                    ; Specifies the maximum number of pages on which every sheet is to be printed. The scale will be reduced as necessary to fit the defined number of pages.

; Search In
Global Const _
		$LOC_SEARCH_IN_FORMULAS = 0, _                          ; Searches for the search string in formulas and in non-calculated values.
		$LOC_SEARCH_IN_VALUES = 1, _                            ; Searches for the search string in values and in formula results.
		$LOC_SEARCH_IN_COMMENTS = 2                             ; Searches for the search string in comments that are attached to the cells.

; Shadow Location
Global Const _                                                  ; com.sun.star.table.ShadowLocation
		$LOC_SHADOW_NONE = 0, _                                 ; No shadow is applied.
		$LOC_SHADOW_TOP_LEFT = 1, _                             ; Shadow is located along the upper and left sides.
		$LOC_SHADOW_TOP_RIGHT = 2, _                            ; Shadow is located along the upper and right sides.
		$LOC_SHADOW_BOTTOM_LEFT = 3, _                          ; Shadow is located along the lower and left sides.
		$LOC_SHADOW_BOTTOM_RIGHT = 4                            ; Shadow is located along the lower and right sides.

; Sheet Link Mode
Global Const _                                                  ; com.sun.star.sheet.SheetLinkMode
		$LOC_SHEET_LINK_MODE_NONE = 0, _                        ; The Sheet is not linked.
		$LOC_SHEET_LINK_MODE_NORMAL = 1, _                      ; All the Sheet's contents are copied, both values and formulas.
		$LOC_SHEET_LINK_MODE_VALUE = 2                          ; Only the Sheet's values and formula results are copied.

; Sheet Cursor Movement Constants.
Global Enum _
		$LOC_SHEETCUR_COLLAPSE_TO_CURRENT_ARRAY, _              ; Collapses or Expands the range to contain the current array formula.
		$LOC_SHEETCUR_COLLAPSE_TO_CURRENT_REGION, _             ; Collapses or Expands the range to contain all contiguous nonempty cells.
		$LOC_SHEETCUR_COLLAPSE_TO_MERGED_AREA, _                ; Collapses or Expand the range to contain merged cells that intersect the range.
		$LOC_SHEETCUR_COLLAPSE_TO_SIZE, _                       ; Beginning with the upper-left corner or the current range, set the cursor range size.
		$LOC_SHEETCUR_EXPAND_TO_ENTIRE_COLUMN, _                ; Expands the range to contain all columns that intersect the range.
		$LOC_SHEETCUR_EXPAND_TO_ENTIRE_ROW, _                   ; Expands the range to contain all rows that intersect the range.
		$LOC_SHEETCUR_GOTO_OFFSET, _                            ; Shift the cursor’s range relative to the current position. Negative numbers shift left/up; positive numbers shift right/down.
		$LOC_SHEETCUR_GOTO_START, _                             ; Move the cursor to the first filled cell at the beginning of a contiguous series of filled cells. This cell may be outside the cursor’s range.
		$LOC_SHEETCUR_GOTO_END, _                               ; Move the cursor to the last filled cell at the end of a contiguous series of filled cells. This cell may be outside the cursor’s range.
		$LOC_SHEETCUR_GOTO_NEXT, _                              ; Move the cursor to the next (right) unprotected cell.
		$LOC_SHEETCUR_GOTO_PREV, _                              ; Move the cursor to the previous (left) unprotected cell.
		$LOC_SHEETCUR_GOTO_USED_AREA_START, _                   ; Set the cursor to the start of the used area.
		$LOC_SHEETCUR_GOTO_USED_AREA_END                        ; Set the cursor to the end of the used area.

; Sort Data Type
Global Const _                                                  ; com.sun.star.table.TableSortFieldType
		$LOC_SORT_DATA_TYPE_AUTO = 0, _                         ; Automatically determine Sort Data type.
		$LOC_SORT_DATA_TYPE_NUMERIC = 1, _                      ; Sort Data type is Numerical.
		$LOC_SORT_DATA_TYPE_ALPHANUMERIC = 2                    ; Sort Data type is Text.

; Strikeout
Global Const _                                                  ; com.sun.star.awt.FontStrikeout
		$LOC_STRIKEOUT_NONE = 0, _                              ; No strike out.
		$LOC_STRIKEOUT_SINGLE = 1, _                            ; Strike out the characters with a single line.
		$LOC_STRIKEOUT_DOUBLE = 2, _                            ; Strike out the characters with a double line.
		$LOC_STRIKEOUT_DONT_KNOW = 3, _                         ; The strikeout mode is not specified. For Read Only.
		$LOC_STRIKEOUT_BOLD = 4, _                              ; Strike out the characters with a bold line.
		$LOC_STRIKEOUT_SLASH = 5, _                             ; Strike out the characters with slashes.
		$LOC_STRIKEOUT_X = 6                                    ; Strike out the characters with X's.

; Text Cursor Movement Constants.
Global Enum _
		$LOC_TEXTCUR_COLLAPSE_TO_START, _                       ; Collapses the current selection to the start of the selection.
		$LOC_TEXTCUR_COLLAPSE_TO_END, _                         ; Collapses the current selection the to end of the selection.
		$LOC_TEXTCUR_GO_LEFT, _                                 ; Move the cursor left by n characters.
		$LOC_TEXTCUR_GO_RIGHT, _                                ; Move the cursor right by n characters.
		$LOC_TEXTCUR_GOTO_START, _                              ; Move the cursor to the start of the text.
		$LOC_TEXTCUR_GOTO_END                                   ; Move the cursor to the end of the text.

; Text Direction
Global Const _                                                  ; com.sun.star.text.WritingMode2
		$LOC_TXT_DIR_LR = 0, _                                  ; Text within lines is written left-to-right. Typically, this is the writing mode for normal "alphabetic" text.
		$LOC_TXT_DIR_RL = 1, _                                  ; Text within a line are written right-to-left. Typically, this writing mode is used in Arabic and Hebrew text.
		$LOC_TXT_DIR_CONTEXT = 4                                ; Obtain actual writing mode from the context of the object.

; Underline/Overline
Global Const _                                                  ; com.sun.star.awt.FontUnderline
		$LOC_UNDERLINE_NONE = 0, _                              ; No Underline or Overline style.
		$LOC_UNDERLINE_SINGLE = 1, _                            ; Single line Underline/Overline style.
		$LOC_UNDERLINE_DOUBLE = 2, _                            ; Double line Underline/Overline style.
		$LOC_UNDERLINE_DOTTED = 3, _                            ; Dotted line Underline/Overline style.
		$LOC_UNDERLINE_DONT_KNOW = 4, _                         ; Unknown Underline/Overline style, for read only.
		$LOC_UNDERLINE_DASH = 5, _                              ; Dashed line Underline/Overline style.
		$LOC_UNDERLINE_LONG_DASH = 6, _                         ; Long Dashed line Underline/Overline style.
		$LOC_UNDERLINE_DASH_DOT = 7, _                          ; Dash Dot line Underline/Overline style.
		$LOC_UNDERLINE_DASH_DOT_DOT = 8, _                      ; Dash Dot Dot line Underline/Overline style.
		$LOC_UNDERLINE_SML_WAVE = 9, _                          ; Small Wave line Underline/Overline style.
		$LOC_UNDERLINE_WAVE = 10, _                             ; Wave line Underline/Overline style.
		$LOC_UNDERLINE_DBL_WAVE = 11, _                         ; Double Wave line Underline/Overline style.
		$LOC_UNDERLINE_BOLD = 12, _                             ; Bold line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_DOTTED = 13, _                      ; Bold Dotted line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_DASH = 14, _                        ; Bold Dashed line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_LONG_DASH = 15, _                   ; Bold Long Dash line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_DASH_DOT = 16, _                    ; Bold Dash Dot line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_DASH_DOT_DOT = 17, _                ; Bold Dash Dot Dot line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_WAVE = 18                           ; Bold Wave line Underline/Overline style.

; Validation Conditions.
Global Const _                                                  ; com.sun.star.sheet.ConditionOperator
		$LOC_VALIDATION_COND_NONE = 0, _                        ; No condition is specified.
		$LOC_VALIDATION_COND_EQUAL = 1, _                       ; The cell value is equal to the specified value.
		$LOC_VALIDATION_COND_NOT_EQUAL = 2, _                   ; The cell value must not be equal to the specified value.
		$LOC_VALIDATION_COND_GREATER = 3, _                     ; The cell value has to be greater than the specified value.
		$LOC_VALIDATION_COND_GREATER_EQUAL = 4, _               ; The cell value is greater or equal to the specified value.
		$LOC_VALIDATION_COND_LESS = 5, _                        ; The cell value has to be less than the specified value.
		$LOC_VALIDATION_COND_LESS_EQUAL = 6, _                  ; The cell value is less or equal to the specified value.
		$LOC_VALIDATION_COND_BETWEEN = 7, _                     ; The cell value has to be between the two specified values.
		$LOC_VALIDATION_COND_NOT_BETWEEN = 8, _                 ; The cell value has to be outside of the two specified values.
		$LOC_VALIDATION_COND_FORMULA = 9                        ; The specified formula has to give a non-zero result.

; Validation On Error Alert Types.
Global Const _                                                  ; com.sun.star.sheet.ValidationAlertStyle
		$LOC_VALIDATION_ERROR_ALERT_STOP = 0, _                 ; Error message is shown and the change is rejected.
		$LOC_VALIDATION_ERROR_ALERT_WARNING = 1, _              ; Warning message is shown and the user is asked whether the change will be accepted (defaulted to "No").
		$LOC_VALIDATION_ERROR_ALERT_INFO = 2, _                 ; Information message is shown and the user is asked whether the change will be accepted (defaulted to "Yes").
		$LOC_VALIDATION_ERROR_ALERT_MACRO = 3                   ; A macro is executed.

; Validation List Visibility.
Global Const _                                                  ; com.sun.star.sheet.TableValidationVisibility
		$LOC_VALIDATION_LIST_INVISIBLE = 0, _                   ; The List is not shown.
		$LOC_VALIDATION_LIST_UNSORTED = 1, _                    ; The List is shown unsorted.
		$LOC_VALIDATION_LIST_SORT_ASCENDING = 2                 ; The List is shown sorted ascending.

; Validation Types.
Global Const _                                                  ; com.sun.star.sheet.ValidationType
		$LOC_VALIDATION_TYPE_ANY = 0, _                         ; Any cell content is valid; no conditions are used.
		$LOC_VALIDATION_TYPE_WHOLE = 1, _                       ; Any whole number matching the specified condition is valid.
		$LOC_VALIDATION_TYPE_DECIMAL = 2, _                     ; Any number matching the specified condition is valid.
		$LOC_VALIDATION_TYPE_DATE = 3, _                        ; Any date value matching the specified condition is valid.
		$LOC_VALIDATION_TYPE_TIME = 4, _                        ; Any time value matching the specified condition is valid.
		$LOC_VALIDATION_TYPE_TEXT_LEN = 5, _                    ; String is valid if its length matches the specified condition.
		$LOC_VALIDATION_TYPE_LIST = 6, _                        ; Only strings from a specified list are valid.
		$LOC_VALIDATION_TYPE_CUSTOM = 7                         ; The specified formula determines which contents are valid.

; Weight/Bold
Global Const _                                                  ; com.sun.star.awt.FontWeight
		$LOC_WEIGHT_DONT_KNOW = 0, _                            ; The font weight is not specified/unknown. For Read Only.
		$LOC_WEIGHT_THIN = 50, _                                ; A 50% (Thin) font weight.
		$LOC_WEIGHT_ULTRA_LIGHT = 60, _                         ; A 60% (Ultra Light) font weight.
		$LOC_WEIGHT_LIGHT = 75, _                               ; A 75% (Light) font weight.
		$LOC_WEIGHT_SEMI_LIGHT = 90, _                          ; A 90% (Semi-Light) font weight.
		$LOC_WEIGHT_NORMAL = 100, _                             ; A 100% (Normal) font weight.
		$LOC_WEIGHT_SEMI_BOLD = 110, _                          ; A 110% (Semi-Bold) font weight.
		$LOC_WEIGHT_BOLD = 150, _                               ; A 150% (Bold) font weight.
		$LOC_WEIGHT_ULTRA_BOLD = 175, _                         ; A 175% (Ultra-Bold) font weight.
		$LOC_WEIGHT_BLACK = 200                                 ; A 200% (Black) font weight.

; Zoom Type Constants
Global Const _                                                  ; com.sun.star.view.DocumentZoomType
		$LOC_ZOOMTYPE_OPTIMAL = 0, _                            ; The page content width (excluding margins) at the current selection is fit into the view.
		$LOC_ZOOMTYPE_PAGE_WIDTH = 1, _                         ; The page width at the current selection is fit into the view.
		$LOC_ZOOMTYPE_ENTIRE_PAGE = 2, _                        ; A complete page of the document is fit into the view.
		$LOC_ZOOMTYPE_BY_VALUE = 3, _                           ; The Zoom property is relative, and set using Zoom Value.
		$LOC_ZOOMTYPE_PAGE_WIDTH_EXACT = 4                      ; The Page width at the current selection is fit into the view with the view ends exactly at the end of the page.
