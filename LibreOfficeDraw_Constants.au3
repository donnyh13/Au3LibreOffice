#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice Impress Constants for the LibreOffice UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various Impress functions in the LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Note ..........: Descriptions for some Constants are taken from the LibreOffice SDK API documentation.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LODCONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LODCONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LODCONST_SLEEP_DIV = 0

#Tidy_ILC_Pos=70

; Vertical Alignment
Global Const _                                                       ; com.sun.star.style.VerticalAlignment
		$LOD_ALIGN_VERT_TOP = 0, _                                   ; Vertically Align the object to the Top.
		$LOD_ALIGN_VERT_MIDDLE = 1, _                                ; Vertically Align the object to the Middle.
		$LOD_ALIGN_VERT_BOTTOM = 2                                   ; Vertically Align the object to the Bottom.

; Anchor Type
Global Const _                                                       ; com.sun.star.text.TextContentAnchorType
		$LOD_ANCHOR_AT_PARAGRAPH = 0, _                              ; Anchors the object to the current paragraph.
		$LOD_ANCHOR_AS_CHARACTER = 1, _                              ; Anchors the Object as character. The height of the current line is resized to match the height of the selection.
		$LOD_ANCHOR_AT_PAGE = 2, _                                   ; Anchors the Object to the current page.
		$LOD_ANCHOR_AT_FRAME = 3, _                                  ; Anchors the object to the surrounding frame.
		$LOD_ANCHOR_AT_CHARACTER = 4                                 ; Anchors the Object to a character.

; Animation Direction
Global Const _                                                       ; com.sun.star.drawing.TextAnimationDirection
		$LOD_ANIMATION_DIR_LEFT = 0, _                               ; The Animation begins at the Right and goes to the Left.
		$LOD_ANIMATION_DIR_RIGHT = 1, _                              ; The Animation begins at the Left and goes to the Right.
		$LOD_ANIMATION_DIR_UP = 2, _                                 ; The Animation begins at the Bottom and goes to the Top.
		$LOD_ANIMATION_DIR_DOWN = 3                                  ; The Animation begins at the Top and goes to the Bottom.

; Animation Kind
Global Const _                                                       ; com.sun.star.drawing.TextAnimationKind
		$LOD_ANIMATION_TYPE_NONE = 0, _                              ; No Animation is applied.
		$LOD_ANIMATION_TYPE_BLINK = 1, _                             ; The text switches its state from visible to invisible continuously.
		$LOD_ANIMATION_TYPE_SCROLL_THROUGH = 2, _                    ; The text scrolls.
		$LOD_ANIMATION_TYPE_SCROLL_ALTERNATE = 3, _                  ; The text scrolls from one side to the other and back.
		$LOD_ANIMATION_TYPE_SCROLL_IN = 4                            ; The text Scrolls from one side to the final position and stops there.

; Fill Style Type Constants
Global Enum _                                                        ; com.sun.star.drawing.FillStyle
		$LOD_AREA_FILL_STYLE_OFF, _                                  ; 0 Fill Style is off.
		$LOD_AREA_FILL_STYLE_SOLID, _                                ; 1 Fill Style is a solid color.
		$LOD_AREA_FILL_STYLE_GRADIENT, _                             ; 2 Fill Style is a gradient color.
		$LOD_AREA_FILL_STYLE_HATCH, _                                ; 3 Fill Style is a Hatch style color.
		$LOD_AREA_FILL_STYLE_BITMAP                                  ; 4 Fill Style is a Bitmap.

; Case Constants
Global Const _                                                       ; com.sun.star.style.CaseMap
		$LOD_CHAR_CASEMAP_NONE = 0, _                                ; The case of the characters is unchanged.
		$LOD_CHAR_CASEMAP_UPPER = 1, _                               ; All characters are put in upper case.
		$LOD_CHAR_CASEMAP_LOWER = 2, _                               ; All characters are put in lower case.
		$LOD_CHAR_CASEMAP_TITLE = 3, _                               ; The first character of each word is put in upper case.
		$LOD_CHAR_CASEMAP_SM_CAPS = 4                                ; All characters are put in upper case, but with a smaller font height.

; Posture/Italic
Global Const _                                                       ; com.sun.star.awt.FontSlant
		$LOD_CHAR_POSTURE_NONE = 0, _                                ; Specifies a font without slant.
		$LOD_CHAR_POSTURE_OBLIQUE = 1, _                             ; Specifies an oblique font (slant not designed into the font).
		$LOD_CHAR_POSTURE_ITALIC = 2, _                              ; Specifies an italic font (slant designed into the font).
		$LOD_CHAR_POSTURE_DONTKNOW = 3, _                            ; Specifies a font with an unknown slant. For Read Only.
		$LOD_CHAR_POSTURE_REV_OBLIQUE = 4, _                         ; Specifies a reverse oblique font (slant not designed into the font).
		$LOD_CHAR_POSTURE_REV_ITALIC = 5                             ; Specifies a reverse italic font (slant designed into the font).

; Relief
Global Const _                                                       ; com.sun.star.text.FontRelief
		$LOD_CHAR_RELIEF_NONE = 0, _                                 ; No relief is applied.
		$LOD_CHAR_RELIEF_EMBOSSED = 1, _                             ; The font relief is embossed.
		$LOD_CHAR_RELIEF_ENGRAVED = 2                                ; The font relief is engraved.

; Strikeout
Global Const _                                                       ; com.sun.star.awt.FontStrikeout
		$LOD_CHAR_STRIKEOUT_NONE = 0, _                              ; No strike out.
		$LOD_CHAR_STRIKEOUT_SINGLE = 1, _                            ; Strike out the characters with a single line.
		$LOD_CHAR_STRIKEOUT_DOUBLE = 2, _                            ; Strike out the characters with a double line.
		$LOD_CHAR_STRIKEOUT_DONT_KNOW = 3, _                         ; The strikeout mode is not specified. For Read Only.
		$LOD_CHAR_STRIKEOUT_BOLD = 4, _                              ; Strike out the characters with a bold line.
		$LOD_CHAR_STRIKEOUT_SLASH = 5, _                             ; Strike out the characters with slashes.
		$LOD_CHAR_STRIKEOUT_X = 6                                    ; Strike out the characters with X's.

; Underline/Overline
Global Const _                                                       ; com.sun.star.awt.FontUnderline
		$LOD_CHAR_UNDERLINE_NONE = 0, _                              ; No Underline or Overline style.
		$LOD_CHAR_UNDERLINE_SINGLE = 1, _                            ; Single line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_DOUBLE = 2, _                            ; Double line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_DOTTED = 3, _                            ; Dotted line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_DONT_KNOW = 4, _                         ; Unknown Underline/Overline style, for read only.
		$LOD_CHAR_UNDERLINE_DASH = 5, _                              ; Dashed line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_LONG_DASH = 6, _                         ; Long Dashed line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_DASH_DOT = 7, _                          ; Dash Dot line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_DASH_DOT_DOT = 8, _                      ; Dash Dot Dot line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_SML_WAVE = 9, _                          ; Small Wave line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_WAVE = 10, _                             ; Wave line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_DBL_WAVE = 11, _                         ; Double Wave line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_BOLD = 12, _                             ; Bold line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_BOLD_DOTTED = 13, _                      ; Bold Dotted line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_BOLD_DASH = 14, _                        ; Bold Dashed line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_BOLD_LONG_DASH = 15, _                   ; Bold Long Dash line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_BOLD_DASH_DOT = 16, _                    ; Bold Dash Dot line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_BOLD_DASH_DOT_DOT = 17, _                ; Bold Dash Dot Dot line Underline/Overline style.
		$LOD_CHAR_UNDERLINE_BOLD_WAVE = 18                           ; Bold Wave line Underline/Overline style.

; Weight/Bold
Global Const _                                                       ; com.sun.star.awt.FontWeight
		$LOD_CHAR_WEIGHT_DONT_KNOW = 0, _                            ; The font weight is not specified/unknown. For Read Only.
		$LOD_CHAR_WEIGHT_THIN = 50, _                                ; A 50% (Thin) font weight.
		$LOD_CHAR_WEIGHT_ULTRA_LIGHT = 60, _                         ; A 60% (Ultra Light) font weight.
		$LOD_CHAR_WEIGHT_LIGHT = 75, _                               ; A 75% (Light) font weight.
		$LOD_CHAR_WEIGHT_SEMI_LIGHT = 90, _                          ; A 90% (Semi-Light) font weight.
		$LOD_CHAR_WEIGHT_NORMAL = 100, _                             ; A 100% (Normal) font weight.
		$LOD_CHAR_WEIGHT_SEMI_BOLD = 110, _                          ; A 110% (Semi-Bold) font weight.
		$LOD_CHAR_WEIGHT_BOLD = 150, _                               ; A 150% (Bold) font weight.
		$LOD_CHAR_WEIGHT_ULTRA_BOLD = 175, _                         ; A 175% (Ultra-Bold) font weight.
		$LOD_CHAR_WEIGHT_BLACK = 200                                 ; A 200% (Black) font weight.

; Shape Connector type
Global Const _                                                       ; com.sun.star.drawing.ConnectorType
		$LOD_DRAWSHAPE_CONNECTOR_TYPE_STANDARD = 0, _                ; The connector is drawn with three lines, with the middle line perpendicular to the other two.
		$LOD_DRAWSHAPE_CONNECTOR_TYPE_CURVE = 1, _                   ; The connector is drawn as a curve.
		$LOD_DRAWSHAPE_CONNECTOR_TYPE_STRAIGHT = 2, _                ; The connector is drawn as a straight line.
		$LOD_DRAWSHAPE_CONNECTOR_TYPE_LINE = 3                       ; The connector is drawn with three lines.

; Dimension Line Text Horizontal Position.
Global Const _                                                       ; com.sun.star.drawing.MeasureTextHorzPos
		$LOD_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_AUTO = 0, _           ; Select the best horizontal position for the text Automatically.
		$LOD_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_LEFT = 1, _           ; The text is positioned at the Left.
		$LOD_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_CENTER = 2, _         ; The text is positioned at the Horizontal Center.
		$LOD_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_RIGHT = 3             ; The text is positioned at the Right.

; Dimension Line Text Vertical Position.
Global Const _                                                       ; com.sun.star.drawing.MeasureTextVertPos
		$LOD_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_AUTO = 0, _           ; Select the best vertical position for the text automatically
		$LOD_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_TOP = 1, _            ; The text is positioned at the Top.
		$LOD_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_BOTTOM = 3, _         ; The text is positioned at the Bottom. (2 is not used?)
		$LOD_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_MIDDLE = 4            ; The text is positioned in the Vertical Middle.

; Dimension Line Unit type.
Global Const _                                                       ; ?
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_OFF = -1, _               ; No Measurement units are used.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_AUTO = 0, _               ; Measurement units are automatically determined.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_MM = 1, _                 ; Measurement units are in Millimeters.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_CM = 2, _                 ; Measurement units are in Centimeters.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_METER = 3, _              ; Measurement units are in Meters.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_KILOMETER = 4, _          ; Measurement units are in Kilometers.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_POINT = 6, _              ; Measurement units are in Points.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_PICA = 7, _               ; Measurement units are in Picas.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_INCH = 8, _               ; Measurement units are in Inches.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_FOOT = 9, _               ; Measurement units are in Feet.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_MILES = 10, _             ; Measurement units are in Miles.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_CHAR = 14, _              ; Measurement units are in Characters.
		$LOD_DRAWSHAPE_DIMENSION_UNIT_TYPE_LINE = 15                 ; Measurement units are in Lines.

; Polygon Flags
Global Const _                                                       ; com.sun.star.drawing.PolygonFlags
		$LOD_DRAWSHAPE_POINT_TYPE_NORMAL = 0, _                      ; the point is normal, from the curve discussion view.
		$LOD_DRAWSHAPE_POINT_TYPE_SMOOTH = 1, _                      ; the point is smooth, the first derivation from the curve discussion view.
		$LOD_DRAWSHAPE_POINT_TYPE_CONTROL = 2, _                     ; the point is a control point, to control the curve from the user interface.
		$LOD_DRAWSHAPE_POINT_TYPE_SYMMETRIC = 3                      ; the point is symmetric, the second derivation from the curve discussion view.

; Drawing Shape Type Constants.
Global Enum _
		$LOD_DRAWSHAPE_TYPE_3D_CONE, _                               ; 0 -- A 3D Cone.
		$LOD_DRAWSHAPE_TYPE_3D_CUBE, _                               ; 1 -- A 3D Cube.
		$LOD_DRAWSHAPE_TYPE_3D_CYLINDER, _                           ; 2 -- A 3D Cylinder.
		$LOD_DRAWSHAPE_TYPE_3D_HALF_SPHERE, _                        ; 3 -- A 3D Half-Sphere.
		$LOD_DRAWSHAPE_TYPE_3D_PYRAMID, _                            ; 4 -- A 3D Pyramid.
		$LOD_DRAWSHAPE_TYPE_3D_SHELL, _                              ; 5 -- A 3D Shell.
		$LOD_DRAWSHAPE_TYPE_3D_SPHERE, _                             ; 6 -- A 3D Sphere.
		$LOD_DRAWSHAPE_TYPE_3D_TORUS, _                              ; 7 -- A 3D Torus.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY, _                    ; 8 -- A Four-way Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY, _            ; 9 -- A Four-way Callout Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN, _             ; 10 -- A Downward Callout Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT, _             ; 11 -- A Left hand Callout Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT, _       ; 12 -- A Left and Right Callout Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT, _            ; 13 -- A Right hand Callout Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP, _               ; 14 -- A Upward Callout Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN, _          ; 15 -- A Upward and Downward Callout Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT, _         ; 16 -- Upward and Right hand Callout Arrow. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CIRCULAR, _                 ; 17 -- A Circular Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, _             ; 18 -- A Right hand Corner Arrow. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_DOWN, _                     ; 19 -- A Downward Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT, _                     ; 20 -- A Left hand Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT, _               ; 21 -- A Left and Right Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT, _            ; 22 -- A Notched Right Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT, _                    ; 23 -- A Right hand Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, _            ; 24 -- A Right or Left Arrow. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, _                 ; 25 -- A "S"-Shaped Arrow. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, _                    ; 26 -- A Split Arrow. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT, _            ; 27 -- A Striped Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_UP, _                       ; 28 -- A Upward Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_DOWN, _                  ; 29 -- A Up and Down Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT, _                 ; 30 -- A Upward and Right hand Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, _            ; 31 -- A Upward, Right hand and Downward Arrow. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_ARROWS_CHEVRON, _                        ; 32 -- A Chevron Shape Arrow.
		$LOD_DRAWSHAPE_TYPE_ARROWS_PENTAGON, _                       ; 33 -- A Pentagon Shape Arrow.
		$LOD_DRAWSHAPE_TYPE_BASIC_ARC, _                             ; 34 -- An Arc Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_ARC_BLOCK, _                       ; 35 -- A Block Arc Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_CIRCLE, _                          ; 36 -- A Circle.
		$LOD_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, _                      ; 37 -- A Pie Circle. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT, _                  ; 38 -- A Segment Circle.
		$LOD_DRAWSHAPE_TYPE_BASIC_CROSS, _                           ; 39 -- A Cross Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_CUBE, _                            ; 40 -- A Cube Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_CYLINDER, _                        ; 41 -- A Cylinder Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_DIAMOND, _                         ; 42 -- A Diamond Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_ELLIPSE, _                         ; 43 -- An Ellipse Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_FOLDED_CORNER, _                   ; 44 -- A Paper Shape with a Folded Corner.
		$LOD_DRAWSHAPE_TYPE_BASIC_FRAME, _                           ; 45 -- A Frame Shape. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_BASIC_HEXAGON, _                         ; 46 -- A Hexagon Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_OCTAGON, _                         ; 47 -- A Octagon Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_PARALLELOGRAM, _                   ; 48 -- A Parallelogram Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_RECTANGLE, _                       ; 49 -- A Rectangle.
		$LOD_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED, _               ; 50 -- A Rectangle with rounded corners.
		$LOD_DRAWSHAPE_TYPE_BASIC_REGULAR_PENTAGON, _                ; 51 -- A regular Pentagon.
		$LOD_DRAWSHAPE_TYPE_BASIC_RING, _                            ; 52 -- A Ring Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_SQUARE, _                          ; 53 -- A Square.
		$LOD_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED, _                  ; 54 -- A Square with rounded corners.
		$LOD_DRAWSHAPE_TYPE_BASIC_TRAPEZOID, _                       ; 55 -- A Trapezoid Shape.
		$LOD_DRAWSHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES, _              ; 56 -- An Isosceles Triangle.
		$LOD_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT, _                  ; 57 -- A Right Angle Triangle.
		$LOD_DRAWSHAPE_TYPE_CALLOUT_CLOUD, _                         ; 58 -- A Cloud Shaped Callout.
		$LOD_DRAWSHAPE_TYPE_CALLOUT_LINE_1, _                        ; 59 -- A Callout with Line style #1.
		$LOD_DRAWSHAPE_TYPE_CALLOUT_LINE_2, _                        ; 60 -- A Callout with Line style #2.
		$LOD_DRAWSHAPE_TYPE_CALLOUT_LINE_3, _                        ; 61 -- A Callout with Line style #3.
		$LOD_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR, _                   ; 62 -- A Rectangular Callout.
		$LOD_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED, _           ; 63 -- A Rectangular Callout with rounded corners.
		$LOD_DRAWSHAPE_TYPE_CALLOUT_ROUND, _                         ; 64 -- A Round Callout.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR, _                             ; 65 -- A plain connector.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_ARROWS, _                      ; 66 -- A Plain connector with Arrows on both ends.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_CURVED, _                      ; 67 -- A Curved connector.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ARROWS, _               ; 68 -- A Curved connector with Arrows on both ends.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ENDS_ARROW, _           ; 69 -- A Curved connector with an Arrow on one end.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_ENDS_ARROW, _                  ; 70 -- A Plain connector with an Arrow on one end.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_LINE, _                        ; 71 -- A Line connector.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_LINE_ARROWS, _                 ; 72 -- A Line connector with Arrows on both ends.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_LINE_ENDS_ARROW, _             ; 73 -- A Line connector with an Arrow on one end.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT, _                    ; 74 -- A Straight connector.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ARROWS, _             ; 75 -- A Straight connector with Arrows on both ends.
		$LOD_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW, _         ; 76 -- A Straight connector with an Arrow on one end.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_CARD, _                        ; 77 -- A Card Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_COLLATE, _                     ; 78 -- A Collate Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR, _                   ; 79 -- A Connector Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE, _          ; 80 -- A Off-Page Connector Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_DATA, _                        ; 81 -- A Data Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_DECISION, _                    ; 82 -- A Decision Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_DELAY, _                       ; 83 -- A Delay Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE, _       ; 84 -- A Direct Access Storage Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_DISPLAY, _                     ; 85 -- A Display Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_DOCUMENT, _                    ; 86 -- A Document Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_EXTRACT, _                     ; 87 -- A Extract Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE, _            ; 88 -- A Internal Storage Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_MAGNETIC_DISC, _               ; 89 -- A Magnetic Disc Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_INPUT, _                ; 90 -- A Manual Input Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_OPERATION, _            ; 91 -- A Manual Operation Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_MERGE, _                       ; 92 -- A Merge Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_MULTIDOCUMENT, _               ; 93 -- A Multi-document Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_OR, _                          ; 94 -- A Or Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_PREPARATION, _                 ; 95 -- A Preparation Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_PROCESS, _                     ; 96 -- A Process Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE, _           ; 97 -- A Alternate Process Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED, _          ; 98 -- A Predefined Process Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_PUNCHED_TAPE, _                ; 99 -- A Punched Tape Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS, _           ; 100 -- A Sequential Access Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_SORT, _                        ; 101 -- A Sort Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_STORED_DATA, _                 ; 102 -- A Stored Data Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION, _            ; 103 -- A Summing Junction Flowchart.
		$LOD_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR, _                  ; 104 -- A Terminator Flowchart.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL, _                     ; 105 -- "Air Mail" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_ASPHALT, _                      ; 106 -- "Asphalt" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_BLUE, _                         ; 107 -- "Blue" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_BLUE_MOON, _                    ; 108 -- "Blue Moon" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_BURN, _                         ; 109 -- "Burn" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_BUSTER, _                       ; 110 -- "Buster" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_DONUTS_DONUTS, _                ; 111 -- "Donuts" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_FASTER, _                       ; 112 -- "Faster" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_FAVORITE_2, _                   ; 113 -- "Favorite 2" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_FAVORITE_16, _                  ; 114 -- "Favorite 16" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_GRAY, _                         ; 115 -- "Gray" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_GREEN_TO_BLUE, _                ; 116 -- "Green To Blue" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_GHOST, _                        ; 117 -- "Ghost" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_GOLD_WAVE, _                    ; 118 -- "Gold Wave" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_HEAVY_METAL, _                  ; 119 -- "Heavy Metal" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_HULK, _                         ; 120 -- "Hulk" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_NEWSPAPER, _                    ; 121 -- "Newspaper" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_NO_WAY, _                       ; 122 -- "No Way" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_NOTE, _                         ; 123 -- "Note" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_OPEN, _                         ; 124 -- "Open" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_OUTLINE, _                      ; 125 -- "Outline" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_OUTLINE_BLUE, _                 ; 126 -- "Outline Blue" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_PLANET, _                       ; 127 -- "Planet" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_POLICE_DRAMA, _                 ; 128 -- "Police Drama" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_POW, _                          ; 129 -- "Pow!" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_PURPLE_SOLID, _                 ; 130 -- "Purple Solid" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_RETRO, _                        ; 131 -- "Retro" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_SHADOW, _                       ; 132 -- "Shadow" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_SHADOW_BLUE, _                  ; 133 -- "Shadow Blue" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_SIMPLE, _                       ; 134 -- "Simple" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_SNOW, _                         ; 135 -- "Snow" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_STAR_WARS, _                    ; 136 -- "Star Wars" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_STONE, _                        ; 137 -- "Stone" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_STYLE, _                        ; 138 -- "Style" Fontwork.
		$LOD_DRAWSHAPE_TYPE_FONTWORK_TRICOLORE, _                    ; 139 -- "Tricolore" Fontwork.
		$LOD_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS, _                ; 140 -- A line with an Arrow on either end.
		$LOD_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_CIRCLE, _          ; 141 -- A line that starts with an Arrow, and ends with a Circle.
		$LOD_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_SQUARE, _          ; 142 -- A line that starts with an Arrow, and ends with a Square.
		$LOD_DRAWSHAPE_TYPE_LINE_ARROW_LINE_CIRCLE_ARROW, _          ; 143 -- A line that starts with a Circle, and ends with an Arrow.
		$LOD_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ENDS_ARROW, _            ; 144 -- A line that ends with an Arrow.
		$LOD_DRAWSHAPE_TYPE_LINE_ARROW_LINE_SQUARE_ARROW, _          ; 145 -- A line that starts with a Square, and ends with an Arrow.
		$LOD_DRAWSHAPE_TYPE_LINE_ARROW_LINE_STARTS_ARROW, _          ; 146 -- A line that starts with an Arrow.
		$LOD_DRAWSHAPE_TYPE_LINE_CURVE, _                            ; 147 -- A Curve.
		$LOD_DRAWSHAPE_TYPE_LINE_CURVE_FILLED, _                     ; 148 -- A Filled Curve.
		$LOD_DRAWSHAPE_TYPE_LINE_DIMENSION, _                        ; 149 -- A dimension line.
		$LOD_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE, _                    ; 150 -- A Freeform Line.
		$LOD_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED, _             ; 151 -- A Filled Freeform Line.
		$LOD_DRAWSHAPE_TYPE_LINE_LINE, _                             ; 152 -- A Line.
		$LOD_DRAWSHAPE_TYPE_LINE_LINE_45, _                          ; 153 -- A 45º line.
		$LOD_DRAWSHAPE_TYPE_LINE_POLYGON, _                          ; 154 -- A Polygon.
		$LOD_DRAWSHAPE_TYPE_LINE_POLYGON_45, _                       ; 155 -- A 45 degree Polygon.
		$LOD_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED, _                ; 156 -- A Filled 45 degree Polygon.
		$LOD_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED, _                   ; 157 -- A Filled Polygon.
		$LOD_DRAWSHAPE_TYPE_STARS_4_POINT, _                         ; 158 -- A 4 Pointed Star.
		$LOD_DRAWSHAPE_TYPE_STARS_5_POINT, _                         ; 159 -- A 5 Pointed Star.
		$LOD_DRAWSHAPE_TYPE_STARS_6_POINT, _                         ; 160 -- A 6 Pointed Star. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE, _                 ; 161 -- A Concave 6 Pointed Star. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_STARS_8_POINT, _                         ; 162 -- A 8 Pointed Star.
		$LOD_DRAWSHAPE_TYPE_STARS_12_POINT, _                        ; 163 -- A 12 Pointed Star. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_STARS_24_POINT, _                        ; 164 -- A 24 Pointed Star.
		$LOD_DRAWSHAPE_TYPE_STARS_DOORPLATE, _                       ; 165 -- A Doorplate Shape.
		$LOD_DRAWSHAPE_TYPE_STARS_EXPLOSION, _                       ; 166 -- A Explosion Shape.
		$LOD_DRAWSHAPE_TYPE_STARS_SCROLL_HORIZONTAL, _               ; 167 -- A Horizontal Scroll.
		$LOD_DRAWSHAPE_TYPE_STARS_SCROLL_VERTICAL, _                 ; 168 -- A Vertical Scroll.
		$LOD_DRAWSHAPE_TYPE_STARS_SIGNET, _                          ; 169 -- A Signet Shape. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, _                  ; 170 -- A Diamond Bevel. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, _                  ; 171 -- A Octagon Bevel. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_BEVEL_SQUARE, _                   ; 172 -- A Square Bevel.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_BRACE_DOUBLE, _                   ; 173 -- A Double Brace.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_BRACE_LEFT, _                     ; 174 -- A Left hand Brace.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_BRACE_RIGHT, _                    ; 175 -- A Right hand Brace.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_BRACKET_DOUBLE, _                 ; 176 -- A Double Bracket.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_BRACKET_LEFT, _                   ; 177 -- A Left hand Bracket.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_BRACKET_RIGHT, _                  ; 178 -- A Right hand Bracket.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_CLOUD, _                          ; 179 -- A Cloud Shape. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_FLOWER, _                         ; 180 -- A Flower Shape. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_HEART, _                          ; 181 -- A Heart Shape.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING, _                      ; 182 -- A Lightning Shape. ## Note: Lightning is visually different than the one available in L.O. Shapes U.I.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_MOON, _                           ; 183 -- A Moon Shape.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_SMILEY, _                         ; 184 -- A Smiley Shape.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_SUN, _                            ; 185 -- A Sun Shape.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_PROHIBITED, _                     ; 186 -- A Prohibited Shape.
		$LOD_DRAWSHAPE_TYPE_SYMBOL_PUZZLE                            ; 187 -- A Puzzle Piece Shape. ## Not implemented into LibreOffice SDK as of 7.3.4.2 or higher.

; Field Author Display Format
Global Const _                                                       ; com.sun.star.text.AuthorDisplayFormat
		$LOD_FIELD_AUTH_NAME_FULL = 0, _                             ; The full name of the author is displayed.
		$LOD_FIELD_AUTH_NAME_LAST = 1, _                             ; Only the last name of the author is displayed.
		$LOD_FIELD_AUTH_NAME_FIRST = 2, _                            ; Only the first name of the author is displayed.
		$LOD_FIELD_AUTH_NAME_INITIALS = 3                            ; The initials of the author are displayed.

; Field Date Display Format
Global Const _
		$LOD_FIELD_DATE_FMT_STANDARD_SHORT = 2, _                    ; Standard Short Date (e.g., 03/28/92)
		$LOD_FIELD_DATE_FMT_STANDARD_LONG = 3, _                     ; Standard Long Date (e.g., Saturday, March 28, 1992)
		$LOD_FIELD_DATE_FMT_MMDDYY = 4, _                            ; Numerical Month, Day, Two-digit year (03/28/92)
		$LOD_FIELD_DATE_FMT_MMDDYYYY = 5, _                          ; Numerical Month, Day, Four-digit year (03/28/1992)
		$LOD_FIELD_DATE_FMT_MMM_DD_YYYY = 6, _                       ; Abbreviated Month Name, Day, Year (Mar 28, 1992)
		$LOD_FIELD_DATE_FMT_MMMM_DD_YYYY = 7, _                      ; Full Month Name, Day, Year (March 28, 1992)
		$LOD_FIELD_DATE_FMT_DOW_MMM_DD_YYYY = 8, _                   ; Abbreviated Day of Week + Full Month (Sat, March 28, 1992)
		$LOD_FIELD_DATE_FMT_DOW_MMMM_DD_YYYY = 9                     ; Full Day of Week + Full Month (Saturday, March 28, 1992)

; File Name Field Display Format
Global Const _                                                       ; com.sun.star.text.FilenameDisplayFormat
		$LOD_FIELD_FILENAME_FULL_PATH = 0, _                         ; The Path and File name is displayed.
		$LOD_FIELD_FILENAME_PATH = 1, _                              ; Only the path of the file is displayed.
		$LOD_FIELD_FILENAME_NAME = 2, _                              ; Only the name of the file without the file extension is displayed.
		$LOD_FIELD_FILENAME_NAME_AND_EXT = 3                         ; The file name including the file extension is displayed.

; Field Time Display Format
Global Const _
		$LOD_FIELD_TIME_FMT_STANDARD = 2, _                          ; Standard Time format (HH:MM)
		$LOD_FIELD_TIME_FMT_24H_HM = 3, _                            ; 24-Hour: Hours and Minutes (15:24)
		$LOD_FIELD_TIME_FMT_24H_HMS = 4, _                           ; 24-Hour: Hours, Minutes, Seconds (15:24:55)
		$LOD_FIELD_TIME_FMT_24H_HMS_MS = 5, _                        ; 24-Hour: With Milliseconds (15:24:55.32)
		$LOD_FIELD_TIME_FMT_12H_HM_AMPM = 6, _                       ; 12-Hour: Hours and Minutes AM/PM (5:02 PM)
		$LOD_FIELD_TIME_FMT_12H_HMS_AMPM = 7, _                      ; 12-Hour: Hours, Minutes, Seconds AM/PM (5:02:43 PM)
		$LOD_FIELD_TIME_FMT_12H_HMS_MS_AMPM = 8                      ; 12-Hour: With Milliseconds AM/PM (5:02:43.23 PM)

; Field Types
Global Enum Step *2 _
		$LOD_FIELD_TYPE_AUTHOR = 1, _                                ; 1 An Author field.
		$LOD_FIELD_TYPE_DATE_TIME, _                                 ; 2 A Date or Time field.
		$LOD_FIELD_TYPE_FILE_NAME, _                                 ; 4 A File Name field.
		$LOD_FIELD_TYPE_SLIDE_COUNT, _                               ; 8 A total Slide Count field.
		$LOD_FIELD_TYPE_SLIDE_NUM, _                                 ; 16 A Slide Number field.
		$LOD_FIELD_TYPE_SLIDE_TITLE, _                               ; 32 A Slide Title field.
		$LOD_FIELD_TYPE_URL, _                                       ; 64 A Hyperlink/URL field.
		$LOD_FIELD_TYPE_ALL = 127                                    ; 127 Returns an array of all field types listed above.

; Gradient Names
Global Const _
		$LOD_GRAD_NAME_PASTEL_BOUQUET = "Pastel Bouquet", _          ; The "Pastel Bouquet" Gradient Preset.
		$LOD_GRAD_NAME_PASTEL_DREAM = "Pastel Dream", _              ; The "Pastel Dream" Gradient Preset.
		$LOD_GRAD_NAME_BLUE_TOUCH = "Blue Touch", _                  ; The "Blue Touch" Gradient Preset.
		$LOD_GRAD_NAME_BLANK_W_GRAY = "Blank with Gray", _           ; The "Blank with Gray" Gradient Preset.
		$LOD_GRAD_NAME_LONDON_MIST = "London Mist", _                ; The "London Mist" Gradient Preset.
		$LOD_GRAD_NAME_SUBMARINE = "Submarine", _                    ; The "Submarine" Gradient Preset.
		$LOD_GRAD_NAME_MIDNIGHT = "Midnight", _                      ; The "Midnight" Gradient Preset.
		$LOD_GRAD_NAME_DEEP_OCEAN = "Deep Ocean", _                  ; The "Deep Ocean" Gradient Preset.
		$LOD_GRAD_NAME_MAHOGANY = "Mahogany", _                      ; The "Mahogany" Gradient Preset.
		$LOD_GRAD_NAME_GREEN_GRASS = "Green Grass", _                ; The "Green Grass" Gradient Preset.
		$LOD_GRAD_NAME_NEON_LIGHT = "Neon Light", _                  ; The "Neon Light" Gradient Preset.
		$LOD_GRAD_NAME_SUNSHINE = "Sunshine", _                      ; The "Sunshine" Gradient Preset.
		$LOD_GRAD_NAME_RAINBOW = "Rainbow", _                        ; The "Rainbow" Gradient Preset. L.O. 7.6+
		$LOD_GRAD_NAME_SUNRISE = "Sunrise", _                        ; The "Sunrise" Gradient Preset. L.O. 7.6+
		$LOD_GRAD_NAME_SUNDOWN = "Sundown"                           ; The "Sundown" Gradient Preset. L.O. 7.6+

; Gradient Type
Global Const _                                                       ; com.sun.star.awt.GradientStyle
		$LOD_GRAD_TYPE_OFF = -1, _                                   ; Turn the Gradient off.
		$LOD_GRAD_TYPE_LINEAR = 0, _                                 ; Linear type Gradient
		$LOD_GRAD_TYPE_AXIAL = 1, _                                  ; Axial type Gradient
		$LOD_GRAD_TYPE_RADIAL = 2, _                                 ; Radial type Gradient
		$LOD_GRAD_TYPE_ELLIPTICAL = 3, _                             ; Elliptical type Gradient
		$LOD_GRAD_TYPE_SQUARE = 4, _                                 ; Square type Gradient
		$LOD_GRAD_TYPE_RECT = 5                                      ; Rectangle type Gradient

; Handout layout arrangements.
Global Const _
		$LOD_HANDOUT_LAYOUT_ONE_SLIDE = 22, _                        ; The Handout page will contain one slide placeholder.
		$LOD_HANDOUT_LAYOUT_TWO_SLIDES = 23, _                       ; The Handout page will contain two slide placeholders.
		$LOD_HANDOUT_LAYOUT_THREE_SLIDES = 24, _                     ; The Handout page will contain three slide placeholders.
		$LOD_HANDOUT_LAYOUT_FOUR_SLIDES = 25, _                      ; The Handout page will contain four slide placeholders.
		$LOD_HANDOUT_LAYOUT_SIX_SLIDES = 26, _                       ; The Handout page will contain six slide placeholders.
		$LOD_HANDOUT_LAYOUT_NINE_SLIDES = 31                         ; The Handout page will contain nine slide placeholders.

; Numbering Style Type
Global Const _                                                       ; com.sun.star.style.NumberingType
		$LOD_NUM_FRMT_CHARS_UPPER_LETTER = 0, _                      ; Numbering is put in upper case letters. ("A, B, C, D)
		$LOD_NUM_FRMT_CHARS_LOWER_LETTER = 1, _                      ; Numbering is in lower case letters. (a, b, c, d)
		$LOD_NUM_FRMT_ROMAN_UPPER = 2, _                             ; Numbering is in Roman numbers with upper case letters. (I, II, III)
		$LOD_NUM_FRMT_ROMAN_LOWER = 3, _                             ; Numbering is in Roman numbers with lower case letters. (i, ii, iii).
		$LOD_NUM_FRMT_ARABIC = 4, _                                  ; Numbering is in Arabic numbers. (1, 2, 3, 4),
		$LOD_NUM_FRMT_NUMBER_NONE = 5, _                             ; Numbering is invisible.
		$LOD_NUM_FRMT_CHAR_SPECIAL = 6, _                            ; Use a character from a specified font.
		$LOD_NUM_FRMT_PAGE_DESCRIPTOR = 7, _                         ; Numbering is specified in the page style.
		$LOD_NUM_FRMT_BITMAP = 8, _                                  ; Numbering is displayed as a bitmap graphic.
		$LOD_NUM_FRMT_CHARS_UPPER_LETTER_N = 9, _                    ; Numbering is put in upper case letters. (A, B, Y, Z, AA, BB)
		$LOD_NUM_FRMT_CHARS_LOWER_LETTER_N = 10, _                   ; Numbering is put in lower case letters. (a, b, y, z, aa, bb)
		$LOD_NUM_FRMT_TRANSLITERATION = 11, _                        ; A transliteration module will be used to produce numbers in Chinese, Japanese, etc.
		$LOD_NUM_FRMT_NATIVE_NUMBERING = 12, _                       ; The NativeNumberSupplier service will be called to produce numbers in native languages.
		$LOD_NUM_FRMT_FULLWIDTH_ARABIC = 13, _                       ; Numbering for full width Arabic number.
		$LOD_NUM_FRMT_CIRCLE_NUMBER = 14, _                          ; Bullet for Circle Number.
		$LOD_NUM_FRMT_NUMBER_LOWER_ZH = 15, _                        ; Numbering for Chinese lower case number.
		$LOD_NUM_FRMT_NUMBER_UPPER_ZH = 16, _                        ; Numbering for Chinese upper case number.
		$LOD_NUM_FRMT_NUMBER_UPPER_ZH_TW = 17, _                     ; Numbering for Traditional Chinese upper case number.
		$LOD_NUM_FRMT_TIAN_GAN_ZH = 18, _                            ; Bullet for Chinese Tian Gan.
		$LOD_NUM_FRMT_DI_ZI_ZH = 19, _                               ; Bullet for Chinese Di Zi.
		$LOD_NUM_FRMT_NUMBER_TRADITIONAL_JA = 20, _                  ; Numbering for Japanese traditional number.
		$LOD_NUM_FRMT_AIU_FULLWIDTH_JA = 21, _                       ; Bullet for Japanese AIU fullwidth.
		$LOD_NUM_FRMT_AIU_HALFWIDTH_JA = 22, _                       ; Bullet for Japanese AIU halfwidth.
		$LOD_NUM_FRMT_IROHA_FULLWIDTH_JA = 23, _                     ; Bullet for Japanese IROHA fullwidth.
		$LOD_NUM_FRMT_IROHA_HALFWIDTH_JA = 24, _                     ; Bullet for Japanese IROHA halfwidth.
		$LOD_NUM_FRMT_NUMBER_UPPER_KO = 25, _                        ; Numbering for Korean upper case number.
		$LOD_NUM_FRMT_NUMBER_HANGUL_KO = 26, _                       ; Numbering for Korean Hangul number.
		$LOD_NUM_FRMT_HANGUL_JAMO_KO = 27, _                         ; Bullet for Korean Hangul Jamo.
		$LOD_NUM_FRMT_HANGUL_SYLLABLE_KO = 28, _                     ; Bullet for Korean Hangul Syllable.
		$LOD_NUM_FRMT_HANGUL_CIRCLED_JAMO_KO = 29, _                 ; Bullet for Korean Hangul Circled Jamo.
		$LOD_NUM_FRMT_HANGUL_CIRCLED_SYLLABLE_KO = 30, _             ; Bullet for Korean Hangul Circled Syllable.
		$LOD_NUM_FRMT_CHARS_ARABIC = 31, _                           ; Numbering in Arabic alphabet letters.
		$LOD_NUM_FRMT_CHARS_THAI = 32, _                             ; Numbering in Thai alphabet letters.
		$LOD_NUM_FRMT_CHARS_HEBREW = 33, _                           ; Numbering in Hebrew alphabet letters.
		$LOD_NUM_FRMT_CHARS_NEPALI = 34, _                           ; Numbering in Nepali alphabet letters.
		$LOD_NUM_FRMT_CHARS_KHMER = 35, _                            ; Numbering in Khmer alphabet letters.
		$LOD_NUM_FRMT_CHARS_LAO = 36, _                              ; Numbering in Lao alphabet letters.
		$LOD_NUM_FRMT_CHARS_TIBETAN = 37, _                          ; Numbering in Tibetan/Dzongkha alphabet letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_UPPER_LETTER_BG = 38, _         ; Numbering in Cyrillic alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_LOWER_LETTER_BG = 39, _         ; Numbering in Cyrillic alphabet lower case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_UPPER_LETTER_N_BG = 40, _       ; Numbering in Cyrillic alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_LOWER_LETTER_N_BG = 41, _       ; Numbering in Cyrillic alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_UPPER_LETTER_RU = 42, _         ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_LOWER_LETTER_RU = 43, _         ; Numbering in Russian Cyrillic alphabet lower case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_UPPER_LETTER_N_RU = 44, _       ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_LOWER_LETTER_N_RU = 45, _       ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_PERSIAN = 46, _                          ; Numbering in Persian alphabet letters.
		$LOD_NUM_FRMT_CHARS_MYANMAR = 47, _                          ; Numbering in Myanmar alphabet letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_UPPER_LETTER_SR = 48, _         ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_LOWER_LETTER_SR = 49, _         ; Numbering in Russian Serbian alphabet lower case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_UPPER_LETTER_N_SR = 50, _       ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_CYRILLIC_LOWER_LETTER_N_SR = 51, _       ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_GREEK_UPPER_LETTER = 52, _               ; Numbering in Greek alphabet upper case letters.
		$LOD_NUM_FRMT_CHARS_GREEK_LOWER_LETTER = 53, _               ; Numbering in Greek alphabet lower case letters.
		$LOD_NUM_FRMT_CHARS_ARABIC_ABJAD = 54, _                     ; Numbering in Arabic alphabet using abjad sequence.
		$LOD_NUM_FRMT_CHARS_PERSIAN_WORD = 55, _                     ; Numbering in Persian words.
		$LOD_NUM_FRMT_NUMBER_HEBREW = 56, _                          ; Numbering in Hebrew numerals.
		$LOD_NUM_FRMT_NUMBER_ARABIC_INDIC = 57, _                    ; Numbering in Arabic-Indic numerals.
		$LOD_NUM_FRMT_NUMBER_EAST_ARABIC_INDIC = 58, _               ; Numbering in East Arabic-Indic numerals.
		$LOD_NUM_FRMT_NUMBER_INDIC_DEVANAGARI = 59, _                ; Numbering in Indic Devanagari numerals.
		$LOD_NUM_FRMT_TEXT_NUMBER = 60, _                            ; Numbering in ordinal numbers of the language of the text node. (1st, 2nd, 3rd)
		$LOD_NUM_FRMT_TEXT_CARDINAL = 61, _                          ; Numbering in cardinal numbers of the language of the text node. (One, Two)
		$LOD_NUM_FRMT_TEXT_ORDINAL = 62, _                           ; Numbering in ordinal numbers of the language of the text node. (First, Second)
		$LOD_NUM_FRMT_SYMBOL_CHICAGO = 63, _                         ; Footnoting symbols according the University of Chicago style.
		$LOD_NUM_FRMT_ARABIC_ZERO = 64, _                            ; Numbering is in Arabic numbers, padded with zero to have a length of at least two. (01, 02)
		$LOD_NUM_FRMT_ARABIC_ZERO3 = 65, _                           ; Numbering is in Arabic numbers, padded with zero to have a length of at least three.
		$LOD_NUM_FRMT_ARABIC_ZERO4 = 66, _                           ; Numbering is in Arabic numbers, padded with zero to have a length of at least four.
		$LOD_NUM_FRMT_ARABIC_ZERO5 = 67, _                           ; Numbering is in Arabic numbers, padded with zero to have a length of at least five.
		$LOD_NUM_FRMT_SZEKELY_ROVAS = 68, _                          ; Numbering is in Szekely rovas (Old Hungarian) numerals.
		$LOD_NUM_FRMT_NUMBER_DIGITAL_KO = 69, _                      ; Numbering is in Korean Digital number.
		$LOD_NUM_FRMT_NUMBER_DIGITAL2_KO = 70, _                     ; Numbering is in Korean Digital Number, reserved "koreanDigital2".
		$LOD_NUM_FRMT_NUMBER_LEGAL_KO = 71                           ; Numbering is in Korean Legal Number, reserved "koreanLegal".

; Horizontal Orientation
Global Const _                                                       ; com.sun.star.text.HoriOrientation
		$LOD_ORIENT_HORI_NONE = 0, _                                 ; No hard alignment is applied. Equal to "From Left" in L.O. U.I.
		$LOD_ORIENT_HORI_RIGHT = 1, _                                ; The object is aligned at the right side.
		$LOD_ORIENT_HORI_CENTER = 2, _                               ; The object is aligned at the middle.
		$LOD_ORIENT_HORI_LEFT = 3, _                                 ; The object is aligned at the left side.
		$LOD_ORIENT_HORI_FULL = 6, _                                 ; The table uses the full space (for text tables only).
		$LOD_ORIENT_HORI_LEFT_AND_WIDTH = 7                          ; The left offset and the width of the table are defined.

; Vertical Orientation
Global Const _                                                       ; com.sun.star.text.VertOrientation
		$LOD_ORIENT_VERT_NONE = 0, _                                 ; No hard alignment. The same as "From Top"/From Bottom" in L.O. U.I., the only difference is the combination setting of Vertical Relation.
		$LOD_ORIENT_VERT_TOP = 1, _                                  ; Aligned at the top.
		$LOD_ORIENT_VERT_CENTER = 2, _                               ; Aligned at the center.
		$LOD_ORIENT_VERT_BOTTOM = 3, _                               ; Aligned at the bottom.
		$LOD_ORIENT_VERT_CHAR_TOP = 4, _                             ; Aligned at the top of a character. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Top, and "To" = Character.
		$LOD_ORIENT_VERT_CHAR_CENTER = 5, _                          ; Aligned at the center of a character. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Character.
		$LOD_ORIENT_VERT_CHAR_BOTTOM = 6, _                          ; Aligned at the bottom of a character. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Character.
		$LOD_ORIENT_VERT_LINE_TOP = 7, _                             ; Aligned at the top of the line. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Top, and "To" = Row.
		$LOD_ORIENT_VERT_LINE_CENTER = 8, _                          ; Aligned at the center of the line. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Row.
		$LOD_ORIENT_VERT_LINE_BOTTOM = 9                             ; Aligned at the bottom of the line. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Row.

; Slide Page Height in Hundredths of a Millimeter
Global Const _
		$LOD_PAGE_HEIGHT_A6 = 14808, _                               ; A6 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_A5 = 21000, _                               ; A5 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_A4 = 29700, _                               ; A4 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_A3 = 42012, _                               ; A3 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_A2 = 59411, _                               ; A2 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_A1 = 84099, _                               ; A1 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_A0 = 11890, _                               ; A0 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_B6ISO = 17600, _                            ; B6ISO page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_B5ISO = 25000, _                            ; B5ISO page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_B4ISO = 35300, _                            ; B4ISO page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_LETTER = 27940, _                           ; Letter page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_LEGAL = 35560, _                            ; Legal page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_LONG_BOND = 33020, _                        ; Long Bond page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_TABLOID = 43180, _                          ; Tabloid page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_B6JIS = 18212, _                            ; B6JIS page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_B5JIS = 25705, _                            ; B5JIS page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_B4JIS = 36400, _                            ; B4JIS page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_16KAI = 26010, _                            ; 16KAI page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_32KAI = 18390, _                            ; 32KAI page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_BIG_32KAI = 20300, _                        ; Big 32KAI page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_DLENVELOPE = 22000, _                       ; DL Envelope page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_C6ENVELOPE = 16200, _                       ; C6 Envelope page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_C6_5_ENVELOPE = 22911, _                    ; C6/5 Envelope page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_C5ENVELOPE = 22911, _                       ; C5 Envelope page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_C4ENVELOPE = 32410, _                       ; C4 Envelope page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_DIA_SLIDE = 27000, _                        ; Dia Slide page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_SCREEN_4_3 = 28000, _                       ; Screen 4:3 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_SCREEN_16_9 = 28000, _                      ; Screen 16:9 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_SCREEN_16_10 = 28000, _                     ; Screen 16:10 page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_WIDESCREEN = 33866, _                       ; Widescreen page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_ON_SCREEN_SHOW_4_3 = 25400, _               ; On Screen Show (4:3) page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_ON_SCREEN_SHOW_16_9 = 25400, _              ; On Screen Show (16:9) page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_ON_SCREEN_SHOW_16_10 = 25400, _             ; On Screen Show (16:10) page height in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_HEIGHT_JAP_POSTCARD = 14800                        ; Japanese Postcard page height in Hundredths of a Millimeter (HMM).

; Slide Page Orientation Constants.
Global Const _                                                       ; com.sun.star.view.PaperOrientation
		$LOD_PAGE_ORIENT_PORTRAIT = 0, _                             ; Portrait Page Orientation.
		$LOD_PAGE_ORIENT_LANDSCAPE = 1                               ; Landscape Page Orientation.

; Current Document View Modes
Global Enum _
		$LOD_PAGE_VIEW_SLIDE = 0, _                                  ; 0 Slide viewing mode.
		$LOD_PAGE_VIEW_SLIDE_OUTLINE, _                              ; 1 Slide Outline viewing mode.
		$LOD_PAGE_VIEW_SLIDE_NOTES, _                                ; 2 Slide Notes viewing mode.
		$LOD_PAGE_VIEW_SLIDE_SORTER, _                               ; 3 Slide Sorter viewing mode.
		$LOD_PAGE_VIEW_MASTER, _                                     ; 4 Master Slide viewing mode.
		$LOD_PAGE_VIEW_MASTER_NOTES, _                               ; 5 Master Slide Notes viewing mode.
		$LOD_PAGE_VIEW_MASTER_HANDOUT                                ; 6 Master Slide Handout viewing mode.

; Slide Page Width in Hundredths of a Millimeter
Global Const _
		$LOD_PAGE_WIDTH_A6 = 10490, _                                ; A6 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_A5 = 14800, _                                ; A5 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_A4 = 21000, _                                ; A4 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_A3 = 29693, _                                ; A3 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_A2 = 42012, _                                ; A2 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_A1 = 59411, _                                ; A1 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_A0 = 84100, _                                ; A0 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_B6ISO = 12500, _                             ; B6ISO page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_B5ISO = 17600, _                             ; B5ISO page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_B4ISO = 25000, _                             ; B4ISO page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_LETTER = 21590, _                            ; Letter page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_LEGAL = 21590, _                             ; Legal page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_LONG_BOND = 21590, _                         ; Long Bond page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_TABLOID = 27940, _                           ; Tabloid page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_B6JIS = 12802, _                             ; B6JIS page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_B5JIS = 18212, _                             ; B5JIS page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_B4JIS = 25700, _                             ; B4JIS page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_16KAI = 18390, _                             ; 16KAI page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_32KAI = 13005, _                             ; 32KAI page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_BIG_32KAI = 14000, _                         ; Big 32KAI page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_DLENVELOPE = 11000, _                        ; DL Envelope page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_C6ENVELOPE = 11400, _                        ; C6 Envelope page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_C6_5_ENVELOPE = 11405, _                     ; C6/5 Envelope page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_C5ENVELOPE = 16205, _                        ; C5 Envelope page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_C4ENVELOPE = 22911, _                        ; C4 Envelope page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_DIA_SLIDE = 18009, _                         ; Dia Slide page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_SCREEN_4_3 = 21000, _                        ; Screen 4:3 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_SCREEN_16_9 = 15750, _                       ; Screen 16:9 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_SCREEN_16_10 = 17500, _                      ; Screen 16:10 page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_WIDESCREEN = 19050, _                        ; Widescreen page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_ON_SCREEN_SHOW_4_3 = 19050, _                ; On Screen Show (4:3) page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_ON_SCREEN_SHOW_16_9 = 14300, _               ; On Screen Show (16:9) page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_ON_SCREEN_SHOW_16_10 = 15875, _              ; On Screen Show (16:10) page width in Hundredths of a Millimeter (HMM).
		$LOD_PAGE_WIDTH_JAP_POSTCARD = 10000                         ; Japanese Postcard page width in Hundredths of a Millimeter (HMM).

; Paragraph Horizontal Align
Global Const _                                                       ; com.sun.star.style.ParagraphAdjust
		$LOD_PAR_ALIGN_HOR_LEFT = 0, _                               ; The Paragraph is left-aligned between the borders.
		$LOD_PAR_ALIGN_HOR_RIGHT = 1, _                              ; The Paragraph is right-aligned between the borders.
		$LOD_PAR_ALIGN_HOR_JUSTIFIED = 2, _                          ; The Paragraph is adjusted / stretched to both borders.
		$LOD_PAR_ALIGN_HOR_CENTER = 3, _                             ; The Paragraph is centered between the left and right borders.
		$LOD_PAR_ALIGN_HOR_STRETCH = 4                               ; HoriAlign 4 does nothing??

; Paragraph Vertical Align
Global Const _                                                       ; com.sun.star.text.ParagraphVertAlign
		$LOD_PAR_ALIGN_VERT_AUTO = 0, _                              ; Automatic vertical alignment mode. In automatic mode, horizontal text is aligned to the baseline. The same applies to text that is rotated 90°. Text that is rotated 270 ° is aligned to the center.
		$LOD_PAR_ALIGN_VERT_BASELINE = 1, _                          ; The text is aligned to the baseline.
		$LOD_PAR_ALIGN_VERT_TOP = 2, _                               ; The text is aligned to the top.
		$LOD_PAR_ALIGN_VERT_CENTER = 3, _                            ; The text is aligned to the center.
		$LOD_PAR_ALIGN_VERT_BOTTOM = 4                               ; The text is aligned to bottom.

; Paragraph Last Line Alignment
Global Const _
		$LOD_PAR_LAST_LINE_START = 0, _                              ; The Paragraph is aligned either to the Left border or the right, depending on the current text direction.
		$LOD_PAR_LAST_LINE_JUSTIFIED = 2, _                          ; The Paragraph is adjusted to both borders / stretched.
		$LOD_PAR_LAST_LINE_CENTER = 3                                ; The Paragraph is centered between the left and right borders.

; Line Spacing
Global Const _                                                       ; com.sun.star.style.LineSpacingMode
		$LOD_PAR_LINE_SPC_MODE_PROP = 0, _                           ; Specifies the height value as a proportional value. Min 6% Max 65,535%. (without percentage sign)
		$LOD_PAR_LINE_SPC_MODE_MIN = 1, _                            ; Specifies the height as the minimum line height. [Minimum/At least in L.O. U.I.] Min 0, Max 10008 (HMM)
		$LOD_PAR_LINE_SPC_MODE_LEADING = 2, _                        ; Specifies the height value as the distance to the previous line. Min 0, Max 10008 Hundredths of a Millimeter (HMM).
		$LOD_PAR_LINE_SPC_MODE_FIX = 3                               ; Specifies the height value as a fixed line height. Min 51, Max 10008 Hundredths of a Millimeter (HMM).

; Tab Alignment
Global Const _                                                       ; com.sun.star.style.TabAlign
		$LOD_PAR_TAB_ALIGN_LEFT = 0, _                               ; Aligns the left edge of the text to the tab stop and extends the text to the right.
		$LOD_PAR_TAB_ALIGN_CENTER = 1, _                             ; Aligns the center of the text to the tab stop.
		$LOD_PAR_TAB_ALIGN_RIGHT = 2, _                              ; Aligns the right edge of the text to the tab stop and extends the text to the left of the tab stop.
		$LOD_PAR_TAB_ALIGN_DECIMAL = 3, _                            ; Aligns the decimal separator of a number to the center of the tab stop and text to the left of the tab.
		$LOD_PAR_TAB_ALIGN_DEFAULT = 4                               ; This setting is the default setting when no TabStops are present. Setting any Tabstop to this constant will make it disappear from the TabStop list. It is therefore only listed here for property reading purposes.

; Horizontal Text Alignment
Global Const _                                                       ; com.sun.star.drawing.TextHorizontalAdjust
		$LOD_PAR_TEXT_ALIGN_HORI_LEFT = 0, _                         ; The left edge of the text is adjusted to the left edge of the shape.
		$LOD_PAR_TEXT_ALIGN_HORI_CENTER = 1, _                       ; The text is centered horizontally inside the shape.
		$LOD_PAR_TEXT_ALIGN_HORI_RIGHT = 2, _                        ; The right edge of the text is adjusted to the right edge of the shape.
		$LOD_PAR_TEXT_ALIGN_HORI_BLOCK = 3                           ; The text extends from the left to the right edge of the shape.

; Vertical Text Alignment
Global Const _                                                       ; com.sun.star.drawing.TextVerticalAdjust
		$LOD_PAR_TEXT_ALIGN_VERT_TOP = 0, _                          ; The top edge of the text is adjusted to the top edge of the shape.
		$LOD_PAR_TEXT_ALIGN_VERT_CENTER = 1, _                       ; The text is centered vertically inside the shape.
		$LOD_PAR_TEXT_ALIGN_VERT_BOTTOM = 2, _                       ; The bottom edge of the text is adjusted to the bottom edge of the shape.
		$LOD_PAR_TEXT_ALIGN_VERT_BLOCK = 3                           ; The text extends from the top to the bottom edge of the shape.

; Text Anchor Position
Global Enum _
		$LOD_PAR_TEXT_ANCHOR_TOP_LEFT, _                             ; 0 The text is positioned in the Upper-Left corner of the Shape.
		$LOD_PAR_TEXT_ANCHOR_TOP_CENTER, _                           ; 1 The text is positioned in the Upper-Center of the Shape.
		$LOD_PAR_TEXT_ANCHOR_TOP_RIGHT, _                            ; 2 The text is positioned in the Upper-Right of the Shape.
		$LOD_PAR_TEXT_ANCHOR_MIDDLE_LEFT, _                          ; 3 The text is positioned in the Middle-Left corner of the Shape.
		$LOD_PAR_TEXT_ANCHOR_MIDDLE_CENTER, _                        ; 4 The text is positioned in the Middle-Center of the Shape.
		$LOD_PAR_TEXT_ANCHOR_MIDDLE_RIGHT, _                         ; 5 The text is positioned in the Middle-Right of the Shape.
		$LOD_PAR_TEXT_ANCHOR_BOTTOM_LEFT, _                          ; 6 The text is positioned in the Lower-Left corner of the Shape.
		$LOD_PAR_TEXT_ANCHOR_BOTTOM_CENTER, _                        ; 7 The text is positioned in the Lower-Center of the Shape.
		$LOD_PAR_TEXT_ANCHOR_BOTTOM_RIGHT                            ; 8 The text is positioned in the Lower-Right of the Shape.

; Text Direction
Global Const _                                                       ; com.sun.star.text.WritingMode2
		$LOD_PAR_TXT_DIR_LR_TB = 0, _                                ; Text within lines is written left-to-right. Lines and blocks are placed top-to-bottom. Typically, this is the writing mode for normal "alphabetic" text.
		$LOD_PAR_TXT_DIR_RL_TB = 1, _                                ; Text within a line are written right-to-left. Lines and blocks are placed top-to-bottom. Typically, this writing mode is used in Arabic and Hebrew text.
		$LOD_PAR_TXT_DIR_TB_RL = 2, _                                ; Text within a line is written top-to-bottom. Lines and blocks are placed right-to-left. Typically, this writing mode is used in Chinese and Japanese text.
		$LOD_PAR_TXT_DIR_TB_LR = 3, _                                ; Text within a line is written top-to-bottom. Lines and blocks are placed left-to-right. Typically, this writing mode is used in Mongolian text.
		$LOD_PAR_TXT_DIR_CONTEXT = 4, _                              ; Obtain actual writing mode from the context of the object.
		$LOD_PAR_TXT_DIR_BT_LR = 5                                   ; Text within a line is written bottom-to-top. Lines and blocks are placed left-to-right. (LibreOffice 6.3).

; Relative to
Global Const _                                                       ; com.sun.star.text.RelOrientation
		$LOD_RELATIVE_ROW = -1, _                                    ; Position an object considering the row height.
		$LOD_RELATIVE_PARAGRAPH = 0, _                               ; The Object is placed considering the available paragraph space, including indent spacing. [Also called "Margin" or "Baseline" in L.O. UI]
		$LOD_RELATIVE_PARAGRAPH_TEXT = 1, _                          ; The Object is placed considering the available paragraph space, excluding indent spacing.
		$LOD_RELATIVE_CHARACTER = 2, _                               ; The Object is placed considering the available character space.
		$LOD_RELATIVE_PAGE_LEFT = 3, _                               ; The Object is placed considering the available space between the left page border and the left Paragraph border. [Same as Left Page Border in L.O. UI]
		$LOD_RELATIVE_PAGE_RIGHT = 4, _                              ; The Object is placed considering the available space between the Right page border and the Right Paragraph border. [Same as Right Page Border in L.O. UI]
		$LOD_RELATIVE_PARAGRAPH_LEFT = 5, _                          ; The Object is placed considering the available indent space to the left of the paragraph.
		$LOD_RELATIVE_PARAGRAPH_RIGHT = 6, _                         ; The Object is placed considering the available indent space to the right of the paragraph.
		$LOD_RELATIVE_PAGE = 7, _                                    ; The Object is placed considering the available space between the right and left, or top and bottom page borders.
		$LOD_RELATIVE_PAGE_PRINT = 8, _                              ; The Object is placed considering the available space between the right and left, or top and bottom page margins. [Same as Page Text Area in L.O. UI]
		$LOD_RELATIVE_TEXT_LINE = 9, _                               ; The Object is placed considering the height of the line.
		$LOD_RELATIVE_PAGE_PRINT_BOTTOM = 10, _                      ; The Object is placed considering the space available in the page footer(?)
		$LOD_RELATIVE_PAGE_PRINT_TOP = 11                            ; The Object is placed considering the space available in the page header(?)

; Table Border Style
Global Const _                                                       ; com.sun.star.table.BorderLineStyle
		$LOD_SHAPE_BORDER_STYLE_NONE = 0x7FFF, _                     ; No border line.
		$LOD_SHAPE_BORDER_STYLE_SOLID = 0, _                         ; Solid border line.
		$LOD_SHAPE_BORDER_STYLE_DOTTED = 1, _                        ; Dotted border line.
		$LOD_SHAPE_BORDER_STYLE_DASHED = 2, _                        ; Dashed border line.
		$LOD_SHAPE_BORDER_STYLE_DOUBLE = 3, _                        ; Double border line.
		$LOD_SHAPE_BORDER_STYLE_THINTHICK_SMALLGAP = 4, _            ; Double border line with a thin line outside and a thick line inside separated by a small gap.
		$LOD_SHAPE_BORDER_STYLE_THINTHICK_MEDIUMGAP = 5, _           ; Double border line with a thin line outside and a thick line inside separated by a medium gap.
		$LOD_SHAPE_BORDER_STYLE_THINTHICK_LARGEGAP = 6, _            ; Double border line with a thin line outside and a thick line inside separated by a large gap.
		$LOD_SHAPE_BORDER_STYLE_THICKTHIN_SMALLGAP = 7, _            ; Double border line with a thick line outside and a thin line inside separated by a small gap.
		$LOD_SHAPE_BORDER_STYLE_THICKTHIN_MEDIUMGAP = 8, _           ; Double border line with a thick line outside and a thin line inside separated by a medium gap.
		$LOD_SHAPE_BORDER_STYLE_THICKTHIN_LARGEGAP = 9, _            ; Double border line with a thick line outside and a thin line inside separated by a large gap.
		$LOD_SHAPE_BORDER_STYLE_EMBOSSED = 10, _                     ; 3D embossed border line.
		$LOD_SHAPE_BORDER_STYLE_ENGRAVED = 11, _                     ; 3D engraved border line.
		$LOD_SHAPE_BORDER_STYLE_OUTSET = 12, _                       ; Outset border line.
		$LOD_SHAPE_BORDER_STYLE_INSET = 13, _                        ; Inset border line.
		$LOD_SHAPE_BORDER_STYLE_FINE_DASHED = 14, _                  ; Finely dashed border line.
		$LOD_SHAPE_BORDER_STYLE_DOUBLE_THIN = 15, _                  ; Double border line consisting of two fixed thin lines separated by a variable gap.
		$LOD_SHAPE_BORDER_STYLE_DASH_DOT = 16, _                     ; Line consisting of a repetition of one dash and one dot.
		$LOD_SHAPE_BORDER_STYLE_DASH_DOT_DOT = 17                    ; Line consisting of a repetition of one dash and 2 dots.

; Border Width
Global Const _
		$LOD_SHAPE_BORDER_WIDTH_HAIRLINE = 2, _                      ; Hairline Border line width.
		$LOD_SHAPE_BORDER_WIDTH_VERY_THIN = 18, _                    ; Very Thin Border line width.
		$LOD_SHAPE_BORDER_WIDTH_THIN = 26, _                         ; Thin Border line width.
		$LOD_SHAPE_BORDER_WIDTH_MEDIUM = 53, _                       ; Medium Border line width.
		$LOD_SHAPE_BORDER_WIDTH_THICK = 79, _                        ; Thick Border line width.
		$LOD_SHAPE_BORDER_WIDTH_EXTRA_THICK = 159                    ; Extra Thick Border line width.

; Shape Use Slide Background color.
Global Const _
		$LOD_SHAPE_COLOR_USE_SLIDE_BACKGROUND = -2                   ; Use the Slide's background color as the Shape's background color.

; Shape Interaction Action on Click
Global Const _                                                       ; com.sun.star.presentation.ClickAction
		$LOD_SHAPE_INTERACTION_ACTION_NONE = 0, _                    ; No action is performed on click.
		$LOD_SHAPE_INTERACTION_ACTION_PREV_PAGE = 1, _               ; The presentation jumps to the previous page.
		$LOD_SHAPE_INTERACTION_ACTION_NEXT_PAGE = 2, _               ; The presentation jumps to the next page.
		$LOD_SHAPE_INTERACTION_ACTION_FIRST_PAGE = 3, _              ; The presentation continues with the first page.
		$LOD_SHAPE_INTERACTION_ACTION_LAST_PAGE = 4, _               ; The presentation continues with the last page.
		$LOD_SHAPE_INTERACTION_ACTION_GOTO_PAGE_OBJ = 5, _           ; The presentation jumps to a Slide or Object. Call $sTarget with the Slide name, or shape name to jump to.
		$LOD_SHAPE_INTERACTION_ACTION_DOCUMENT = 6, _                ; The presentation jumps to another document. Call $sTarget with the path to the Document to open.
		$LOD_SHAPE_INTERACTION_ACTION_INVISIBLE = 7, _               ; [Not used?] The object renders itself invisible after a click.
		$LOD_SHAPE_INTERACTION_ACTION_SOUND = 8, _                   ; A sound is played after a click. Call $sTarget with the path to the sound file to play.
		$LOD_SHAPE_INTERACTION_ACTION_OBJ_ACTION = 9, _              ; An OLE verb is performed on this object. Call $sTarget and $iVerb with the appropriate flags for the action to perform on the OLE Object.
		$LOD_SHAPE_INTERACTION_ACTION_VANISH = 10, _                 ; [Not used?] The object vanishes with its effect.
		$LOD_SHAPE_INTERACTION_ACTION_PROGRAM = 11, _                ; Another program is executed after a click. Call $sTarget with the path to the program to run.
		$LOD_SHAPE_INTERACTION_ACTION_MACRO = 12, _                  ; A macro is executed after the click. Call $sTarget with the appropriate Macro URL to call.
		$LOD_SHAPE_INTERACTION_ACTION_EXIT = 13                      ; The presentation is stopped after the click.

; Arrowhead Type Constants
Global Enum _
		$LOD_SHAPE_LINE_ARROW_TYPE_NONE, _                           ; 0 -- No Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_ARROW_SHORT, _                    ; 1 --Short Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CONCAVE_SHORT, _                  ; 2 -- Short Concave Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_ARROW, _                          ; 3 -- Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_TRIANGLE, _                       ; 4 -- Triangle Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CONCAVE, _                        ; 5 -- Concave Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_ARROW_LARGE, _                    ; 6 -- Large Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CIRCLE, _                         ; 7 -- Circle Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_SQUARE, _                         ; 8 -- Square Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_SQUARE_45, _                      ; 9 -- Square Arrow head rotated 45 degrees.
		$LOD_SHAPE_LINE_ARROW_TYPE_DIAMOND, _                        ; 10 -- Diamond Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE, _                    ; 11 -- Half Circle Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINES, _              ; 12 -- Dimension Lines head.
		$LOD_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW, _         ; 13 -- Dimension Line Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_DIMENSION_LINE, _                 ; 14 -- Dimension Line head.
		$LOD_SHAPE_LINE_ARROW_TYPE_LINE_SHORT, _                     ; 15 -- Short Line head.
		$LOD_SHAPE_LINE_ARROW_TYPE_LINE, _                           ; 16 -- Line head.
		$LOD_SHAPE_LINE_ARROW_TYPE_TRIANGLE_UNFILLED, _              ; 17 -- Unfilled Triangle Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_DIAMOND_UNFILLED, _               ; 18 -- Unfilled Diamond Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CIRCLE_UNFILLED, _                ; 19 -- Unfilled Circle Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, _             ; 20 -- Unfilled Square Arrow head, rotated 45 degrees.
		$LOD_SHAPE_LINE_ARROW_TYPE_SQUARE_UNFILLED, _                ; 21 -- Unfilled Square Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED, _           ; 22 -- Unfilled Half Circle Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_LEFT, _                ; 23 -- Half Arrow left Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_RIGHT, _               ; 24 -- Half Arrow right Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_REVERSED_ARROW, _                 ; 25 -- Reversed Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_DOUBLE_ARROW, _                   ; 26 -- Double Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CF_ONE, _                         ; 27 -- CF One Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CF_ONLY_ONE, _                    ; 28 -- CF Only One Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CF_MANY, _                        ; 29 -- CF Many Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CF_MANY_ONE, _                    ; 30 -- CF Many One Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CF_ZERO_ONE, _                    ; 31 -- CF Zero One Arrow head.
		$LOD_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY                      ; 32 -- CF Zero Many Arrow head.

; Shape Line End Cap Constants.
Global Const _                                                       ; com.sun.star.drawing.LineCap
		$LOD_SHAPE_LINE_CAP_FLAT = 0, _                              ; Also called Butt, the line will end without any additional shape.
		$LOD_SHAPE_LINE_CAP_ROUND = 1, _                             ; The line will get a half circle as additional cap.
		$LOD_SHAPE_LINE_CAP_SQUARE = 2                               ; The line uses a square for the line end.

; Shape Line Joint Constants.
Global Const _                                                       ; com.sun.star.drawing.LineJoint
		$LOD_SHAPE_LINE_JOINT_NONE = 0, _                            ; The joint between lines will not be connected.
		$LOD_SHAPE_LINE_JOINT_MIDDLE = 1, _                          ; The middle value between the joints is used. ## Note used?
		$LOD_SHAPE_LINE_JOINT_BEVEL = 2, _                           ; The edges of the thick lines will be joined by lines.
		$LOD_SHAPE_LINE_JOINT_MITER = 3, _                           ; The lines join at intersections.
		$LOD_SHAPE_LINE_JOINT_ROUND = 4                              ; The lines join with an arc.

; Shape Line Style Constants.
Global Enum _
		$LOD_SHAPE_LINE_STYLE_NONE, _                                ; 0 -- No Line is applied.
		$LOD_SHAPE_LINE_STYLE_CONTINUOUS, _                          ; 1 -- A Solid Line.
		$LOD_SHAPE_LINE_STYLE_DOT, _                                 ; 2 -- A Dotted Line.
		$LOD_SHAPE_LINE_STYLE_DOT_ROUNDED, _                         ; 3 -- A Rounded Dotted Line.
		$LOD_SHAPE_LINE_STYLE_LONG_DOT, _                            ; 4 -- A Long Dotted Line.
		$LOD_SHAPE_LINE_STYLE_LONG_DOT_ROUNDED, _                    ; 5 -- A Rounded Long Dotted Line.
		$LOD_SHAPE_LINE_STYLE_DASH, _                                ; 6 -- A Dashed Line.
		$LOD_SHAPE_LINE_STYLE_DASH_ROUNDED, _                        ; 7 -- A Rounded Dashed Line.
		$LOD_SHAPE_LINE_STYLE_LONG_DASH, _                           ; 8 -- A Long Dashed Line.
		$LOD_SHAPE_LINE_STYLE_LONG_DASH_ROUNDED, _                   ; 9 -- A Rounded Long Dashed Line.
		$LOD_SHAPE_LINE_STYLE_DOUBLE_DASH, _                         ; 10 -- A Double Dashed Line.
		$LOD_SHAPE_LINE_STYLE_DOUBLE_DASH_ROUNDED, _                 ; 11 -- A Rounded Double Dash.
		$LOD_SHAPE_LINE_STYLE_DASH_DOT, _                            ; 12 -- A Dashed and Dotted Line.
		$LOD_SHAPE_LINE_STYLE_DASH_DOT_ROUNDED, _                    ; 13 -- A Rounded Dashed and Dotted Line.
		$LOD_SHAPE_LINE_STYLE_LONG_DASH_DOT, _                       ; 14 -- A Long Dashed and Dotted Line.
		$LOD_SHAPE_LINE_STYLE_LONG_DASH_DOT_ROUNDED, _               ; 15 -- A Rounded Long Dashed and Dotted Line.
		$LOD_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT, _                     ; 16 -- A Double Dash Dot Line.
		$LOD_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED, _             ; 17 -- A Rounded Double Dash Dot Line
		$LOD_SHAPE_LINE_STYLE_DASH_DOT_DOT, _                        ; 18 -- A Dash Dot Dot Line.
		$LOD_SHAPE_LINE_STYLE_DASH_DOT_DOT_ROUNDED, _                ; 19 -- A Rounded Dash Dot Dot Line.
		$LOD_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT, _                 ; 20 -- A Double Dash Dot Dot Line.
		$LOD_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED, _         ; 21 -- A Rounded Double Dash Dot Dot Line.
		$LOD_SHAPE_LINE_STYLE_ULTRAFINE_DOTTED, _                    ; 22 -- A Ultrafine Dotted Line.
		$LOD_SHAPE_LINE_STYLE_FINE_DOTTED, _                         ; 23 -- A Fine Dotted Line.
		$LOD_SHAPE_LINE_STYLE_ULTRAFINE_DASHED, _                    ; 24 -- A Ultrafine Dashed Line.
		$LOD_SHAPE_LINE_STYLE_FINE_DASHED, _                         ; 25 -- A Fine Dashed Line.
		$LOD_SHAPE_LINE_STYLE_DASHED, _                              ; 26 -- A Dashed Line.
		$LOD_SHAPE_LINE_STYLE_SPARSE_DASH, _                         ; 27 -- A Sparse Dash. Before version 24.2 this was called Line Style 9.
		$LOD_SHAPE_LINE_STYLE_3_DASHES_3_DOTS, _                     ; 28 -- A Line consisting of 3 Dashes and 3 Dots.
		$LOD_SHAPE_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES, _           ; 29 -- A Ultrafine Line consisting of 2 Dots and 3 Dashes.
		$LOD_SHAPE_LINE_STYLE_2_DOTS_1_DASH, _                       ; 30 -- A Line consisting of 2 Dots and 1 Dash.
		$LOD_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS                    ; 31 -- A Line with Fine Dots.

; Shape Shadow Position
Global Enum _
		$LOD_SHAPE_SHADOW_LOCATION_TOP_LEFT, _                       ; 0 The Shadow is positioned in the Upper-Left corner of the shape.
		$LOD_SHAPE_SHADOW_LOCATION_TOP_CENTER, _                     ; 1 The Shadow is positioned in the Upper-Center of the shape.
		$LOD_SHAPE_SHADOW_LOCATION_TOP_RIGHT, _                      ; 2 The Shadow is positioned in the Upper-Right corner of the shape.
		$LOD_SHAPE_SHADOW_LOCATION_MIDDLE_LEFT, _                    ; 3 The Shadow is positioned in the Middle-Left corner of the shape.
		$LOD_SHAPE_SHADOW_LOCATION_MIDDLE_CENTER, _                  ; 4 The Shadow is positioned in the Middle-Center of the shape.
		$LOD_SHAPE_SHADOW_LOCATION_MIDDLE_RIGHT, _                   ; 5 The Shadow is positioned in the Middle-Right of the shape.
		$LOD_SHAPE_SHADOW_LOCATION_BOTTOM_LEFT, _                    ; 6 The Shadow is positioned in the Lower-Left corner of the shape.
		$LOD_SHAPE_SHADOW_LOCATION_BOTTOM_CENTER, _                  ; 7 The Shadow is positioned in the Lower-Center of the shape.
		$LOD_SHAPE_SHADOW_LOCATION_BOTTOM_RIGHT                      ; 8 The Shadow is positioned in the Lower-Right corner of the shape.

; Table Cell Type
Global Const _                                                       ; com.sun.star.table.CellContentType
		$LOD_SHAPE_TABLE_CELL_TYPE_EMPTY = 0, _                      ; Cell is empty.
		$LOD_SHAPE_TABLE_CELL_TYPE_VALUE = 1, _                      ; Cell contains a value.
		$LOD_SHAPE_TABLE_CELL_TYPE_TEXT = 2, _                       ; Cell contains text.
		$LOD_SHAPE_TABLE_CELL_TYPE_FORMULA = 3                       ; Cell contains a formula.

; Text Box type Constants.
Global Enum _
		$LOD_SHAPE_TEXTBOX_TYPE_TEXTBOX, _                           ; 0 - A Text Box, including Hyperlinks, and most Fields.
		$LOD_SHAPE_TEXTBOX_TYPE_OUTLINE, _                           ; 1 - A Slide Outline Text Box.
		$LOD_SHAPE_TEXTBOX_TYPE_SUBTITLE, _                          ; 2 - A Slide Subtitle Text Box.
		$LOD_SHAPE_TEXTBOX_TYPE_TITLE                                ; 3 - A Slide Title Text Box.

; Shape Type Constants.
Global Enum Step *2 _
		$LOD_SHAPE_TYPE_CALC = 1, _                                  ; 1 Calc sheet in an Impress document. (I have not encountered this shape yet, but it is included here for error prevention.)
		$LOD_SHAPE_TYPE_CHART, _                                     ; 2 Chart sheet in an Impress document. (I have not encountered this shape yet, but it is included here for error prevention.)
		$LOD_SHAPE_TYPE_DATETIME, _                                  ; 4 A Date/Time shape, such as is found in a Header or Footer or the Notes, Handouts or Master slides.
		$LOD_SHAPE_TYPE_DRAWING_SHAPE, _                             ; 8 - All shapes, 3D Shapes, Basic Shapes, Block Arrows, Flowcharts, Callouts, Lines, Connectors, Fontwork etc.
		$LOD_SHAPE_TYPE_FOOTER, _                                    ; 16 A Footer text shape, as is found in the footer of Notes, Handouts or Master slide.
		$LOD_SHAPE_TYPE_FORM_CONTROL, _                              ; 32 - Form Controls.
		$LOD_SHAPE_TYPE_HANDOUT, _                                   ; 64 A Handouts page shape, as found in the Master Handouts preview page.
		$LOD_SHAPE_TYPE_HEADER, _                                    ; 128 A Header text shape, as is found in the footer of Notes, Handouts or Master slide.
		$LOD_SHAPE_TYPE_IMAGE, _                                     ; 256 - An Image, Barcode or QR code.
		$LOD_SHAPE_TYPE_MEDIA, _                                     ; 512 - A Video or Audio shape.
		$LOD_SHAPE_TYPE_NOTES, _                                     ; 1024 A Notes page shape, as found in the slide and master slide notes pages.
		$LOD_SHAPE_TYPE_OLE2, _                                      ; 2048 - An OLE2 shape, such as a Chart, Formula etc.
		$LOD_SHAPE_TYPE_ORG_CHART, _                                 ; 4096 An Org Chart shape. (I have not encountered this shape yet, but it is included here for error prevention.)
		$LOD_SHAPE_TYPE_PAGE, _                                      ; 8192 A Page preview shape, as found in the notes and handouts pages.
		$LOD_SHAPE_TYPE_SLIDE_NUM, _                                 ; 16384 A slide number shape, as found in notes, handouts and master slides.
		$LOD_SHAPE_TYPE_TABLE, _                                     ; 32768 - A Table.
		$LOD_SHAPE_TYPE_TEXTBOX, _                                   ; 65536 - A Text Box, including Hyperlinks, and most Fields.
		$LOD_SHAPE_TYPE_TEXTBOX_SUBTITLE, _                          ; 131072 - A Slide Subtitle Text Box.
		$LOD_SHAPE_TYPE_TEXTBOX_TITLE, _                             ; 262144 - A Slide Title Text Box.
		$LOD_SHAPE_TYPE_TEXTBOX_OUTLINE, _                           ; 524288 - A Slide Outline Text Box.
		$LOD_SHAPE_TYPE_ALL = 1048575                                ; 1048575 All types above.

; Slide Header/Footer Date and Time Display Format
Global Const _
		$LOD_SLIDE_DT_FMT_MMDDYY = 4, _                              ; Numerical Month, Day, Two-digit year (03/28/92).
		$LOD_SLIDE_DT_FMT_MMDDYYYY = 5, _                            ; Numerical Month, Day, Four-digit year (03/28/1992).
		$LOD_SLIDE_DT_FMT_MMM_DD_YYYY = 6, _                         ; Abbreviated Month Name, Day, Year (Mar 28, 1992).
		$LOD_SLIDE_DT_FMT_MMMM_DD_YYYY = 7, _                        ; Full Month Name, Day, Year (March 28, 1992).
		$LOD_SLIDE_DT_FMT_DOW_MMM_DD_YYYY = 8, _                     ; Abbreviated Day of Week + Full Month (Sat, March 28, 1992).
		$LOD_SLIDE_DT_FMT_DOW_MMMM_DD_YYYY = 9, _                    ; Full Day of Week + Full Month (Saturday, March 28, 1992).
		$LOD_SLIDE_DT_FMT_24H_HM = 48, _                             ; 24-Hour: Hours and Minutes (15:24).
		$LOD_SLIDE_DT_FMT_MMDDYY_24H_HM = 52, _                      ; Numerical Month, Day, Two-digit year (03/28/92), 24-Hour: Hours and Minutes (15:24).
		$LOD_SLIDE_DT_FMT_24H_HMS = 64, _                            ; 24-Hour: Hours, Minutes, Seconds (15:24:55)
		$LOD_SLIDE_DT_FMT_12H_HM_AMPM = 96, _                        ; 12-Hour: Hours and Minutes AM/PM (5:02 PM).
		$LOD_SLIDE_DT_FMT_MMDDYY_12H_HM_AMPM = 100, _                ; Numerical Month, Day, Two-digit year (03/28/92), 12-Hour: Hours and Minutes AM/PM (5:02 PM).
		$LOD_SLIDE_DT_FMT_12H_HMS_AMPM = 112                         ; 12-Hour: Hours, Minutes, Seconds AM/PM (5:02:43 PM).

; Slide layout arrangements.
Global Const _
		$LOD_SLIDE_LAYOUT_TITLE = 0, _                               ; The Slide will contain a Title textbox and a Subtitle textbox.
		$LOD_SLIDE_LAYOUT_TITLE_CONTENT = 1, _                       ; The Slide will contain a Title textbox and a content textbox.
		$LOD_SLIDE_LAYOUT_TITLE_2_CONTENT = 3, _                     ; The Slide will contain a Title textbox and two content textboxes.
		$LOD_SLIDE_LAYOUT_TITLE_CONTENT_AND_2_CONTENT = 12, _        ; The Slide will contain a Title textbox and a content textbox beside the two smaller content boxes.
		$LOD_SLIDE_LAYOUT_TITLE_CONTENT_OVER_CONTENT = 14, _         ; The Slide will contain a Title textbox two content textboxes one positioned over top the other.
		$LOD_SLIDE_LAYOUT_TITLE_2_CONTENT_AND_CONTENT = 15, _        ; The Slide will contain a Title textbox with two smaller content textboxes beside a third content text box.
		$LOD_SLIDE_LAYOUT_TITLE_2_CONTENT_OVER_CONTENT = 16, _       ; The Slide will contain a Title textbox with two smaller content textboxes over top of a third content text box.
		$LOD_SLIDE_LAYOUT_TITLE_4_CONTENT = 18, _                    ; The Slide will contain a Title textbox with four smaller content textboxes.
		$LOD_SLIDE_LAYOUT_TITLE_ONLY = 19, _                         ; The Slide will contain a Title textbox only.
		$LOD_SLIDE_LAYOUT_BLANK = 20, _                              ; The Slide will contain no textbox.
		$LOD_SLIDE_LAYOUT_VERT_TITLE_TEXT_CHART = 27, _              ; The Slide will contain a Vertical Title with a vertical textbox, a horizontal textbox and chart placeholders.
		$LOD_SLIDE_LAYOUT_VERT_TITLE_VERT_TEXT = 28, _               ; The Slide will contain a Vertical Title with vertical textbox.
		$LOD_SLIDE_LAYOUT_TITLE_VERT_TEXT = 29, _                    ; The Slide will contain a Horizontal Title with vertical textbox.
		$LOD_SLIDE_LAYOUT_TITLE_2_VERT_TEXT_CLIPART = 30, _          ; The Slide will contain a Horizontal Title with two vertical textboxes and clipart placeholders.
		$LOD_SLIDE_LAYOUT_CENTERED_TEXT = 32, _                      ; The Slide will contain a content textbox with centered text.
		$LOD_SLIDE_LAYOUT_TITLE_6_CONTENT = 34                       ; The Slide will contain a Title textbox with six smaller content textboxes.

; Slide Transition Effects
Global Enum _
		$LOD_SLIDE_TRANSITION_3D_VENETIAN_VERT, _                    ; 0 The slide will transition with the 3D Venetian effect Vertically.
		$LOD_SLIDE_TRANSITION_3D_VENETIAN_HORI, _                    ; 1 The slide will transition with the 3D Venetian effect Horizontally.
		$LOD_SLIDE_TRANSITION_BARS_VERT, _                           ; 2 The slide will transition with the Bars effect Vertically.
		$LOD_SLIDE_TRANSITION_BARS_HORI, _                           ; 3 The slide will transition with the Bars effect Horizontally.
		$LOD_SLIDE_TRANSITION_BOX_OUT, _                             ; 4 The slide will transition with the Box effect expanding Out.
		$LOD_SLIDE_TRANSITION_BOX_IN, _                              ; 5 The slide will transition with the Box effect shrinking In.
		$LOD_SLIDE_TRANSITION_CHECKERS_DOWN, _                       ; 6 The slide will transition with the Checkers effect Down.
		$LOD_SLIDE_TRANSITION_CHECKERS_ACROSS, _                     ; 7 The slide will transition with the Checkers effect Across.
		$LOD_SLIDE_TRANSITION_CIRCLES, _                             ; 8 The slide will transition with the Circles effect.
		$LOD_SLIDE_TRANSITION_COMB_HORI, _                           ; 9 The slide will transition with the Comb effect Horizontally.
		$LOD_SLIDE_TRANSITION_COMB_VERT, _                           ; 10 The slide will transition with the Comb effect Vertically.
		$LOD_SLIDE_TRANSITION_COVER_TOP_TO_BOTTOM, _                 ; 11 The slide will transition with the Cover effect from Top to Bottom.
		$LOD_SLIDE_TRANSITION_COVER_RIGHT_TO_LEFT, _                 ; 12 The slide will transition with the Cover effect from Right to Left.
		$LOD_SLIDE_TRANSITION_COVER_LEFT_TO_RIGHT, _                 ; 13 The slide will transition with the Cover effect from Left to Right.
		$LOD_SLIDE_TRANSITION_COVER_BOTTOM_TO_TOP, _                 ; 14 The slide will transition with the Cover effect from Bottom to Top.
		$LOD_SLIDE_TRANSITION_COVER_TOP_RIGHT_TO_BOTTOM_LEFT, _      ; 15 The slide will transition with the Cover effect from Top Right to Bottom Left.
		$LOD_SLIDE_TRANSITION_COVER_BOTTOM_RIGHT_TO_TOP_LEFT, _      ; 16 The slide will transition with the Cover effect from Bottom Right to Top Left.
		$LOD_SLIDE_TRANSITION_COVER_TOP_LEFT_TO_BOTTOM_RIGHT, _      ; 17 The slide will transition with the Cover effect from Top Left to Bottom Right.
		$LOD_SLIDE_TRANSITION_COVER_BOTTOM_LEFT_TO_TOP_RIGHT, _      ; 18 The slide will transition with the Cover effect from Bottom Left to Top Right.
		$LOD_SLIDE_TRANSITION_CUBE_OUTSIDE, _                        ; 19 The slide will transition with the Cube effect from the Outside.
		$LOD_SLIDE_TRANSITION_CUBE_INSIDE, _                         ; 20 The slide will transition with the Cube effect from the Inside.
		$LOD_SLIDE_TRANSITION_CUT_THROUGH_BLACK, _                   ; 21 The slide will transition with the Cut effect through the Back.
		$LOD_SLIDE_TRANSITION_DIAGONAL_TOP_RIGHT_TO_BOTTOM_LEFT, _   ; 22 The slide will transition with the Diagonal effect from Top Right to Bottom Left.
		$LOD_SLIDE_TRANSITION_DIAGONAL_BOTTOM_RIGHT_TO_TOP_LEFT, _   ; 23 The slide will transition with the Diagonal effect from Bottom Right to Top Left.
		$LOD_SLIDE_TRANSITION_DIAGONAL_TOP_LEFT_TO_BOTTOM_RIGHT, _   ; 24 The slide will transition with the Diagonal effect from Top Left to Bottom Right.
		$LOD_SLIDE_TRANSITION_DIAGONAL_BOTTOM_LEFT_TO_TOP_RIGHT, _   ; 25 The slide will transition with the Diagonal effect from Bottom Left to Top Right.
		$LOD_SLIDE_TRANSITION_DISSOLVE, _                            ; 26 The slide will transition with the Dissolve effect.
		$LOD_SLIDE_TRANSITION_FADE_THROUGH_BLACK, _                  ; 27 The slide will transition with the Fade effect through Black.
		$LOD_SLIDE_TRANSITION_FADE_THROUGH_WHITE, _                  ; 28 The slide will transition with the Fade effect through White.
		$LOD_SLIDE_TRANSITION_FADE_SMOOTHLY, _                       ; 29 The slide will transition with the Fade effect Smoothly.
		$LOD_SLIDE_TRANSITION_FALL, _                                ; 30 The slide will transition with the Fall effect.
		$LOD_SLIDE_TRANSITION_FINE_DISSOLVE, _                       ; 31 The slide will transition with the Fine Dissolve effect.
		$LOD_SLIDE_TRANSITION_GLITTER, _                             ; 32 The slide will transition with the Glitter effect.
		$LOD_SLIDE_TRANSITION_HELIX, _                               ; 33 The slide will transition with the Helix effect.
		$LOD_SLIDE_TRANSITION_HONEYCOMB, _                           ; 34 The slide will transition with the Honeycomb effect.
		$LOD_SLIDE_TRANSITION_IRIS, _                                ; 35 The slide will transition with the Iris effect.
		$LOD_SLIDE_TRANSITION_NEWSFLASH, _                           ; 36 The slide will transition with the Newsflash effect.
		$LOD_SLIDE_TRANSITION_NONE, _                                ; 37 The slide will transition with no effect.
		$LOD_SLIDE_TRANSITION_PUSH_TOP_TO_BOTTOM, _                  ; 38 The slide will transition with the Push effect from Top to Bottom.
		$LOD_SLIDE_TRANSITION_PUSH_RIGHT_TO_LEFT, _                  ; 39 The slide will transition with the Push effect from Right to Left.
		$LOD_SLIDE_TRANSITION_PUSH_LEFT_TO_RIGHT, _                  ; 40 The slide will transition with the Push effect from Left to Right.
		$LOD_SLIDE_TRANSITION_PUSH_BOTTOM_TO_TOP, _                  ; 41 The slide will transition with the Push effect from Bottom to Top.
		$LOD_SLIDE_TRANSITION_RANDOM, _                              ; 42 The slide will transition with a Random effect.
		$LOD_SLIDE_TRANSITION_RIPPLE, _                              ; 43 The slide will transition with the Ripple effect.
		$LOD_SLIDE_TRANSITION_ROCHADE, _                             ; 44 The slide will transition with the Rochade effect.
		$LOD_SLIDE_TRANSITION_SHAPE_PLUS, _                          ; 45 The slide will transition with the Plus Shape effect.
		$LOD_SLIDE_TRANSITION_SHAPE_DIAMOND, _                       ; 46 The slide will transition with the Diamond Shape effect.
		$LOD_SLIDE_TRANSITION_SHAPE_CIRCLE, _                        ; 47 The slide will transition with the Circle Shape effect.
		$LOD_SLIDE_TRANSITION_SHAPE_OVAL_HORI, _                     ; 48 The slide will transition with the Horizontal Oval Shape effect.
		$LOD_SLIDE_TRANSITION_SHAPE_OVAL_VERT, _                     ; 49 The slide will transition with the Vertical Oval Shape effect.
		$LOD_SLIDE_TRANSITION_SPLIT_HORI_IN, _                       ; 50 The slide will transition with the Split effect Horizontally Inward.
		$LOD_SLIDE_TRANSITION_SPLIT_HORI_OUT, _                      ; 51 The slide will transition with the Split effect Horizontally Outward.
		$LOD_SLIDE_TRANSITION_SPLIT_VERT_IN, _                       ; 52 The slide will transition with the Split effect Vertically Inward.
		$LOD_SLIDE_TRANSITION_SPLIT_VERT_OUT, _                      ; 53 The slide will transition with the Split effect Vertically Outward.
		$LOD_SLIDE_TRANSITION_STATIC, _                              ; 54 The slide will transition with the Static effect.
		$LOD_SLIDE_TRANSITION_TILES, _                               ; 55 The slide will transition with the Tiles effect.
		$LOD_SLIDE_TRANSITION_TURN_AROUND, _                         ; 56 The slide will transition with the Turn Around effect.
		$LOD_SLIDE_TRANSITION_TURN_DOWN, _                           ; 57 The slide will transition with the Turn Downward effect.
		$LOD_SLIDE_TRANSITION_UNCOVER_TOP_TO_BOTTOM, _               ; 58 The slide will transition with the Uncover effect from Top to Bottom.
		$LOD_SLIDE_TRANSITION_UNCOVER_RIGHT_TO_LEFT, _               ; 59 The slide will transition with the Uncover effect from Right to Left.
		$LOD_SLIDE_TRANSITION_UNCOVER_LEFT_TO_RIGHT, _               ; 60 The slide will transition with the Uncover effect from Left to Right.
		$LOD_SLIDE_TRANSITION_UNCOVER_BOTTOM_TO_TOP, _               ; 61 The slide will transition with the Uncover effect from Bottom to Top.
		$LOD_SLIDE_TRANSITION_UNCOVER_TOP_RIGHT_TO_BOTTOM_LEFT, _    ; 62 The slide will transition with the Uncover effect from Top Right to Bottom Left.
		$LOD_SLIDE_TRANSITION_UNCOVER_BOTTOM_RIGHT_TO_TOP_LEFT, _    ; 63 The slide will transition with the Uncover effect from Bottom Right to Top Left.
		$LOD_SLIDE_TRANSITION_UNCOVER_TOP_LEFT_TO_BOTTOM_RIGHT, _    ; 64 The slide will transition with the Uncover effect from Top Left to Bottom Right.
		$LOD_SLIDE_TRANSITION_UNCOVER_BOTTOM_LEFT_TO_TOP_RIGHT, _    ; 65 The slide will transition with the Uncover effect from Bottom Left to Top Right.
		$LOD_SLIDE_TRANSITION_VENETIAN_VERT, _                       ; 66 The slide will transition with the Venetian effect Vertically.
		$LOD_SLIDE_TRANSITION_VENETIAN_HORI, _                       ; 67 The slide will transition with the Venetian effect Horizontally.
		$LOD_SLIDE_TRANSITION_VORTEX, _                              ; 68 The slide will transition with the Vortex effect.
		$LOD_SLIDE_TRANSITION_WEDGE, _                               ; 69 The slide will transition with the Wedge effect.
		$LOD_SLIDE_TRANSITION_WHEEL_1_SPOKE, _                       ; 70 The slide will transition with the Wheel effect with One Spoke.
		$LOD_SLIDE_TRANSITION_WHEEL_2_SPOKE, _                       ; 71 The slide will transition with the Wheel effect with Two Spokes.
		$LOD_SLIDE_TRANSITION_WHEEL_3_SPOKE, _                       ; 72 The slide will transition with the Wheel effect with Three Spokes.
		$LOD_SLIDE_TRANSITION_WHEEL_4_SPOKE, _                       ; 73 The slide will transition with the Wheel effect with Four Spokes.
		$LOD_SLIDE_TRANSITION_WHEEL_8_SPOKE, _                       ; 74 The slide will transition with the Wheel effect with Eight Spokes.
		$LOD_SLIDE_TRANSITION_WIPE_BOTTOM_TO_TOP, _                  ; 75 The slide will transition with the Wipe effect from Bottom to Top.
		$LOD_SLIDE_TRANSITION_WIPE_LEFT_TO_RIGHT, _                  ; 76 The slide will transition with the Wipe effect from Left to Right.
		$LOD_SLIDE_TRANSITION_WIPE_RIGHT_TO_LEFT, _                  ; 77 The slide will transition with the Wipe effect from Right to Left.
		$LOD_SLIDE_TRANSITION_WIPE_TOP_TO_BOTTOM                     ; 78 The slide will transition with the Wipe effect from Top to Bottom.

; Slideshow Presentation Mode.
Global Enum _
		$LOD_SLIDESHOW_VIEW_MODE_FULL_SCREEN, _                      ; 0 The Slideshow is Full Screen.
		$LOD_SLIDESHOW_VIEW_MODE_IN_WINDOW, _                        ; 1 The Slideshow is displayed in the LibreOffice program window.
		$LOD_SLIDESHOW_VIEW_MODE_LOOP                                ; 2 The Slideshow is looped after a set pause.

; Slideshow Pen Width
Global Const _
		$LOD_SLIDESHOW_PEN_WIDTH_VERY_THIN = 4, _                    ; A very thin width pen line for drawing with.
		$LOD_SLIDESHOW_PEN_WIDTH_THIN = 100, _                       ; A thin width pen line for drawing with.
		$LOD_SLIDESHOW_PEN_WIDTH_NORMAL = 150, _                     ; A normal width pen line for drawing with.
		$LOD_SLIDESHOW_PEN_WIDTH_THICK = 200, _                      ; A thick width pen line for drawing with.
		$LOD_SLIDESHOW_PEN_WIDTH_VERY_THICK = 400                    ; A very thick width pen line for drawing with.

; Slideshow active Presentation commands and queries.
Global Enum _
		$LOD_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE, _               ; 0 Returns the Object for the slide that is currently displayed.
		$LOD_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE_INDEX, _         ; 1 Returns the index of the current slide. Index is 0 based.
		$LOD_SLIDESHOW_PRES_QUERY_GET_NEXT_SLIDE_INDEX, _            ; 2 Returns the index for the slide that is displayed next. Index is 0 based.
		$LOD_SLIDESHOW_PRES_QUERY_GET_SLIDE_BY_INDEX, _              ; 3 Returns the Object for the slide at the index. Index is 0 based. Slides are in the order they will be displayed in the presentation which can be different than the orders of slides in the document. Not all slides must be present and each slide can be used more than once.
		$LOD_SLIDESHOW_PRES_QUERY_GET_SLIDE_COUNT, _                 ; 4 Returns the number of slides in this slide show.
		$LOD_SLIDESHOW_PRES_QUERY_IS_ACTIVE, _                       ; 5 Determines if the slide show is active. Returns TRUE for UI active slide show, FALSE otherwise.
		$LOD_SLIDESHOW_PRES_QUERY_IS_ENDLESS, _                      ; 6 Returns TRUE if the slide show was started to run endlessly.
		$LOD_SLIDESHOW_PRES_QUERY_IS_FULLSCREEN, _                   ; 7 Returns TRUE if the slide show was started in full-screen mode.
		$LOD_SLIDESHOW_PRES_QUERY_IS_PAUSED, _                       ; 8 Returns TRUE if the slide show is currently paused.
		$LOD_SLIDESHOW_PRES_COMMAND_ACTIVATE, _                      ; 9 Activates the user interface of this slide show.
		$LOD_SLIDESHOW_PRES_COMMAND_ACTIVATE_BLANK_SCREEN, _         ; 10 >Expects Parameter: Pause Screen Color as a RGB Color Integer.< Pauses the slide show and blanks the screen in the given color. Call Resume to unpause the slide show.
		$LOD_SLIDESHOW_PRES_COMMAND_DEACTIVATE, _                    ; 11 Can be called to deactivate the user interface of this slide show. (Doesn't seem to set IsActive to False!)
		$LOD_SLIDESHOW_PRES_COMMAND_ERASE_ALL_INK, _                 ; 12 Clears ink drawing from the slideshow being played. L.O. 7.2+
		$LOD_SLIDESHOW_PRES_COMMAND_GOTO_FIRST_SLIDE, _              ; 13 Goto and display the first slide.
		$LOD_SLIDESHOW_PRES_COMMAND_GOTO_LAST_SLIDE, _               ; 14 Goto and display last slide. Remaining effects on the current slide will be skipped.
		$LOD_SLIDESHOW_PRES_COMMAND_GOTO_NEXT_EFFECT, _              ; 15 Start next effects that wait on a generic trigger. If no generic triggers are waiting the next slide will be displayed.
		$LOD_SLIDESHOW_PRES_COMMAND_GOTO_NEXT_SLIDE, _               ; 16 Goto and display next slide. Remaining effects on the current slide will be skipped.
		$LOD_SLIDESHOW_PRES_COMMAND_GOTO_PREV_EFFECT, _              ; 17 Undo the last effects that were triggered by a generic trigger. If there is no previous effect that can be undone then the previous slide will be displayed.
		$LOD_SLIDESHOW_PRES_COMMAND_GOTO_PREV_SLIDE, _               ; 18 Goto and display previous slide. Remaining effects on the current slide will be skipped.
		$LOD_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE, _                    ; 19 >Expects Parameter: Slide Object to jump to.< Jumps to the given slide. The slide can also be a slide that would normally not be shown during the current slide show.
		$LOD_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE_BY_INDEX, _           ; 20 >Expects Parameter: Slide's index to jump to.< Jumps to the slide at the given index. 0 based.
		$LOD_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE_BY_NAME, _            ; 21 >Expects Parameter: Slide's name to jump to.< Jumps to the slide with the given name.
		$LOD_SLIDESHOW_PRES_COMMAND_PAUSE, _                         ; 22 Pauses the slide show. All effects are paused. The slide show continues on next user input or if resume is called.
		$LOD_SLIDESHOW_PRES_COMMAND_RESUME, _                        ; 23 Resumes a paused slide show.
		$LOD_SLIDESHOW_PRES_COMMAND_STOP_SOUND                       ; 24 Stop all currently played sounds

; Slideshow Presentation Range
Global Enum _
		$LOD_SLIDESHOW_RANGE_ALL, _                                  ; 0 All the slides in the presentation are included in the Slideshow.
		$LOD_SLIDESHOW_RANGE_FROM, _                                 ; 1 The Slideshow begins at the defined slide.
		$LOD_SLIDESHOW_RANGE_CUSTOM                                  ; 2 A custom Slideshow order is followed.

; Text Cursor Movement Constants.
Global Enum _
		$LOD_TEXTCUR_COLLAPSE_TO_START, _                            ; 0 Collapses the current selection to the start of the selection.
		$LOD_TEXTCUR_COLLAPSE_TO_END, _                              ; 1 Collapses the current selection the to end of the selection.
		$LOD_TEXTCUR_GO_LEFT, _                                      ; 2 Move the cursor left by n characters.
		$LOD_TEXTCUR_GO_RIGHT, _                                     ; 3 Move the cursor right by n characters.
		$LOD_TEXTCUR_GOTO_START, _                                   ; 4 Move the cursor to the start of the text.
		$LOD_TEXTCUR_GOTO_END                                        ; 5 Move the cursor to the end of the text.

; Zoom Type Constants
Global Const _                                                       ; com.sun.star.view.DocumentZoomType
		$LOD_ZOOMTYPE_OPTIMAL = 0, _                                 ; The page content width (excluding margins) at the current selection is fit into the view.
		$LOD_ZOOMTYPE_PAGE_WIDTH = 1, _                              ; The page width at the current selection is fit into the view.
		$LOD_ZOOMTYPE_ENTIRE_PAGE = 2, _                             ; A complete page of the document is fit into the view.
		$LOD_ZOOMTYPE_BY_VALUE = 3, _                                ; The Zoom property is relative, and set using Zoom Value.
		$LOD_ZOOMTYPE_PAGE_WIDTH_EXACT = 4                           ; The Page width at the current selection is fit into the view with the view ends exactly at the end of the page.
